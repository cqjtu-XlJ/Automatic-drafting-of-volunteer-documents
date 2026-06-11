# -*- coding: utf-8 -*-
"""
志愿活动材料自动撰写脚本
功能：根据xlsx数据源自动生成志愿服务活动时长申报表、志愿者名单汇总表、志愿服务活动总结
"""

import os
import sys
import re
from datetime import datetime, timedelta
from collections import defaultdict
from copy import deepcopy

import openpyxl
from docx import Document
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml


# ==================== 配置 ====================

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(SCRIPT_DIR, 'output')


# ==================== 数据读取 ====================

class VolunteerData:
    """志愿者数据类"""
    def __init__(self, name, student_id):
        self.name = name
        self.student_id = student_id
        self.services = []  # [(service_content, service_time, duration_hours)]
        self.total_hours = 0.0

    def add_service(self, service_content, service_time, duration_hours):
        self.services.append((service_content, service_time, duration_hours))

    def calculate_total(self):
        """计算总时长（所有服务时长的简单求和）"""
        self.total_hours = sum(s[2] for s in self.services)
        return self.total_hours

    def calculate_total_dedup(self):
        """计算去重后的总时长（同一时间段只算一次）"""
        # 解析所有时间段
        time_slots = []
        for content, time_str, duration in self.services:
            slots = parse_time_slots(time_str)
            if slots:
                time_slots.extend(slots)
            else:
                # 如果无法解析时间段，直接加时长
                time_slots.append(duration)

        # 如果全是数值（无法解析时间），直接求和
        if all(isinstance(s, (int, float)) for s in time_slots):
            self.total_hours = sum(time_slots)
            return self.total_hours

        # 合并重叠的时间段
        merged = merge_time_slots([s for s in time_slots if isinstance(s, tuple)])
        self.total_hours = sum((end - start).total_seconds() / 3600 for start, end in merged)
        return round(self.total_hours, 1)


class ActivityData:
    """活动数据类"""
    def __init__(self):
        self.volunteers = {}  # {student_id: VolunteerData}
        self.service_types = set()
        self.all_services = []  # [(name, student_id, content, time, duration)]

    def add_record(self, name, student_id, service_content, service_time, duration_hours):
        if student_id not in self.volunteers:
            self.volunteers[student_id] = VolunteerData(name, student_id)
        self.volunteers[student_id].add_service(service_content, service_time, duration_hours)
        self.service_types.add(service_content)
        self.all_services.append((name, student_id, service_content, service_time, duration_hours))

    def get_sorted_volunteers(self):
        """按姓名排序返回志愿者列表"""
        return sorted(self.volunteers.values(), key=lambda v: v.name)

    def get_total_volunteers(self):
        return len(self.volunteers)

    def get_total_hours_dedup(self):
        """获取去重后的总时长"""
        return sum(v.calculate_total_dedup() for v in self.volunteers.values())

    def get_total_hours_from_sub_activities(self):
        """从子活动表格计算总时长（单次活动时长的总和）"""
        sub_activities = self.get_sub_activities()
        return sum(act['duration'] for act in sub_activities)

    def get_sub_activities(self):
        """按服务类型聚合子活动"""
        sub_activities = defaultdict(lambda: {
            'times': [], 'volunteers': set(), 'durations': [],
            'time_slots': []
        })

        for name, student_id, content, time_str, duration in self.all_services:
            sub_activities[content]['times'].append(time_str)
            sub_activities[content]['volunteers'].add(student_id)
            sub_activities[content]['durations'].append(duration)
            slots = parse_time_slots(time_str)
            if slots:
                sub_activities[content]['time_slots'].extend(slots)

        result = []
        for content, data in sub_activities.items():
            volunteer_count = len(data['volunteers'])

            # 计算去重后的总时长
            if data['time_slots']:
                merged = merge_time_slots(data['time_slots'])
                total_duration = sum((end - start).total_seconds() / 3600 for start, end in merged)
            else:
                total_duration = sum(data['durations'])

            # 提取时间范围
            time_range = self._extract_time_range(data['times'])

            # 获取活动时间描述
            time_desc = self._get_time_description(data['times'])

            result.append({
                'name': content,
                'time': time_range,
                'time_desc': time_desc,
                'content': content,
                'volunteer_count': volunteer_count,
                'duration': round(total_duration, 1)
            })

        return result

    def _extract_time_range(self, times):
        """从时间字符串列表中提取时间范围"""
        if not times:
            return ''

        parsed_times = []
        for t in times:
            match = re.search(r'(\d{4})年(\d{1,2})月(\d{1,2})日\s+(\d{1,2}:\d{2})-(\d{1,2}:\d{2})', t)
            if match:
                year, month, day = match.group(1), match.group(2), match.group(3)
                start_time = match.group(4)
                end_time = match.group(5)
                parsed_times.append({
                    'date': f'{year}.{month}.{day}',
                    'start': start_time,
                    'end': end_time,
                })

        if not parsed_times:
            return times[0] if times else ''

        first = parsed_times[0]
        return f"{first['date']} {first['start']}-{first['end']}"

    def _get_time_description(self, times):
        """获取活动时间描述"""
        if not times:
            return ''

        # 提取所有日期
        dates = set()
        for t in times:
            match = re.search(r'(\d{4})年(\d{1,2})月(\d{1,2})日', t)
            if match:
                dates.add(f"{match.group(1)}.{match.group(2)}.{match.group(3)}")

        if len(dates) == 1:
            date_str = list(dates)[0]
            # 获取时间范围
            all_starts = []
            all_ends = []
            for t in times:
                match = re.search(r'(\d{1,2}:\d{2})-(\d{1,2}:\d{2})', t)
                if match:
                    all_starts.append(match.group(1))
                    all_ends.append(match.group(2))
            if all_starts and all_ends:
                return f"{date_str} {min(all_starts)}-{max(all_ends)}"
        elif len(dates) > 1:
            sorted_dates = sorted(dates)
            return f"{sorted_dates[0]}-{sorted_dates[-1]}"

        return times[0] if times else ''


def parse_time_slots(time_str):
    """解析时间字符串为时间段列表 [(datetime, datetime), ...]"""
    slots = []
    # 格式: 2026年3月27日 13:00-14:00 或 2026.3.27 13:00-14:00
    patterns = [
        r'(\d{4})年(\d{1,2})月(\d{1,2})日\s+(\d{1,2}):(\d{2})-(\d{1,2}):(\d{2})',
        r'(\d{4})\.(\d{1,2})\.(\d{1,2})\s+(\d{1,2}):(\d{2})-(\d{1,2}):(\d{2})',
        r'(\d{4})/(\d{1,2})/(\d{1,2})\s+(\d{1,2}):(\d{2})-(\d{1,2}):(\d{2})',
        r'(\d{4})-(\d{1,2})-(\d{1,2})\s+(\d{1,2}):(\d{2})-(\d{1,2}):(\d{2})',
    ]

    for pattern in patterns:
        match = re.search(pattern, time_str)
        if match:
            year, month, day = int(match.group(1)), int(match.group(2)), int(match.group(3))
            start_h, start_m = int(match.group(4)), int(match.group(5))
            end_h, end_m = int(match.group(6)), int(match.group(7))

            start = datetime(year, month, day, start_h, start_m)
            end = datetime(year, month, day, end_h, end_m)
            slots.append((start, end))
            return slots

    # 如果无法解析，尝试计算时长
    duration = calculate_duration_from_time(time_str)
    if duration > 0:
        return [duration]

    return slots


def merge_time_slots(slots):
    """合并重叠的时间段"""
    if not slots:
        return []

    # 按开始时间排序
    sorted_slots = sorted(slots, key=lambda x: x[0])
    merged = [sorted_slots[0]]

    for current in sorted_slots[1:]:
        last = merged[-1]
        if current[0] <= last[1]:  # 重叠
            merged[-1] = (last[0], max(last[1], current[1]))
        else:
            merged.append(current)

    return merged


def parse_duration(duration_str):
    """解析时长字符串，返回小时数"""
    if not duration_str:
        return 0.0

    duration_str = str(duration_str).strip()

    # 处理 "5h" 格式
    match = re.match(r'([\d.]+)h?', duration_str)
    if match:
        return float(match.group(1))

    try:
        return float(duration_str)
    except ValueError:
        return 0.0


def extract_dates_from_time(time_str):
    """从时间字符串中提取日期"""
    dates = []
    if not time_str:
        return dates

    time_str = str(time_str)

    # 支持多种日期格式
    patterns = [
        r'(\d{4})年(\d{1,2})月(\d{1,2})日',
        r'(\d{4})\.(\d{1,2})\.(\d{1,2})',
        r'(\d{4})/(\d{1,2})/(\d{1,2})',
        r'(\d{4})-(\d{1,2})-(\d{1,2})',
    ]

    for pattern in patterns:
        match = re.search(pattern, time_str)
        if match:
            year, month, day = int(match.group(1)), int(match.group(2)), int(match.group(3))
            try:
                date = datetime(year, month, day)
                dates.append(date)
            except:
                pass
            break

    return dates


def read_xlsx_data(file_path):
    """读取xlsx数据源文件（自动检测列格式，自动计算活动时间）"""
    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = wb.active

    activity_data = ActivityData()

    # 读取表头
    headers = []
    for cell in ws[1]:
        headers.append(cell.value)

    print(f"表头: {headers}")

    # 自动检测列索引
    name_col = 0
    student_id_col = 1
    content_col = 2
    time_col = 3
    duration_col = 4

    # 读取数据
    current_name = None
    current_student_id = None
    all_dates = []  # 收集所有日期

    for row in ws.iter_rows(min_row=2, values_only=True):
        name = row[name_col] if len(row) > name_col else None
        student_id = row[student_id_col] if len(row) > student_id_col else None
        service_content = row[content_col] if len(row) > content_col else None
        service_time = row[time_col] if len(row) > time_col else None
        total_hours = row[duration_col] if len(row) > duration_col else None

        if not service_content:
            continue

        if name:
            current_name = str(name).strip()
        if student_id:
            current_student_id = str(student_id).strip()

        # 提取日期
        if service_time:
            dates = extract_dates_from_time(str(service_time))
            all_dates.extend(dates)

        # 解析时长
        duration = 0.0
        if total_hours:
            duration = parse_duration(total_hours)

        # 如果没有时长，尝试从时间计算
        if duration == 0.0 and service_time:
            duration = calculate_duration_from_time(str(service_time))

        if current_name and current_student_id:
            activity_data.add_record(
                current_name,
                current_student_id,
                str(service_content).strip(),
                str(service_time).strip() if service_time else '',
                duration
            )

    # 计算每个人的总时长（简单求和）
    for vol in activity_data.volunteers.values():
        vol.calculate_total()

    # 计算活动实施时间（最早日期-最晚日期，格式：xx年xx月xx日-xx年xx月xx日）
    activity_period = ''
    if all_dates:
        min_date = min(all_dates)
        max_date = max(all_dates)
        activity_period = f'{min_date.year}年{min_date.month}月{min_date.day}日-{max_date.year}年{max_date.month}月{max_date.day}日'

    return activity_data, activity_period


def calculate_duration_from_time(time_str):
    """从时间字符串计算时长（支持多种格式）"""
    if not time_str:
        return 0.0

    time_str = str(time_str)

    # 格式1: 2026年3月27日 13:00-14:00
    match = re.search(r'(\d{1,2}):(\d{2})-(\d{1,2}):(\d{2})', time_str)
    if match:
        start_h, start_m = int(match.group(1)), int(match.group(2))
        end_h, end_m = int(match.group(3)), int(match.group(4))
        duration = (end_h * 60 + end_m - start_h * 60 - start_m) / 60.0
        return round(duration, 1)

    # 格式2: 15:30-17:00 (只有时间)
    match = re.search(r'^(\d{1,2}):(\d{2})-(\d{1,2}):(\d{2})$', time_str)
    if match:
        start_h, start_m = int(match.group(1)), int(match.group(2))
        end_h, end_m = int(match.group(3)), int(match.group(4))
        duration = (end_h * 60 + end_m - start_h * 60 - start_m) / 60.0
        return round(duration, 1)

    return 0.0


# ==================== 文档生成工具函数 ====================

def set_cell_text(cell, text, font_name='宋体', font_size=12, bold=False, alignment=WD_ALIGN_PARAGRAPH.CENTER):
    """设置单元格文本和格式"""
    cell.text = ''
    paragraph = cell.paragraphs[0]
    paragraph.alignment = alignment

    run = paragraph.add_run(str(text))
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)


def set_cell_mixed_text(cell, lines_with_bold, font_name='宋体', font_size=10.5, alignment=WD_ALIGN_PARAGRAPH.LEFT):
    """设置单元格混合格式文本（部分加粗，部分不加粗）
    lines_with_bold: list of (text, is_bold)
    """
    cell.text = ''
    paragraph = cell.paragraphs[0]
    paragraph.alignment = alignment

    for i, (text, is_bold) in enumerate(lines_with_bold):
        if i > 0:
            # 添加换行
            run = paragraph.add_run('\n')
            run.font.name = font_name
            run.font.size = Pt(font_size)
        run = paragraph.add_run(str(text))
        run.font.name = font_name
        run.font.size = Pt(font_size)
        run.font.bold = is_bold
        run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)


def set_cell_vertical_alignment(cell, align='center'):
    """设置单元格垂直对齐"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    vAlign = parse_xml(f'<w:vAlign {nsdecls("w")} w:val="{align}"/>')
    tcPr.append(vAlign)


def set_row_height(row, height_cm):
    """设置表格行高度（厘米）"""
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    trHeight = parse_xml(f'<w:trHeight {nsdecls("w")} w:val="{int(height_cm * 567)}" w:hRule="atLeast"/>')
    trPr.append(trHeight)


def merge_cells_vertical(table, col, start_row, end_row):
    """垂直合并单元格，并清除多余内容"""
    cell_start = table.cell(start_row, col)
    cell_end = table.cell(end_row, col)

    # 保存第一个单元格的内容
    first_para = cell_start.paragraphs[0] if cell_start.paragraphs else None
    first_text = first_para.text if first_para else ''
    first_runs = []
    if first_para:
        for run in first_para.runs:
            first_runs.append({
                'text': run.text,
                'bold': run.font.bold,
                'size': run.font.size,
                'name': run.font.name
            })

    # 合并单元格
    cell_start.merge(cell_end)

    # 清除合并后的多余内容，只保留第一个段落
    # 获取合并后的单元格
    merged_cell = table.cell(start_row, col)

    # 删除多余的段落，只保留第一个
    while len(merged_cell.paragraphs) > 1:
        p = merged_cell.paragraphs[-1]
        p._element.getparent().remove(p._element)

    # 恢复第一个段落的内容
    if first_runs and merged_cell.paragraphs:
        p = merged_cell.paragraphs[0]
        p.clear()
        for run_info in first_runs:
            run = p.add_run(run_info['text'])
            run.font.bold = run_info['bold']
            run.font.size = run_info['size']
            run.font.name = run_info['name']


def set_landscape_orientation(doc):
    """设置页面为横向"""
    for section in doc.sections:
        section.orientation = WD_ORIENT.LANDSCAPE
        # 交换宽高
        new_width = section.page_height
        new_height = section.page_width
        section.page_width = new_width
        section.page_height = new_height


# ==================== 志愿者名单生成（拆分单元格版） ====================

def set_row_cant_split(row):
    """设置表格行不允许跨页断行"""
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    cantSplit = parse_xml(f'<w:cantSplit {nsdecls("w")}/>')
    trPr.append(cantSplit)


def set_table_header_repeat(row):
    """设置表格行在每页重复显示（作为表头）"""
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    tblHeader = parse_xml(f'<w:tblHeader {nsdecls("w")}/>')
    trPr.append(tblHeader)


def generate_volunteer_list(activity_data, output_path):
    """生成志愿者名单汇总表（横向，拆分单元格）"""
    print("\n正在生成志愿者名单汇总表...")

    doc = Document()

    # 设置横向纸张
    set_landscape_orientation(doc)

    # 获取排序后的志愿者列表
    volunteers = activity_data.get_sorted_volunteers()

    # 创建表格（只有表头，数据行动态添加）
    table = doc.add_table(rows=1, cols=7)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # 设置表头（宋体五号 = 10.5pt）
    FONT_SIZE = 10.5  # 五号字体
    headers = ['序号', '姓名', '学号', '服务内容', '服务时间', '本项时长(h)', '总时长(h)']
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        set_cell_text(cell, header, font_name='宋体', font_size=FONT_SIZE, bold=True)
        set_cell_vertical_alignment(cell, 'center')

    # 设置表头在每页重复显示
    set_table_header_repeat(table.rows[0])

    # 填充数据
    current_row = 1  # 当前行号
    for vol_idx, vol in enumerate(volunteers, 1):
        services = vol.services
        num_services = len(services)

        # 为每个志愿者添加行（一个服务一行）
        first_row_idx = current_row
        for service_idx in range(num_services):
            new_row = table.add_row()
            content, time_str, duration = services[service_idx]

            # 服务内容、服务时间、本项时长
            set_cell_text(new_row.cells[3], content, font_name='宋体', font_size=FONT_SIZE)
            set_cell_text(new_row.cells[4], time_str, font_name='宋体', font_size=FONT_SIZE)
            set_cell_text(new_row.cells[5], str(duration), font_name='宋体', font_size=FONT_SIZE)

            # 序号、姓名、学号、总时长列先留空
            set_cell_text(new_row.cells[0], '', font_name='宋体', font_size=FONT_SIZE)
            set_cell_text(new_row.cells[1], '', font_name='宋体', font_size=FONT_SIZE)
            set_cell_text(new_row.cells[2], '', font_name='宋体', font_size=FONT_SIZE)
            set_cell_text(new_row.cells[6], '', font_name='宋体', font_size=FONT_SIZE)

            current_row += 1

        # 设置序号、姓名、学号、总时长（在第一行）
        last_row_idx = current_row - 1
        set_cell_text(table.rows[first_row_idx].cells[0], str(vol_idx), font_name='宋体', font_size=FONT_SIZE)
        set_cell_text(table.rows[first_row_idx].cells[1], vol.name, font_name='宋体', font_size=FONT_SIZE)
        set_cell_text(table.rows[first_row_idx].cells[2], vol.student_id, font_name='宋体', font_size=FONT_SIZE)
        set_cell_text(table.rows[first_row_idx].cells[6], str(vol.total_hours), font_name='宋体', font_size=FONT_SIZE)

        # 如果有多项服务，合并序号、姓名、学号、总时长的单元格
        if num_services > 1:
            merge_cells_vertical(table, 0, first_row_idx, last_row_idx)
            merge_cells_vertical(table, 1, first_row_idx, last_row_idx)
            merge_cells_vertical(table, 2, first_row_idx, last_row_idx)
            merge_cells_vertical(table, 6, first_row_idx, last_row_idx)

        # 垂直居中
        for col in [0, 1, 2, 6]:
            set_cell_vertical_alignment(table.rows[first_row_idx].cells[col], 'center')

        # 设置该志愿者的所有行不允许跨页断行
        for row_idx in range(first_row_idx, last_row_idx + 1):
            set_row_cant_split(table.rows[row_idx])

    doc.save(output_path)
    print(f"志愿者名单汇总表已保存: {output_path}")


# ==================== 申报表生成 ====================

def generate_declaration_form(activity_data, activity_name, output_path,
                               beneficiary_type='重庆交通大学全体在校学生、活动对接的老年群体',
                               beneficiary_count=None,
                               activity_period=None,
                               activity_unit='信息学院学生会',
                               activity_overview=None,
                               activity_background=None,
                               activity_significance=None,
                               activity_results=None):
    """生成志愿服务活动时长申报表（完全按照模板格式）"""
    print("\n正在生成志愿服务活动时长申报表...")

    doc = Document()
    FONT_SIZE = 10.5  # 宋体五号
    HEADER_FONT_SIZE = 12  # 宋体小四（用于标题）

    # 添加标题
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run('志愿服务活动时长申报表')
    run.font.name = '宋体'
    run.font.size = Pt(HEADER_FONT_SIZE)
    run.font.bold = True
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    # ===== 表格1：活动基本信息（8行5列） =====
    table1 = doc.add_table(rows=8, cols=5)
    table1.style = 'Table Grid'
    table1.alignment = WD_TABLE_ALIGNMENT.CENTER

    # 设置行高使布局更稀松
    for row in table1.rows:
        set_row_height(row, 1.0)  # 每行至少1厘米高

    # 第0行：一、活动基本信息（合并全部5列）- 宋体小四加粗
    cell = table1.cell(0, 0).merge(table1.cell(0, 4))
    set_cell_text(cell, '一、活动基本信息', font_name='宋体', font_size=HEADER_FONT_SIZE, bold=True)
    set_cell_vertical_alignment(cell, 'center')

    # 第1行：活动名称（左2列合并为标签，右3列合并为值）
    cell_label = table1.cell(1, 0).merge(table1.cell(1, 1))
    set_cell_text(cell_label, '活动名称', font_name='宋体', font_size=FONT_SIZE)
    cell_value = table1.cell(1, 2).merge(table1.cell(1, 4))
    set_cell_text(cell_value, activity_name, font_name='宋体', font_size=FONT_SIZE)

    # 第2行：活动类型（左右两列）
    cell_label = table1.cell(2, 0).merge(table1.cell(2, 1))
    set_cell_text(cell_label, '活动类型', font_name='宋体', font_size=FONT_SIZE)
    cell_value = table1.cell(2, 2).merge(table1.cell(2, 4))
    # 使用制表符对齐左右两列
    activity_types = ('（1）乡村振兴服务  □\t（2）弱势群体服务  □\n'
                      '（3）社会发展服务  □\t（4）生态文明服务  □\n'
                      '（5）大型赛会服务  □\t（6）抢险救灾服务  □\n'
                      '（7）文明建设服务  □\t（8）校园公益服务  ✓')
    set_cell_text(cell_value, activity_types, font_name='宋体', font_size=FONT_SIZE)

    # 第3行：受益人
    cell_label = table1.cell(3, 0).merge(table1.cell(3, 1))
    set_cell_text(cell_label, '受益人：\n类型及数量', font_name='宋体', font_size=FONT_SIZE)
    set_cell_text(table1.cell(3, 2), beneficiary_type, font_name='宋体', font_size=FONT_SIZE)
    set_cell_text(table1.cell(3, 3), '受益人数', font_name='宋体', font_size=FONT_SIZE)
    set_cell_text(table1.cell(3, 4), str(beneficiary_count or 200), font_name='宋体', font_size=FONT_SIZE)

    # 第4行：活动实施时间
    cell_label = table1.cell(4, 0).merge(table1.cell(4, 1))
    set_cell_text(cell_label, '活动实施时间', font_name='宋体', font_size=FONT_SIZE)
    cell_value = table1.cell(4, 2).merge(table1.cell(4, 4))
    set_cell_text(cell_value, activity_period or '请填写活动时间', font_name='宋体', font_size=FONT_SIZE)

    # 第5行：活动开展单位
    cell_label = table1.cell(5, 0).merge(table1.cell(5, 1))
    set_cell_text(cell_label, '活动开展单位', font_name='宋体', font_size=FONT_SIZE)
    cell_value = table1.cell(5, 2).merge(table1.cell(5, 4))
    set_cell_text(cell_value, activity_unit, font_name='宋体', font_size=FONT_SIZE)

    # 第6行：二、活动详细信息（合并全部5列）- 宋体小四加粗
    cell = table1.cell(6, 0).merge(table1.cell(6, 4))
    set_cell_text(cell, '二、活动详细信息', font_name='宋体', font_size=HEADER_FONT_SIZE, bold=True)

    # 第7行：活动内容
    cell_label = table1.cell(7, 0).merge(table1.cell(7, 1))
    set_cell_text(cell_label, '活动内容', font_name='宋体', font_size=FONT_SIZE)
    cell_value = table1.cell(7, 2).merge(table1.cell(7, 4))

    # 生成活动内容（按照模板格式，标签加粗）
    if activity_overview:
        # 如果用户提供了自定义内容，直接使用
        set_cell_text(cell_value, activity_overview, font_name='宋体', font_size=FONT_SIZE, alignment=WD_ALIGN_PARAGRAPH.LEFT)
    else:
        # 使用混合格式，标签加粗
        overview_lines = generate_activity_overview_lines(
            activity_name,
            activity_background or '校园内',
            activity_significance or '弘扬志愿服务精神，为同学们提供志愿服务',
            activity_results or '活动圆满结束'
        )
        set_cell_mixed_text(cell_value, overview_lines, font_name='宋体', font_size=FONT_SIZE)

    # ===== 表格2：子活动信息（6列） =====
    # 添加分页符，使表格2在新页面开始
    doc.add_page_break()
    table2 = doc.add_table(rows=1, cols=6)
    table2.style = 'Table Grid'
    table2.alignment = WD_TABLE_ALIGNMENT.CENTER

    # 第0行：表头
    sub_headers = ['具体活动', '子活动名称', '活动时间', '活动内容', '志愿者人数', '单次活动服务时长']
    for i, header in enumerate(sub_headers):
        set_cell_text(table2.rows[0].cells[i], header, font_name='宋体', font_size=FONT_SIZE, bold=True)

    # 填充子活动数据
    sub_activities = activity_data.get_sub_activities()
    first_data_row = None
    last_data_row = None
    for idx, sub_act in enumerate(sub_activities):
        row = table2.add_row()
        # 第一列留空，后面合并后统一设置"具体活动"
        set_cell_text(row.cells[0], '', font_name='宋体', font_size=FONT_SIZE)
        set_cell_text(row.cells[1], sub_act['name'], font_name='宋体', font_size=FONT_SIZE)
        set_cell_text(row.cells[2], sub_act['time_desc'], font_name='宋体', font_size=FONT_SIZE)
        set_cell_text(row.cells[3], sub_act['content'], font_name='宋体', font_size=FONT_SIZE)
        set_cell_text(row.cells[4], str(sub_act['volunteer_count']), font_name='宋体', font_size=FONT_SIZE)
        set_cell_text(row.cells[5], f"{sub_act['duration']}h", font_name='宋体', font_size=FONT_SIZE)
        if idx == 0:
            first_data_row = 1
        last_data_row = idx + 1

    # 合并"具体活动"列（所有子活动行），并设置居中显示"具体活动"
    if first_data_row and last_data_row:
        if first_data_row < last_data_row:
            merge_cells_vertical(table2, 0, first_data_row, last_data_row)
        # 设置"具体活动"文字，垂直居中
        merged_cell = table2.cell(first_data_row, 0)
        set_cell_text(merged_cell, '具体活动', font_name='宋体', font_size=FONT_SIZE)
        set_cell_vertical_alignment(merged_cell, 'center')

    # ===== 三、审核意见 =====
    # 添加三、审核意见标题 - 宋体小四加粗
    review_title_row = table2.add_row()
    cell = review_title_row.cells[0].merge(review_title_row.cells[5])
    set_cell_text(cell, '三、审核意见', font_name='宋体', font_size=HEADER_FONT_SIZE, bold=True)

    # 添加申请单位承诺行（按照参考图片格式）
    seal_row = table2.add_row()
    cell = seal_row.cells[0].merge(seal_row.cells[5])
    cell.text = ''
    # 申请单位承诺标题
    p_title = cell.paragraphs[0]
    p_title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run_title = p_title.add_run('申请单位承诺：')
    run_title.font.name = '宋体'
    run_title.font.size = Pt(FONT_SIZE)
    run_title._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    # 3条承诺内容
    commitments = [
        '1.严格遵守《重庆交通大学志愿服务管理办法》相关规定；',
        '2.确保活动期间的安全、有序，维护志愿者合法权益；',
        '3.不擅自任意更改申报内容，无违法、违纪的活动行为。'
    ]
    for commit in commitments:
        p_commit = cell.add_paragraph()
        p_commit.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run_commit = p_commit.add_run(commit)
        run_commit.font.name = '宋体'
        run_commit.font.size = Pt(FONT_SIZE)
        run_commit._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    # 负责人签字和日期（右对齐）
    p_sign = cell.add_paragraph()
    p_sign.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_sign = p_sign.add_run('负责人签字：        日期：')
    run_sign.font.name = '宋体'
    run_sign.font.size = Pt(FONT_SIZE)
    run_sign._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    # 添加志工部审批意见行
    opinion_row = table2.add_row()
    cell = opinion_row.cells[0].merge(opinion_row.cells[5])
    cell.text = ''
    # 志工部审批意见标题
    p_opinion_title = cell.paragraphs[0]
    p_opinion_title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run_opinion_title = p_opinion_title.add_run('志工部审批意见：')
    run_opinion_title.font.name = '宋体'
    run_opinion_title.font.size = Pt(FONT_SIZE)
    run_opinion_title._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    # 添加空行使签字在下方
    for _ in range(3):
        p_empty = cell.add_paragraph()
        p_empty.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # 签字和日期（右对齐）
    p_opinion_sign = cell.add_paragraph()
    p_opinion_sign.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_opinion_sign = p_opinion_sign.add_run('负责人签字：        日期：')
    run_opinion_sign.font.name = '宋体'
    run_opinion_sign.font.size = Pt(FONT_SIZE)
    run_opinion_sign._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    doc.save(output_path)
    print(f"志愿服务活动时长申报表已保存: {output_path}")


def generate_activity_overview(activity_name, background, significance, results):
    """生成活动概述内容（按照模板格式：6个字段，标签加粗）"""
    # 注意：由于set_cell_text无法在同一个单元格中混合加粗和非加粗文本，
    # 这里用特殊标记表示加粗标签，后续可以改进
    overview = f"内容概述\n"
    overview += f"协助开展{activity_name}活动的工作，确保{activity_name}顺利完成。\n\n"
    overview += f"社会问题\n"
    overview += f"当前校园内学生课余生活相对单一，部分同学面临学业压力和心理压力，缺乏有效的放松和调节途径。同时，校园内部分公共区域需要维护和整理，老年群体也需要更多的关怀和陪伴。这些问题需要通过志愿服务活动来缓解和改善。\n\n"
    overview += f"产生原因\n"
    overview += f"学生学业负担较重，课余时间有限，缺乏组织化的志愿服务活动来引导同学们参与社会实践。校园公共区域维护需求增加，而学校后勤人员有限，无法全面覆盖。老年群体因身体原因需要更多的陪伴和帮助，但社会资源分配不均导致服务供给不足。\n\n"
    overview += f"必要性\n"
    overview += f"开展{activity_name}活动有利于增强同学们的社会责任感和奉献精神；有利于缓解同学们的学业压力，丰富课余生活；有利于改善校园环境，提升校园文明程度；有利于弘扬中华民族尊老爱幼的传统美德；有利于培养同学们的团队协作能力和组织能力；有利于促进校园文化建设，营造积极向上的校园氛围；有利于推动志愿服务事业的发展，传递正能量。\n\n"
    overview += f"服务对象\n"
    overview += f"同活动信息中的受益人。\n\n"
    overview += f"预计成效\n"
    overview += f"重庆交通大学信息学院{activity_name}活动圆满结束。"
    return overview


def generate_activity_overview_lines(activity_name, background, significance, results):
    """生成活动概述内容（返回带加粗标记的行列表）"""
    lines = [
        ("内容概述", True),
        ("协助开展" + activity_name + "活动的工作，确保" + activity_name + "顺利完成。", False),
        ("社会问题", True),
        ("当前校园内学生课余生活相对单一，部分同学面临学业压力和心理压力，缺乏有效的放松和调节途径。同时，校园内部分公共区域需要维护和整理，老年群体也需要更多的关怀和陪伴。这些问题需要通过志愿服务活动来缓解和改善。", False),
        ("产生原因", True),
        ("学生学业负担较重，课余时间有限，缺乏组织化的志愿服务活动来引导同学们参与社会实践。校园公共区域维护需求增加，而学校后勤人员有限，无法全面覆盖。老年群体因身体原因需要更多的陪伴和帮助，但社会资源分配不均导致服务供给不足。", False),
        ("必要性", True),
        ("开展" + activity_name + "活动有利于增强同学们的社会责任感和奉献精神；有利于缓解同学们的学业压力，丰富课余生活；有利于改善校园环境，提升校园文明程度；有利于弘扬中华民族尊老爱幼的传统美德；有利于培养同学们的团队协作能力和组织能力；有利于促进校园文化建设，营造积极向上的校园氛围；有利于推动志愿服务事业的发展，传递正能量。", False),
        ("服务对象", True),
        ("同活动信息中的受益人。", False),
        ("预计成效", True),
        ("重庆交通大学信息学院" + activity_name + "活动圆满结束。", False),
    ]
    return lines


# ==================== 活动总结生成 ====================

def generate_summary(activity_data, activity_name, output_path,
                     activity_period=None,
                     activity_overview=None,
                     activity_results=None):
    """生成志愿服务活动总结"""
    print("\n正在生成志愿服务活动总结...")

    doc = Document()

    # 计算统计数据
    total_volunteers = activity_data.get_total_volunteers()
    total_hours = activity_data.get_total_hours_from_sub_activities()

    # 尝试使用方正仿宋字体，如果没有则使用宋体
    font_name = '方正仿宋_GBK'
    # 检测字体是否可用（简化处理，直接尝试使用）

    # 标题
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run(f'信息学院学生会{activity_name}总结')
    run.font.name = font_name
    run.font.size = Pt(20)
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)

    # 正文第一段
    para1 = doc.add_paragraph()
    para1.alignment = WD_ALIGN_PARAGRAPH.LEFT
    para1.paragraph_format.first_line_indent = Pt(32)  # 首行缩进

    if activity_period is None:
        activity_period = '请填写活动时间'

    if activity_overview is None:
        activity_overview = f'{activity_name}，弘扬志愿服务精神，为同学们提供志愿服务，增强同学们的校园归属感'

    text1 = (f'信息学院学生会于{activity_period}开展{activity_name}志愿服务，'
             f'参与志愿者共{total_volunteers}人，志愿时长共计{total_hours}小时。'
             f'有效地保障了{activity_name}活动的顺利开展。'
             f'{activity_overview}，'
             f'活动反响热烈，信息学院和学生会将继续发扬奉献精神和社会责任感，传递正能量。'
             f'希望能为校园文化建设和学院学生会发展出一份力，广受好评。')

    run1 = para1.add_run(text1)
    run1.font.name = font_name
    run1.font.size = Pt(16)
    run1._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)

    # 正文第二段
    para2 = doc.add_paragraph()
    para2.alignment = WD_ALIGN_PARAGRAPH.LEFT
    para2.paragraph_format.first_line_indent = Pt(32)

    if activity_results is None:
        activity_results = f'{activity_name}活动得到了参与志愿活动的老师和同学们的高度认可。'

    run2 = para2.add_run(f'活动成效：{activity_results}')
    run2.font.name = font_name
    run2.font.size = Pt(16)
    run2._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)

    doc.save(output_path)
    print(f"志愿服务活动总结已保存: {output_path}")


# ==================== 活动内容智能生成 ====================

def analyze_activity_type(activity_name):
    """分析活动名称，确定活动类型"""
    activity_name = activity_name.lower()

    # 关键词到活动类型的映射
    keyword_map = {
        '雷锋': '学雷锋',
        '环保': '环保', '绿色': '环保', '生态': '环保',
        '敬老': '敬老', '老人': '敬老', '养老': '敬老',
        '支教': '支教', '教育': '支教', '助学': '支教',
        '社区': '社区', '服务': '社区',
        '公益': '公益', '志愿': '公益',
        '文化': '文化', '传统': '文化',
        '体育': '体育', '运动': '体育',
        '科技': '科技', '创新': '科技',
        '安全': '安全', '消防': '安全',
        '健康': '健康', '医疗': '健康',
    }

    for keyword, activity_type in keyword_map.items():
        if keyword in activity_name:
            return activity_type

    return '通用'


def generate_activity_content(activity_name, activity_type=None):
    """根据活动生成内容（社会问题、产生原因、必要性等）"""
    if activity_type is None:
        activity_type = analyze_activity_type(activity_name)

    # 不同活动类型的内容模板
    templates = {
        '学雷锋': {
            'social_problem': '当前社会中，部分群众特别是老年人和弱势群体面临生活不便、精神孤独等问题。校园内学生课余生活相对单一，部分同学缺乏社会实践经验和志愿服务意识。雷锋精神作为中华民族传统美德的代表，需要在新时代得到传承和弘扬。',
            'cause': '随着社会节奏加快，人们忙于工作和学习，对身边需要帮助的人关注不足。部分学生缺乏主动服务社会的意识和机会，志愿服务活动组织不够系统化。社会资源分配不均，导致部分群体得不到足够的关怀和帮助。',
            'necessity': '开展雷锋月志愿活动有利于弘扬雷锋精神，传承中华民族传统美德；有利于增强同学们的社会责任感和奉献精神；有利于丰富同学们的课余生活，提升综合素质；有利于促进校园精神文明建设；有利于营造互帮互助的良好社会氛围；有利于培养同学们的团队协作能力；有利于推动志愿服务事业的发展。',
            'beneficiary': '校园内学生、社区居民、老年群体',
            'result': '活动圆满结束，得到了参与志愿活动的老师和同学们的高度认可',
        },
        '环保': {
            'social_problem': '当前环境问题日益突出，校园内部分区域存在垃圾乱扔、资源浪费等现象。同学们的环保意识有待提高，绿色生活方式尚未全面普及。生态环境保护需要每个人的参与和努力。',
            'cause': '环保宣传教育不够深入，部分同学对环境保护的重要性认识不足。校园环保设施不够完善，垃圾分类执行不到位。缺乏系统性的环保志愿活动来引导同学们参与环保实践。',
            'necessity': '开展环保志愿活动有利于增强同学们的环保意识和生态文明理念；有利于改善校园环境，打造绿色校园；有利于推动垃圾分类和资源循环利用；有利于培养同学们的环保习惯和责任感；有利于促进可持续发展理念的传播；有利于提升校园文明程度；有利于为建设美丽中国贡献力量。',
            'beneficiary': '校园师生、周边社区居民',
            'result': '活动圆满结束，有效改善了校园环境，提升了同学们的环保意识',
        },
        '敬老': {
            'social_problem': '当前社会老龄化程度加深，老年群体面临生活不便、精神孤独、缺乏陪伴等问题。部分老年人因身体原因行动不便，需要更多的关怀和帮助。大学生群体与老年人之间的交流互动不足。',
            'cause': '随着社会节奏加快，年轻人忙于工作学习，陪伴老人的时间减少。养老机构和社区服务资源有限，无法满足所有老年人的需求。大学生缺乏与老年人交流的平台和机会。',
            'necessity': '开展敬老志愿活动有利于弘扬中华民族尊老爱幼的传统美德；有利于增强同学们的社会责任感和感恩意识；有利于丰富老年人的精神文化生活；有利于促进代际交流和理解；有利于培养同学们的爱心和耐心；有利于营造尊老敬老的良好社会氛围；有利于推动养老服务体系的完善。',
            'beneficiary': '敬老院老人、社区独居老人',
            'result': '活动圆满结束，为老人们带去了温暖和欢乐，得到了老人们和工作人员的高度评价',
        },
        '支教': {
            'social_problem': '当前教育资源分布不均，部分偏远地区和弱势群体家庭的孩子面临教育资源匮乏、学习条件落后等问题。大学生群体拥有丰富的知识和技能，可以通过志愿服务帮助这些孩子。',
            'cause': '城乡教育资源差距较大，部分学校师资力量不足。贫困家庭的孩子缺乏课外辅导和学习指导。大学生志愿服务支教活动可以有效补充教育资源，帮助孩子们开阔视野。',
            'necessity': '开展支教志愿活动有利于促进教育公平，帮助弱势群体获得更好的教育；有利于增强同学们的社会责任感和使命感；有利于丰富孩子们的课余生活，开阔视野；有利于培养同学们的教学能力和沟通能力；有利于促进城乡教育交流；有利于弘扬知识改变命运的理念；有利于推动教育事业的发展。',
            'beneficiary': '偏远地区学生、贫困家庭孩子',
            'result': '活动圆满结束，帮助孩子们提升了学习成绩和综合素质，得到了学校和家长的高度认可',
        },
        '社区': {
            'social_problem': '当前社区治理面临诸多挑战，部分社区存在环境脏乱、邻里关系淡漠、公共服务不足等问题。大学生志愿者可以参与社区服务，为社区建设贡献力量。',
            'cause': '社区工作人员有限，难以全面覆盖各项服务。居民参与社区治理的积极性不高。缺乏有效的社区志愿服务机制来动员社会力量参与社区建设。',
            'necessity': '开展社区志愿活动有利于改善社区环境，提升居民生活质量；有利于增强同学们的社会实践能力和服务意识；有利于促进邻里和谐，增强社区凝聚力；有利于推动社区治理现代化；有利于培养同学们的公民意识和责任感；有利于弘扬志愿服务精神；有利于构建和谐社会。',
            'beneficiary': '社区居民、社区工作者',
            'result': '活动圆满结束，有效改善了社区环境，提升了居民满意度',
        },
        '通用': {
            'social_problem': '当前社会中存在一些需要关注和解决的问题，大学生作为社会的重要力量，可以通过志愿服务活动为社会发展贡献力量。志愿服务活动有助于培养同学们的社会责任感和奉献精神。',
            'cause': '社会发展过程中，部分群体面临各种困难和挑战，需要社会各界的关注和帮助。大学生群体拥有知识和活力，可以通过志愿服务活动为社会做出贡献。',
            'necessity': f'开展{activity_name}活动有利于增强同学们的社会责任感和奉献精神；有利于丰富同学们的课余生活，提升综合素质；有利于促进校园精神文明建设；有利于营造互帮互助的良好社会氛围；有利于培养同学们的团队协作能力；有利于推动志愿服务事业的发展；有利于构建和谐社会。',
            'beneficiary': '校园师生、社会群众',
            'result': '活动圆满结束，得到了参与志愿活动的老师和同学们的高度认可',
        },
    }

    # 获取模板，如果没有匹配的类型则使用通用模板
    template = templates.get(activity_type, templates['通用'])

    return {
        'social_problem': template['social_problem'],
        'cause': template['cause'],
        'necessity': template['necessity'],
        'beneficiary': template['beneficiary'],
        'result': template['result'],
    }


# ==================== 用户交互 ====================

def get_user_input():
    """获取用户输入（简化版）"""
    print("\n" + "="*60)
    print("志愿活动材料自动撰写脚本")
    print("="*60)

    # 查找xlsx文件
    script_dir = os.path.dirname(os.path.abspath(__file__))
    xlsx_files = [f for f in os.listdir(script_dir) if f.endswith('.xlsx') and not f.startswith('~$')]

    print("\n请选择数据源文件：")

    # 显示当前文件夹下的xlsx文件
    if xlsx_files:
        print("当前文件夹下的xlsx文件：")
        for i, f in enumerate(xlsx_files, 1):
            print(f"  {i}. {f}")
        print(f"  {len(xlsx_files) + 1}. 手动输入文件路径")
    else:
        print("当前文件夹下未找到xlsx文件")
        print("  1. 手动输入文件路径")

    # 选择数据源文件
    while True:
        try:
            choice = input("\n请选择（输入编号或直接输入文件路径）: ").strip()

            # 去除可能的引号
            if choice.startswith('"') and choice.endswith('"'):
                choice = choice[1:-1]
            if choice.startswith("'") and choice.endswith("'"):
                choice = choice[1:-1]

            # 检查是否是文件路径（包含路径分隔符或以盘符开头）
            is_path = ('\\' in choice or '/' in choice or
                       (len(choice) >= 2 and choice[1] == ':') or
                       choice.endswith('.xlsx') or choice.endswith('.xls'))

            if is_path:
                # 处理Windows路径
                xlsx_path = os.path.normpath(choice)
                # 尝试绝对路径
                if os.path.isabs(xlsx_path):
                    if os.path.exists(xlsx_path):
                        print(f"已选择文件: {xlsx_path}")
                        break
                    else:
                        print(f"文件不存在: {xlsx_path}")
                        print("请检查路径是否正确，或重新输入")
                        continue
                else:
                    # 相对路径，尝试相对于当前目录
                    abs_path = os.path.abspath(xlsx_path)
                    if os.path.exists(abs_path):
                        xlsx_path = abs_path
                        print(f"已选择文件: {xlsx_path}")
                        break
                    else:
                        print(f"文件不存在: {abs_path}")
                        print("请检查路径是否正确，或重新输入")
                        continue

            # 检查是否是编号
            choice_num = int(choice)
            if xlsx_files and 1 <= choice_num <= len(xlsx_files):
                xlsx_path = os.path.join(script_dir, xlsx_files[choice_num - 1])
                break
            elif choice_num == len(xlsx_files) + 1 or (not xlsx_files and choice_num == 1):
                xlsx_path = input("请输入xlsx文件的完整路径: ").strip()
                # 去除可能的引号
                if xlsx_path.startswith('"') and xlsx_path.endswith('"'):
                    xlsx_path = xlsx_path[1:-1]
                xlsx_path = os.path.normpath(xlsx_path)
                if os.path.exists(xlsx_path):
                    print(f"已选择文件: {xlsx_path}")
                    break
                else:
                    print(f"文件不存在: {xlsx_path}")
                    print("请检查路径是否正确，或重新输入")
            else:
                print("无效的选择，请重新输入")
        except ValueError:
            print("请输入有效的数字或文件路径")

    # 获取活动名称
    activity_name = input("\n请输入志愿活动名称（如：雷锋月志愿活动）: ").strip()
    if not activity_name:
        activity_name = '志愿活动'
        print(f"使用默认名称: {activity_name}")

    # 获取输出目录
    script_dir = os.path.dirname(os.path.abspath(__file__))
    default_output_dir = os.path.join(script_dir, 'output')
    print(f"\n默认输出目录: {default_output_dir}")
    output_dir_input = input("请输入输出目录（留空使用默认目录）: ").strip()

    # 去除可能的引号
    if output_dir_input.startswith('"') and output_dir_input.endswith('"'):
        output_dir_input = output_dir_input[1:-1]
    if output_dir_input.startswith("'") and output_dir_input.endswith("'"):
        output_dir_input = output_dir_input[1:-1]

    if output_dir_input:
        output_dir = os.path.normpath(output_dir_input)
    else:
        output_dir = default_output_dir

    return {
        'xlsx_path': xlsx_path,
        'activity_name': activity_name,
        'output_dir': output_dir,
    }


# ==================== 主程序 ====================

def main():
    """主程序"""
    config = get_user_input()

    print(f"\n正在读取数据源: {config['xlsx_path']}")
    activity_data, activity_period = read_xlsx_data(config['xlsx_path'])

    total_hours = activity_data.get_total_hours_from_sub_activities()
    print(f"\n数据读取完成:")
    print(f"  - 志愿者人数: {activity_data.get_total_volunteers()}")
    print(f"  - 总服务时长: {total_hours} 小时")
    print(f"  - 服务类型: {len(activity_data.service_types)} 种")
    print(f"  - 活动实施时间: {activity_period}")

    # 智能生成活动内容
    activity_name = config['activity_name']
    activity_type = analyze_activity_type(activity_name)
    content = generate_activity_content(activity_name, activity_type)

    print(f"\n活动类型识别: {activity_type}")
    print("已自动生成活动内容（社会问题、产生原因、必要性等）")

    # 使用用户指定的输出目录
    output_dir = config['output_dir']
    os.makedirs(output_dir, exist_ok=True)

    safe_name = activity_name.replace('/', '_').replace('\\', '_')

    volunteer_list_path = os.path.join(output_dir, f'{safe_name}志愿者名单.docx')
    declaration_path = os.path.join(output_dir, f'{safe_name}申报表.docx')
    summary_path = os.path.join(output_dir, f'{safe_name}活动总结.docx')

    # 删除旧文件
    for old_file in [volunteer_list_path, declaration_path, summary_path]:
        if os.path.exists(old_file):
            try:
                os.remove(old_file)
                print(f"已删除旧文件: {os.path.basename(old_file)}")
            except PermissionError:
                print(f"警告: 无法删除旧文件 {os.path.basename(old_file)}，请先关闭该文件")

    # 1. 生成志愿者名单汇总表
    generate_volunteer_list(activity_data, volunteer_list_path)

    # 2. 生成志愿服务活动时长申报表
    generate_declaration_form(
        activity_data,
        activity_name,
        declaration_path,
        activity_period=activity_period,
        activity_background=content['social_problem'],
        activity_significance=content['necessity'],
        activity_results=content['result'],
        beneficiary_type=content['beneficiary'],
    )

    # 3. 生成志愿服务活动总结
    generate_summary(
        activity_data,
        activity_name,
        summary_path,
        activity_period=activity_period,
        activity_results=content['result'],
    )

    print("\n" + "="*60)
    print("所有材料生成完成！")
    print("="*60)
    print(f"\n输出目录: {output_dir}")
    print(f"\n生成的文件:")
    print(f"  1. {safe_name}志愿者名单.docx")
    print(f"  2. {safe_name}申报表.docx")
    print(f"  3. {safe_name}活动总结.docx")

    os.startfile(output_dir)


if __name__ == '__main__':
    main()
