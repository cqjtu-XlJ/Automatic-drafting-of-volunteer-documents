# -*- coding: utf-8 -*-
"""
志愿活动材料自动撰写脚本 - 前端界面
基于 Streamlit 构建
"""

import os
import sys
import tempfile
import streamlit as st

# 添加当前目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from generate_materials import (
    read_xlsx_data, generate_activity_content, analyze_activity_type,
    generate_volunteer_list, generate_declaration_form, generate_summary
)
from llm_generator import generate_activity_content_with_llm, get_available_providers

# 页面配置
st.set_page_config(
    page_title="志愿活动材料自动撰写",
    page_icon="📝",
    layout="wide"
)

# 标题
st.title("📝 志愿活动材料自动撰写脚本")
st.markdown("---")

# 侧边栏配置
with st.sidebar:
    st.header("⚙️ 配置")

    # LLM 提供商选择
    st.subheader("🤖 AI 设置")
    available_providers = get_available_providers()

    provider_names = {
        'zhipu': '智谱AI (GLM-4-Flash, 免费)',
        'deepseek': 'DeepSeek',
        'qwen': '通义千问',
        'moonshot': '月之暗面 Kimi',
        'mimo': '小米 MIMO',
        'doubao': '火山引擎 Doubao'
    }

    provider_links = {
        'zhipu': 'https://open.bigmodel.cn',
        'deepseek': 'https://platform.deepseek.com',
        'qwen': 'https://dashscope.aliyuncs.com',
        'moonshot': 'https://platform.moonshot.cn',
        'mimo': 'https://token-plan-cn.xiaomimimo.com',
        'doubao': 'https://console.volcengine.com'
    }

    use_llm = False
    llm_provider = None
    api_key = None

    if available_providers:
        use_llm = st.checkbox("启用 AI 生成内容", value=True)
        if use_llm:
            provider_options = [provider_names.get(p, p) for p in available_providers]
            selected = st.selectbox("选择 AI 提供商", provider_options)
            llm_provider = available_providers[provider_options.index(selected)]

            # API Key 输入
            api_key = st.text_input(
                "API Key",
                type="password",
                help=f"请输入 {provider_names.get(llm_provider, llm_provider)} 的 API Key"
            )

            if not api_key:
                st.warning("⚠️ 请输入 API Key 后才能使用 AI 功能")
                st.markdown(f"[获取 API Key →]({provider_links.get(llm_provider, '#')})")
    else:
        st.info("未检测到可用的 AI 提供商")

    st.markdown("---")

    # 输出目录
    st.subheader("📁 输出设置")
    default_output = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'output')
    output_dir = st.text_input("输出目录", value=default_output)

# 主区域
col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("📤 上传数据源")
    uploaded_file = st.file_uploader(
        "选择 xlsx 文件",
        type=['xlsx'],
        help="上传包含志愿者信息的 Excel 文件"
    )

    if uploaded_file:
        st.success(f"已上传: {uploaded_file.name}")

with col2:
    st.subheader("📝 活动信息")
    activity_name = st.text_input(
        "活动名称 *",
        placeholder="例如：雷锋月志愿活动",
        help="必填，用于生成材料标题和内容"
    )

    beneficiary_type = st.text_input(
        "受益人类型",
        value="重庆交通大学全体在校学生、活动对接的老年群体",
        help="活动的受益群体描述"
    )

st.markdown("---")

# 生成按钮
if st.button("🚀 生成材料", type="primary", use_container_width=True):
    # 验证输入
    if not uploaded_file:
        st.error("请上传 xlsx 数据源文件")
    elif not activity_name:
        st.error("请输入活动名称")
    elif use_llm and not api_key:
        st.error("请输入 API Key 后才能使用 AI 功能")
    else:
        # 保存上传的文件到临时目录
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            tmp.write(uploaded_file.getvalue())
            xlsx_path = tmp.name

        try:
            # 进度条
            progress_bar = st.progress(0, text="正在处理...")

            # 1. 读取数据
            progress_bar.progress(10, text="正在读取数据源...")
            activity_data, activity_period = read_xlsx_data(xlsx_path)

            total_hours = activity_data.get_total_hours_from_sub_activities()
            st.info(f"📊 数据读取完成：{activity_data.get_total_volunteers()} 名志愿者，总时长 {total_hours} 小时，活动时间 {activity_period}")

            # 2. 生成内容
            progress_bar.progress(30, text="正在生成活动内容...")
            activity_type = analyze_activity_type(activity_name)

            content = None
            if use_llm and llm_provider and api_key:
                progress_bar.progress(40, text=f"正在调用 AI 生成内容...")
                service_contents = ', '.join(activity_data.service_types)
                content = generate_activity_content_with_llm(
                    activity_name, activity_type,
                    activity_data.get_total_volunteers(), total_hours,
                    service_contents, provider=llm_provider, api_key=api_key
                )
                if content:
                    st.success("✅ AI 内容生成成功")
                else:
                    st.warning("⚠️ AI 生成失败，将使用静态模板")

            if not content:
                content = generate_activity_content(activity_name, activity_type)
                st.info("📝 已使用静态模板生成内容")

            # 3. 创建输出目录
            os.makedirs(output_dir, exist_ok=True)

            safe_name = activity_name.replace('/', '_').replace('\\', '_')

            # 4. 生成文件
            progress_bar.progress(60, text="正在生成志愿者名单...")
            volunteer_list_path = os.path.join(output_dir, f'{safe_name}志愿者名单.docx')
            generate_volunteer_list(activity_data, volunteer_list_path)

            progress_bar.progress(70, text="正在生成申报表...")
            declaration_path = os.path.join(output_dir, f'{safe_name}申报表.docx')
            generate_declaration_form(
                activity_data, activity_name, declaration_path,
                activity_period=activity_period,
                activity_background=content['social_problem'],
                activity_significance=content['necessity'],
                activity_results=f"重庆交通大学{activity_name}活动圆满结束。",
                beneficiary_type=beneficiary_type,
            )

            progress_bar.progress(85, text="正在生成活动总结...")
            summary_path = os.path.join(output_dir, f'{safe_name}活动总结.docx')
            generate_summary(
                activity_data, activity_name, summary_path,
                activity_period=activity_period,
                activity_results=f"重庆交通大学{activity_name}活动圆满结束。"
            )

            progress_bar.progress(100, text="✅ 全部完成！")

            st.markdown("---")
            st.subheader("📄 生成结果")

            # 显示生成的文件
            col1, col2, col3 = st.columns(3)

            with col1:
                st.metric("志愿者名单", f"{safe_name}志愿者名单.docx")
                with open(volunteer_list_path, 'rb') as f:
                    st.download_button(
                        "📥 下载志愿者名单",
                        f, file_name=f'{safe_name}志愿者名单.docx',
                        mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                    )

            with col2:
                st.metric("申报表", f"{safe_name}申报表.docx")
                with open(declaration_path, 'rb') as f:
                    st.download_button(
                        "📥 下载申报表",
                        f, file_name=f'{safe_name}申报表.docx',
                        mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                    )

            with col3:
                st.metric("活动总结", f"{safe_name}活动总结.docx")
                with open(summary_path, 'rb') as f:
                    st.download_button(
                        "📥 下载活动总结",
                        f, file_name=f'{safe_name}活动总结.docx',
                        mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                    )

            st.success(f"📁 文件已保存到: {output_dir}")

        except Exception as e:
            st.error(f"生成失败: {str(e)}")
            import traceback
            st.code(traceback.format_exc())
        finally:
            # 清理临时文件
            os.unlink(xlsx_path)

# 页脚
st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; color: gray;'>
        志愿活动材料自动撰写脚本 | 重庆交通大学信息学院学生会
    </div>
    """,
    unsafe_allow_html=True
)
