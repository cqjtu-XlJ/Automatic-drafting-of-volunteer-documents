# -*- coding: utf-8 -*-
"""
LLM 内容生成模块
支持多个 AI 提供商，默认使用智谱AI GLM-4-Flash（免费）
"""

import os
import yaml
import logging
from openai import OpenAI

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# 默认配置文件路径
CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'config.yaml')


def load_config():
    """加载配置文件"""
    if not os.path.exists(CONFIG_FILE):
        logger.warning(f"配置文件不存在: {CONFIG_FILE}")
        return None

    with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
        config = yaml.safe_load(f)

    return config


def get_provider_config(config, provider=None):
    """获取指定提供商的配置"""
    if provider is None:
        provider = config['llm'].get('provider', 'zhipu')

    provider_config = config['llm'].get(provider, {})
    return provider_config, provider


class LLMGenerator:
    """LLM 内容生成器"""

    def __init__(self, provider=None, api_key=None):
        self.config = load_config()
        self.provider = provider
        self.api_key = api_key
        self.client = None
        self.provider_config = None
        self._init_client()

    def _init_client(self):
        """初始化 LLM 客户端"""
        if not self.config or not self.config['llm'].get('enabled', False):
            logger.info("LLM 未启用")
            return

        self.provider_config, self.provider = get_provider_config(self.config, self.provider)

        # 优先使用传入的 API Key，否则从配置文件读取
        api_key = self.api_key or self.provider_config.get('api_key', '')
        if not api_key:
            logger.warning(f"未配置 {self.provider} 的 API Key")
            return

        base_url = self.provider_config.get('base_url', '')
        self.client = OpenAI(
            api_key=api_key,
            base_url=base_url
        )
        logger.info(f"LLM 客户端初始化成功: {self.provider}")

    def is_available(self):
        """检查 LLM 是否可用"""
        return self.client is not None

    def generate(self, prompt, max_tokens=None):
        """调用 LLM 生成内容"""
        if not self.is_available():
            raise RuntimeError("LLM 客户端未初始化")

        if max_tokens is None:
            max_tokens = self.provider_config.get('max_tokens', 1500)

        model = self.provider_config.get('model', 'glm-4-flash')
        temperature = self.provider_config.get('temperature', 0.7)
        timeout = self.config['llm'].get('timeout', 30)

        logger.info(f"调用 {self.provider} API，模型: {model}")

        response = self.client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": prompt}],
            max_tokens=max_tokens,
            temperature=temperature,
            timeout=timeout
        )

        content = response.choices[0].message.content
        logger.info(f"LLM 生成成功，消耗 tokens: {response.usage.total_tokens}")

        return content


def build_activity_content_prompt(activity_name, activity_type, volunteer_count, total_hours, service_contents):
    """构建活动内容生成的 Prompt"""

    prompt = f"""你是一个专业的大学志愿服务活动材料撰写助手。请根据以下信息，生成志愿活动的详细描述内容。

【活动信息】
- 活动名称：{activity_name}
- 活动类型：{activity_type}
- 参与志愿者人数：{volunteer_count}人
- 总服务时长：{total_hours}小时
- 主要服务内容：{service_contents}

【生成要求】
请严格按照以下格式生成4个部分的内容，每部分80-120字：

【内容概述】
简述活动的主要内容和目标，突出大学生志愿服务的特点。

【社会问题】
从大学校园角度分析相关社会问题，例如大学生课余生活单一、缺乏社会实践、校园文化建设需要更多志愿精神等。

【产生原因】
从大学生群体特点分析问题产生的原因，例如学业压力大、缺乏组织化的志愿服务平台、社会实践机会不足等。

【必要性】
说明开展该活动的必要性，必须包含7个"有利于"的排比句，从增强学生社会责任感、丰富课余生活、提升综合素质、促进校园文明建设、培养团队协作能力、弘扬志愿服务精神、推动校园文化建设等角度展开。

【重要提示】
- 必须严格按照上述格式输出，每个部分以"【标题】"开头
- 不要输出"服务对象"和"预计成效"，这两个由系统自动填写
- 内容要贴合大学校园环境
- 必要性部分必须使用"有利于...；有利于...；有利于..."的排比句式
- 语言正式、符合官方材料风格"""

    return prompt


def parse_llm_response(response_text):
    """解析 LLM 返回的内容"""
    sections = {
        'content_overview': '',  # 内容概述
        'social_problem': '',    # 社会问题
        'cause': '',             # 产生原因
        'necessity': '',         # 必要性
        'beneficiary': '',       # 服务对象
        'result': ''             # 预计成效
    }

    # 解析各个部分
    current_section = None
    lines = response_text.split('\n')

    for line in lines:
        line = line.strip()
        if not line:
            continue

        # 识别章节标题并提取内容
        if '【内容概述】' in line:
            current_section = 'content_overview'
            # 提取标题后的内容
            content_after = line.split('】')[-1].strip()
            if content_after:
                sections[current_section] = content_after
            continue
        elif '【社会问题】' in line:
            current_section = 'social_problem'
            content_after = line.split('】')[-1].strip()
            if content_after:
                sections[current_section] = content_after
            continue
        elif '【产生原因】' in line:
            current_section = 'cause'
            content_after = line.split('】')[-1].strip()
            if content_after:
                sections[current_section] = content_after
            continue
        elif '【必要性】' in line:
            current_section = 'necessity'
            content_after = line.split('】')[-1].strip()
            if content_after:
                sections[current_section] = content_after
            continue
        elif '【服务对象】' in line:
            current_section = 'beneficiary'
            content_after = line.split('】')[-1].strip()
            if content_after:
                sections[current_section] = content_after
            continue
        elif '【预计成效】' in line or '【预计效果】' in line:
            current_section = 'result'
            content_after = line.split('】')[-1].strip()
            if content_after:
                sections[current_section] = content_after
            continue

        # 添加内容到当前章节（去除可能的【标题】标记）
        if current_section:
            # 清理行中的【标题】标记
            cleaned_line = line
            for marker in ['【内容概述】', '【社会问题】', '【产生原因】', '【必要性】', '【服务对象】', '【预计成效】', '【预计效果】']:
                cleaned_line = cleaned_line.replace(marker, '')

            if sections[current_section]:
                sections[current_section] += '\n' + cleaned_line
            else:
                sections[current_section] = cleaned_line

    # 清理内容（去除首尾空白）
    for key in sections:
        sections[key] = sections[key].strip()

    return sections


def generate_activity_content_with_llm(activity_name, activity_type, volunteer_count, total_hours, service_contents, provider=None, api_key=None):
    """使用 LLM 生成活动内容（带降级处理）"""
    try:
        generator = LLMGenerator(provider=provider, api_key=api_key)

        if not generator.is_available():
            logger.warning("LLM 不可用，将使用静态模板")
            return None

        # 构建 Prompt
        prompt = build_activity_content_prompt(
            activity_name, activity_type, volunteer_count, total_hours, service_contents
        )

        # 调用 LLM
        response = generator.generate(prompt)

        # 解析返回内容
        content = parse_llm_response(response)

        # 验证内容完整性（只检查必需的部分）
        required_keys = ['social_problem', 'cause', 'necessity']
        missing_keys = []
        for key in required_keys:
            if not content.get(key):
                missing_keys.append(key)

        if missing_keys:
            logger.warning(f"LLM 返回内容不完整，缺少: {missing_keys}")
            return None

        # 对于 beneficiary 和 result，如果 LLM 没有生成，使用默认值
        if not content.get('beneficiary'):
            content['beneficiary'] = '校园师生'
        if not content.get('result'):
            content['result'] = f'{activity_name}活动圆满结束。'

        logger.info("LLM 内容生成成功")
        return content

    except Exception as e:
        logger.error(f"LLM 调用失败: {e}")
        return None


# 便捷函数：获取可用的 AI 提供商列表
def get_available_providers():
    """获取可用的 AI 提供商列表（返回配置文件中所有已配置的提供商）"""
    config = load_config()
    if not config:
        return []

    providers = []
    llm_config = config.get('llm', {})

    for provider in ['zhipu', 'deepseek', 'qwen', 'moonshot', 'mimo', 'doubao']:
        provider_config = llm_config.get(provider, {})
        # 只要提供商配置存在（有 base_url），就认为可用
        if provider_config.get('base_url'):
            providers.append(provider)

    return providers


if __name__ == '__main__':
    # 测试代码
    print("测试 LLM 生成模块...")

    generator = LLMGenerator()
    if generator.is_available():
        print(f"LLM 可用，提供商: {generator.provider}")

        # 测试生成
        prompt = "请用一句话介绍志愿服务的意义。"
        try:
            result = generator.generate(prompt, max_tokens=100)
            print(f"生成结果: {result}")
        except Exception as e:
            print(f"生成失败: {e}")
    else:
        print("LLM 不可用，请检查配置文件和 API Key")
