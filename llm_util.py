"""LLM 调用公共件：client 构造与模型常量（critic / template / html_gen 共用）。

此前 critic.py 与 template.py 各自复制了一份 _client()，超时等参数已经漂移
（60s vs 120s）。收敛到单一来源，各调用方按需传超时——设计任务的 300s 长超时
与单轮问答的 60s 短超时都是合理差异，但实现只能有一份。
"""
import os

from dotenv import load_dotenv
from zhipuai import ZhipuAI

load_dotenv()

VISION_MODEL = os.getenv("ZHIPUAI_VISION_MODEL") or "agnes-2.5-pro"
TEXT_MODEL = os.getenv("ZHIPUAI_CHAT_MODEL") or "agnes-2.0-flash"


def llm_client(timeout: float = 60.0) -> ZhipuAI:
    """构造 ZhipuAI 兼容客户端；未配 key 直接抛错（快速失败优于静默降级）。"""
    api_key = os.getenv("ZHIPUAI_API_KEY")
    if not api_key or api_key == "your-key-here":
        raise RuntimeError("未配置 ZHIPUAI_API_KEY")
    return ZhipuAI(api_key=api_key, base_url=os.getenv("ZHIPUAI_BASE_URL") or None, timeout=timeout)
