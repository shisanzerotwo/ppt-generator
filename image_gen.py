"""第二步：文生图适配器。默认连 Agnes AI Hub 的 agnes-image-2.5-flash，换供应商只改此文件。"""

import os
import time

import httpx
from dotenv import load_dotenv
from zhipuai import ZhipuAI

from llm_util import get_model

load_dotenv()

PROVIDER = "agnes"
SIZE = "1024x1024"
# 全篇统一画风后缀，保证视觉一致（ppt-maker 视觉规范）
STYLE_SUFFIX = "，扁平插画风格，柔和高级配色，干净构图"


def _download(url: str, save_path: str) -> str:
    """下载图片到 save_path，若不是图片响应则抛异常。"""
    with httpx.Client(timeout=60) as http:
        r = http.get(url)
        r.raise_for_status()
        if "image" not in r.headers.get("content-type", ""):
            raise ValueError(f"非图片响应: {r.headers.get('content-type')}")
        with open(save_path, "wb") as f:
            f.write(r.content)
    return save_path


def generate_image(prompt: str, save_path: str) -> str:
    """按提示词生成一张图并保存到 save_path，返回 save_path。失败抛异常。"""
    api_key = os.getenv("ZHIPUAI_API_KEY")
    if not api_key or api_key == "your-key-here":
        raise RuntimeError("未配置 ZHIPUAI_API_KEY")
    base_url = os.getenv("ZHIPUAI_BASE_URL") or None
    client = ZhipuAI(api_key=api_key, base_url=base_url, timeout=60.0)
    model = get_model("image")  # 运行时可切换（界面"模型设置"）

    last_err = None
    # Agnes 生图为异步：偶发返回任务引用 URL（非图片），重试拿最终图
    for _ in range(4):
        try:
            resp = client.images.generations(model=model, prompt=prompt + STYLE_SUFFIX, size=SIZE)
            return _download(resp.data[0].url, save_path)
        except Exception as e:
            last_err = e
            time.sleep(2)
    raise RuntimeError(f"生图失败（已重试 3 次）: {last_err}")
