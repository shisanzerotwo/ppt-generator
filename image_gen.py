"""第二步：文生图适配器。默认智谱 cogview-3-flash，换供应商只改此文件。"""

import os

import httpx
from dotenv import load_dotenv
from zhipuai import ZhipuAI

load_dotenv()

PROVIDER = "zhipu"
MODEL = "cogview-3-flash"
SIZE = "1024x1024"


def generate_image(prompt: str, save_path: str) -> str:
    """按提示词生成一张图并保存到 save_path，返回 save_path。失败抛异常。"""
    api_key = os.getenv("ZHIPUAI_API_KEY")
    if not api_key or api_key == "your-key-here":
        raise RuntimeError("未配置 ZHIPUAI_API_KEY")

    client = ZhipuAI(api_key=api_key)
    resp = client.images.generations(model=MODEL, prompt=prompt, size=SIZE)
    url = resp.data[0].url

    with httpx.Client(timeout=60) as http:
        r = http.get(url)
        r.raise_for_status()
        with open(save_path, "wb") as f:
            f.write(r.content)
    return save_path
