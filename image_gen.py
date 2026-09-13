"""第二步：文生图适配器。默认连当前渠道（默认 .env 的 Agnes AI Hub），换供应商在界面「模型设置」加渠道。"""

import base64
import io
import time

import httpx
from PIL import Image

from llm_util import get_model, llm_client

PROVIDER = "agnes"
SIZE = "1024x1024"
# 全篇统一画风后缀，保证视觉一致（ppt-maker 视觉规范）
STYLE_SUFFIX = "，扁平插画风格，柔和高级配色，干净构图"
# 生图是同步长任务：免费上游排队实测可达 200s+，超时给足；重试只兜瞬时错误
IMAGE_TIMEOUT = 360.0
IMAGE_RETRIES = 2

_PNG_MAGIC = b"\x89PNG\r\n\x1a\n"


def _write_png(data: bytes, save_path: str) -> str:
    """按调用方的 .png 契约落盘。供应商可能回 WebP（AI Horde 就是），而 python-pptx
    只认 PNG/JPEG 等少数格式，非 PNG 一律转码，否则导出 PPT 会直接报 unsupported image format。"""
    if data.startswith(_PNG_MAGIC):
        with open(save_path, "wb") as f:
            f.write(data)
        return save_path
    with Image.open(io.BytesIO(data)) as im:
        im.save(save_path, "PNG")
    return save_path


def copy_as_png(src: str, dst: str) -> str:
    """把 src 按 PNG 写到 dst。缓存复用路径用——缓存可能是修复前写入的 WebP。"""
    with open(src, "rb") as f:
        return _write_png(f.read(), dst)


def _download(url: str, save_path: str) -> str:
    """下载图片到 save_path，若不是图片响应则抛异常。"""
    with httpx.Client(timeout=60) as http:
        r = http.get(url)
        r.raise_for_status()
        if "image" not in r.headers.get("content-type", ""):
            raise ValueError(f"非图片响应: {r.headers.get('content-type')}")
        return _write_png(r.content, save_path)


def _save_image(item, save_path: str) -> str:
    """把响应落盘。供应商不同返回不同：有的给 url，有的内联 b64_json，两种都支持。"""
    url = getattr(item, "url", None)
    if url:
        return _download(url, save_path)
    b64 = getattr(item, "b64_json", None)
    if b64:
        return _write_png(base64.b64decode(b64), save_path)
    raise ValueError("生图响应既无 url 也无 b64_json")


def generate_image(prompt: str, save_path: str) -> str:
    """按提示词生成一张图并保存到 save_path，返回 save_path。失败抛异常。"""
    client = llm_client(IMAGE_TIMEOUT)
    model = get_model("image")  # 运行时可切换（界面"模型设置"）

    last_err = None
    for _ in range(IMAGE_RETRIES):
        try:
            resp = client.images.generations(model=model, prompt=prompt + STYLE_SUFFIX, size=SIZE)
            return _save_image(resp.data[0], save_path)
        except Exception as e:
            last_err = e
            time.sleep(2)
    raise RuntimeError(f"生图失败（已重试 {IMAGE_RETRIES - 1} 次）: {last_err}")
