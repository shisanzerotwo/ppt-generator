"""阶段一 M1：设计稿 HTML → 逐页 PNG（1280×720，16:9）。

Playwright 复用本机 Chrome/Edge（channel 启动），免下载浏览器内核：
逐页 scrollIntoView({behavior:"instant"}) 定位 .slide section → viewport 截图。
等待策略：networkidle + document.fonts.ready + 固定缓冲——防止 LLM 生成的
入场动画/过渡（scroll-behavior:smooth 等）截到半渲染状态。
"""

import os
from pathlib import Path

from playwright.sync_api import sync_playwright

W, H = 1280, 720  # 16:9，与设计稿 print CSS 的 @page 尺寸一致
_SETTLE_MS = 300  # 截图前固定缓冲：留入场动画/图片解码时间


def _launch_browser(p):
    """优先系统 Chrome，其次 Edge（同为 Chromium），都没有给明确安装指引。"""
    last_err = None
    for channel in ("chrome", "msedge"):
        try:
            return p.chromium.launch(channel=channel, headless=True)
        except Exception as e:  # 该 channel 浏览器未安装
            last_err = e
    raise RuntimeError(
        "未找到可用的 Chrome/Edge。请安装 Google Chrome，或运行 "
        "`playwright install chromium` 下载内置内核"
    ) from last_err


def shot_deck(html_path: str, out_dir: str) -> list[str]:
    """截取设计稿每一页，返回图片路径列表（slide_1.png 起，与页序一致）。"""
    os.makedirs(out_dir, exist_ok=True)
    # as_uri() 自动百分号编码：稿名含 #/% 时裸拼 file:/// 会被截断（审计 M1）
    url = Path(os.path.abspath(html_path)).as_uri()
    shots: list[str] = []
    with sync_playwright() as p:
        browser = _launch_browser(p)
        try:
            page = browser.new_page(viewport={"width": W, "height": H})
            page.goto(url, wait_until="networkidle")
            page.evaluate("document.fonts.ready.then(() => true)")
            page.wait_for_timeout(_SETTLE_MS)

            n = page.evaluate("document.querySelectorAll('.slide').length")
            if n == 0:
                raise ValueError("设计稿中未找到 .slide 页面（HTML 结构异常）")

            for i in range(n):
                # behavior:"instant" 无视 CSS scroll-behavior:smooth，避免截到滚动中途
                page.evaluate(
                    "document.querySelectorAll('.slide')[%d]."
                    "scrollIntoView({behavior: 'instant', block: 'start'})" % i
                )
                page.wait_for_timeout(_SETTLE_MS)
                path = os.path.join(out_dir, f"slide_{i + 1}.png")
                page.screenshot(path=path)
                shots.append(path)
        finally:
            browser.close()
    return shots
