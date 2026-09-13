"""阶段一 M1：设计稿 HTML → 逐页 PNG（1280×720，16:9）。

Playwright 复用本机 Chrome/Edge（channel 启动），免下载浏览器内核：
逐页 scrollIntoView({behavior:"instant"}) 定位 .slide section → viewport 截图。

等待/确定化策略（OmniRoute vision 看图实测抓到的缺陷驱动）：
LLM 设计稿的入场与图表动效几乎都由 IntersectionObserver 触发，而无头截图上下文里
IO 回调**不可靠**——实测等 3s 也不触发，柱高永远是 0，于是空图表被截进图片和视频。
单纯加长缓冲治不了，故改为三重确定化：
  ① 注入 IO shim（`_IO_SHIM_JS`）：回调在加载后立即以 isIntersecting=true 派发；
  ② `screenshot(animations="disabled")`：CSS 动画/过渡快进到终态；
  ③ 稳定性轮询兜底：连续两帧字节一致才算稳，并设上限防无限动画卡死导出。
"""

import os
from pathlib import Path

from playwright.sync_api import sync_playwright

W, H = 1280, 720  # 16:9，与设计稿 print CSS 的 @page 尺寸一致
_MIN_SETTLE_MS = 300   # 最小缓冲：留首帧渲染/图片解码
_STABLE_POLL_MS = 200  # 稳定性轮询间隔
_MAX_SETTLE_MS = 2000  # 单页最长等待：到点用最后一帧，不阻塞出图

# 立即派发 isIntersecting=true，让「滚动才出现」的图表/入场动效在截图前就开始跑。
# 只实现截图需要的最小面：observe 即回调一次；unobserve/disconnect 空实现（不影响最终画面）。
_IO_SHIM_JS = """
(() => {
  window.IntersectionObserver = class {
    constructor(cb) { this._cb = cb; }
    observe(el) {
      setTimeout(() => this._cb(
        [{ isIntersecting: true, intersectionRatio: 1, target: el }], this), 0);
    }
    unobserve() {}
    disconnect() {}
    takeRecords() { return []; }
  };
})();
"""


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


def _screenshot_settled(page, min_ms: int, max_ms: int) -> bytes:
    """等到画面不再变化再返回截图：连续两帧字节一致即认为动画停了。"""
    page.wait_for_timeout(min_ms)
    prev = page.screenshot(animations="disabled")
    waited = min_ms
    while waited < max_ms:
        page.wait_for_timeout(_STABLE_POLL_MS)
        waited += _STABLE_POLL_MS
        cur = page.screenshot(animations="disabled")
        if cur == prev:
            break
        prev = cur
    return prev


def shot_deck(html_path: str, out_dir: str,
              min_settle_ms: int = _MIN_SETTLE_MS,
              max_settle_ms: int = _MAX_SETTLE_MS) -> list[str]:
    """截取设计稿每一页，返回图片路径列表（slide_1.png 起，与页序一致）。

    min_settle_ms / max_settle_ms 传 0 可关掉稳定性轮询（压测/调试用）。
    """
    os.makedirs(out_dir, exist_ok=True)
    # as_uri() 自动百分号编码：稿名含 #/% 时裸拼 file:/// 会被截断（审计 M1）
    url = Path(os.path.abspath(html_path)).as_uri()
    shots: list[str] = []
    with sync_playwright() as p:
        browser = _launch_browser(p)
        try:
            page = browser.new_page(viewport={"width": W, "height": H})
            page.add_init_script(_IO_SHIM_JS)  # 必须在 goto 前，抢在稿件脚本之前改写 IO
            page.goto(url, wait_until="networkidle")
            page.evaluate("document.fonts.ready.then(() => true)")

            n = page.evaluate("document.querySelectorAll('.slide').length")
            if n == 0:
                raise ValueError("设计稿中未找到 .slide 页面（HTML 结构异常）")

            for i in range(n):
                # behavior:"instant" 无视 CSS scroll-behavior:smooth，避免截到滚动中途
                page.evaluate(
                    "document.querySelectorAll('.slide')[%d]."
                    "scrollIntoView({behavior: 'instant', block: 'start'})" % i
                )
                png = _screenshot_settled(page, min_settle_ms, max_settle_ms)
                path = os.path.join(out_dir, f"slide_{i + 1}.png")
                with open(path, "wb") as f:
                    f.write(png)
                shots.append(path)
        finally:
            browser.close()
    return shots
