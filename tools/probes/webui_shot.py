"""WebUI 截图探针：起真实端口 + Playwright 拍工作台三个状态。

用法：python tools/probes/webui_shot.py <输出目录> <前缀>
产物：<输出目录>/<前缀>_home.png / _deck.png / _modal.png
验证口径：退出码非 0 = 页面 JS 报错或截图失败（pageerror 会先落 stderr）。
"""
import os
import sys
import threading

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from playwright.sync_api import sync_playwright
from werkzeug.serving import make_server

import app as app_mod
import shot as shot_mod


def main():
    out_dir, prefix = sys.argv[1], sys.argv[2]
    server = make_server("127.0.0.1", 0, app_mod.app)
    port = server.server_port
    threading.Thread(target=server.serve_forever, daemon=True).start()

    with sync_playwright() as p:
        browser = shot_mod._launch_browser(p)
        page = browser.new_page(viewport={"width": 1600, "height": 1000})
        errors = []
        page.on("pageerror", lambda e: errors.append(str(e)))

        page.goto(f"http://127.0.0.1:{port}/", wait_until="load")
        page.wait_for_timeout(1200)
        page.screenshot(path=f"{out_dir}/{prefix}_home.png")

        # 展开历史项目并载入第一个，拍「有内容」的编辑区状态
        page.evaluate("document.getElementById('project-list').closest('details').open = true")
        page.wait_for_timeout(300)
        items = page.query_selector_all("#project-list button[data-project]")
        if items:
            items[0].click()
            page.wait_for_timeout(1500)
        page.screenshot(path=f"{out_dir}/{prefix}_deck.png")

        # 打开模型与渠道对话框
        page.click("#qb-trigger")
        page.wait_for_timeout(800)
        page.screenshot(path=f"{out_dir}/{prefix}_modal.png")
        page.keyboard.press("Escape")
        page.wait_for_timeout(300)

        # 「新建演示」应回首屏：载入项目后 hero 被隐藏，点击后必须恢复可见
        # （原版按钮无 id，用侧栏主按钮类定位，兼容原版与重设计版）
        if items:
            page.click("#sidebar .side-item.primary")
            page.wait_for_timeout(500)
            if page.locator("#hero").is_hidden():
                print("PAGEERROR: new-deck clicked but #hero still hidden", file=sys.stderr)
                return 1

        browser.close()
    if errors:
        print("PAGEERROR:", errors, file=sys.stderr)
        return 1
    print("OK: 3 shots ->", out_dir)
    return 0


if __name__ == "__main__":
    sys.exit(main())
