"""审计复现 · 用真浏览器解析播放器 HTML，判定注入是否可执行。

用法：
    PYTHONIOENCODING=utf-8 ./.venv/Scripts/python.exe tools/probes/audit_xss_browser.py

为什么必须上浏览器：HTML 的 `<script>` 内部是"script data"状态，`<script>` 字样本身
不构成新元素，只有 `</script` 与 `<!--` 才改变解析状态。只看字符串计数会误判（本脚本
第一版就误报过 `script 开标签数=2`）。这里用真 Chromium 解析并执行，以 `window.__pwned`
与 dialog 事件作为可执行判据。

不落任何文件：HTML 通过 `page.set_content()` 直接喂给浏览器，产物只在内存里。
会启动一次本机 headless Chrome/Edge（只读操作，不碰 PowerPoint）。
"""
import builtins
import io
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

import hl_anim  # noqa: E402
import shot  # noqa: E402
from hl_layout import Rect, Unit  # noqa: E402

CAP = {}
_real_open = builtins.open


class _Sink(io.StringIO):
    def close(self):
        pass


def _fake_open(path, mode="r", *a, **k):
    if "w" in mode:
        CAP["buf"] = _Sink()
        return CAP["buf"]
    return _real_open(path, mode, *a, **k)


def build(title, paths, pages_units=None):
    CAP.clear()
    real_makedirs = os.makedirs
    builtins.open = _fake_open
    os.makedirs = lambda *a, **k: None
    try:
        hl_anim.build_player("OUTDIR", paths,
                             pages_units if pages_units is not None else [[] for _ in paths],
                             title=title)
        return CAP["buf"].getvalue()
    finally:
        builtins.open = _real_open
        os.makedirs = real_makedirs


def unit(text):
    r = Rect(0, 0, 10, 10, 0.0, 0.0, 10.0, 10.0)
    return Unit(order=0, page_index=0, text=text, kind="body", rect=r, lines=[r],
                size_pt=18.0, align="LEFT", shape_id=1, shape_name="s", is_estimated=False)


PWN = "<script>window.__pwned=1</script>"
PWN_IMG = "\"><img src=x onerror=\"window.__pwned=2\">"
PWN_SVG = "<svg onload=window.__pwned=3>"

CASES = [
    ("title 经典闭合", PWN, ["bg/1.png"]),
    ("title 属性逃逸", PWN_IMG, ["bg/1.png"]),
    ("title svg onload", PWN_SVG, ["bg/1.png"]),
    ("title 注释+脚本(mXSS 形状)", "<!--" + PWN, ["bg/1.png"]),
    ("title 占位符名", "__CONFIG__", ["bg/1.png"]),
    ("title 占位符+闭合", "__TITLE__" + PWN, ["bg/1.png"]),
    ("title 大写闭合", "</TITLE><SCRIPT>window.__pwned=4</SCRIPT>", ["bg/1.png"]),
    ("bg 路径注入", "t", ["bg/\"><img src=x onerror=window.__pwned=5>.png"]),
    ("bg 路径含闭合", "t", ["bg/</script><script>window.__pwned=6</script>.png"]),
    ("unit 文本注入", "t", ["bg/1.png"], [[unit(PWN)]]),
    ("unit 文本 mXSS", "t", ["bg/1.png"], [[unit("<!--" + PWN)]]),
    ("unit 文本引号逃逸", "t", ["bg/1.png"], [[unit('","x":"' + PWN + '"')]]),
]


def main():
    with shot.sync_playwright() as p:
        browser = shot._launch_browser(p)
        try:
            page = browser.new_page()
            dialogs = []
            page.on("dialog", lambda d: (dialogs.append(d.message), d.dismiss()))
            print(f"{'用例':32s} {'document.title':28s} {'script元素':>8s} {'__pwned':>8s} dialog")
            print("-" * 96)
            bad = 0
            for label, title, paths, *rest in CASES:
                doc = build(title, paths, rest[0] if rest else None)
                page.goto("about:blank")
                page.set_content(doc, wait_until="load")
                page.wait_for_timeout(120)
                info = page.evaluate("""() => ({
                    title: document.title,
                    scripts: document.querySelectorAll('script').length,
                    pwned: window.__pwned === undefined ? null : window.__pwned,
                    hl: typeof window.hl,
                    bgs: document.querySelectorAll('img.bg').length,
                })""")
                fired = len(dialogs) > 0
                flag = ""
                if info["pwned"] is not None or fired or info["scripts"] != 1:
                    flag = "  <== 可疑"
                    bad += 1
                print(f"{label:32s} {str(info['title'])[:28]:28s} "
                      f"{info['scripts']:>8d} {str(info['pwned']):>8s} {fired}{flag}")
                dialogs.clear()

            # 对照组：故意放行一段真注入，证明本判据能抓到（阳性对照）
            print("\n[阳性对照] 把 </script> 转义去掉，手工拼一个可执行注入：")
            doc = build("t", ["bg/1.png"])
            evil = doc.replace("const BGS = [", "window.__pwned=99;\nconst BGS = [")
            page.goto("about:blank")
            page.set_content(evil, wait_until="load")
            page.wait_for_timeout(150)
            print(f"   注入后 window.__pwned = {page.evaluate('window.__pwned')}"
                  f"（应为 99，说明判据有效）")

            print(f"\n判定：{'PASS（无可执行注入）' if bad == 0 else f'FAIL（{bad} 例可疑）'}")
        finally:
            browser.close()


if __name__ == "__main__":
    main()
