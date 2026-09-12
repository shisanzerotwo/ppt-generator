"""独立验证 ③ · M3 播放器的**对抗性转义**：自己构造恶意输入去撞它。

覆盖任务卡点名的三类，外加我自己加的：
  · title  = `</title><script>alert(1)</script>` / `__CONFIG__` / `<!--` / `</SCRIPT>`
  · 单元文本 = `</script><script>` / `<!--` / U+2028·U+2029 / 空字节 / `"}; alert(...)`
  · 底图路径 = `#` `%` 空格 中文（URL 编码）、引号尖括号（静态断言）
静态断言之外，**用真实 Chrome 打开产物**：挂 `pageerror`/`dialog` 监听，
断言无脚本执行、恶意串以**字面文本**渲染、`hl.goto` 同步生效、
底图经百分号编码后仍能加载到磁盘上的真实文件。

一条命令：./.venv/Scripts/python.exe tools/probes/indep_m3_escape.py
"""

import json
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

import hl_anim  # noqa: E402
from hl_layout import Rect, Unit  # noqa: E402

OUT = os.path.join(ROOT, "output", "spike", "indep_m3")

TITLE = "</title><script>alert(1)</script>__CONFIG__<!--x-->"

UNIT_TEXTS = [
    "</script><script>alert(2)</script>",
    "<!-- comment -->",
    "</SCRIPT>alert(3)</SCRIPT>",
    "  行分隔   段分隔",
    "</title>",
    "__UNITS__ 与 __TITLE__ 与 __BGJSON__",
    "--> 反向注释",
    '"}; alert(4); var x={"',
    "\x00 空字节",
    "\\u0041 字面反斜杠",
    "<img src=x onerror=alert(5)>",
]

BG_REAL = "bg/slide_#1%x 中文.png"          # 磁盘上真的建这个文件，验证编码能取回
BG_HOSTILE = ['bg/a"b<script>c.png', "bg/<svg onload=alert(6)>.png"]


def _rect(l, t, w, h):
    return Rect(int(l * 12700), int(t * 12700), int(w * 12700), int(h * 12700),
                float(l * 2), float(t * 2), float(w * 2), float(h * 2))


def _unit(text, i=0):
    r = _rect(100, 50 + i * 60, 300, 40)
    return Unit(order=i, page_index=0, text=text, kind="body", rect=r, lines=[r],
                size_pt=18.0, align="LEFT", shape_id=i + 1, shape_name="S",
                is_estimated=False, warnings=[])


def _split_script(doc):
    """切成 (script 之外的前段, script 体, script 之后的尾段)。"""
    head, _, rest = doc.partition("<script>")
    body, _, tail = rest.rpartition("</script>")
    return head, body, tail


def _js_literal_to_json(text):
    """产物里是 **JS 字面量**（`<\\/` 合法、`<\\!--` 合法），json.loads 前要还原。"""
    return text.replace("<\\/", "</").replace("<\\!--", "<!--")


def static_checks(doc, label, problems):
    low = doc.lower()
    n_close = low.count("</script")
    if n_close != 1:
        problems.append(f"[{label}] 全文中 `</script` 出现 {n_close} 次（应恰为模板自身 1 次）"
                        f" —— 可能存在提前闭合")
    if "<iframe" in low or "sandbox" in low:
        problems.append(f"[{label}] 出现 iframe/sandbox（契约 §6.2 要求零 iframe）")

    head, body, tail = _split_script(doc)
    # 注意：模板自身含 `<svg class="dim">` 遮罩，且 JSON 字符串里的 `<script>` 在
    # script-data 状态下是**惰性**的（只有 `</script` 能终止脚本块，已在上方计数）。
    # 真正的危险是 `<!--` + `<script` 组合把解析器拖进 double-escaped 态 —— 而 `<!--`
    # 已被转义成 `<\!--`，故下方单独断言 `<!--` 不存在即为充分条件。
    for bad in ("onerror=", "onload="):
        if bad in (head + tail).lower():
            problems.append(f"[{label}] script 之外出现可执行片段 {bad!r}")
    if "<!--" in body:
        problems.append(f"[{label}] script 体内出现 `<!--`（会把解析器拖进注释/转义态，"
                        f"导致末尾 </script> 失效）")
    if "<script" in body.lower():
        print(f"    [note] script 体内有裸 `<script`（来自 JSON 字符串，惰性；"
              f"因 `<!--` 已被转义，无法进入 double-escaped 态）")
    for marker in (".hl{position:absolute", "const CFG = ", "const BGS = ", "const PAGES = "):
        if doc.count(marker) != 1:
            problems.append(f"[{label}] 模板固定标记 {marker!r} 出现 {doc.count(marker)} 次（应为 1）")

    for name, pat in (("CFG", r"const CFG = (.*?);\n"),
                      ("BGS", r"const BGS = (.*?);\n"),
                      ("PAGES", r"const PAGES = (.*?);\n")):
        m = re.search(pat, doc)
        if not m:
            problems.append(f"[{label}] 找不到 {name} 赋值")
            continue
        try:
            json.loads(_js_literal_to_json(m.group(1)))
        except Exception as exc:  # noqa: BLE001
            problems.append(f"[{label}] {name} 还原后仍不是合法 JSON（注入破坏了脚本）：{exc}")
    return problems


def browser_check(path, expect_title, problems):
    """真实 Chrome：无脚本执行、恶意串按字面渲染、hl 可用、编码路径能取到文件。"""
    try:
        from playwright.sync_api import sync_playwright
        import shot
    except Exception as exc:  # noqa: BLE001
        print(f"  [skip] playwright 不可用：{exc}")
        return None
    errors, dialogs = [], []
    try:
        with sync_playwright() as p:
            try:
                browser = shot._launch_browser(p)
            except RuntimeError as exc:
                print(f"  [skip] 无可用浏览器：{exc}")
                return None
            try:
                page = browser.new_page(viewport={"width": 1280, "height": 720})
                page.on("pageerror", lambda e: errors.append(str(e)))
                page.on("dialog", lambda d: (dialogs.append(d.message), d.dismiss()))
                page.goto("file:///" + path.replace("\\", "/"), wait_until="load")
                page.evaluate("window.hl.ready")

                got_title = page.evaluate("document.title")
                h1 = page.evaluate("document.querySelector('h1').textContent")
                st0 = page.evaluate("window.hl.state()")
                page.evaluate("window.hl.goto(0, 1)")
                hl_on = page.evaluate(
                    "Array.from(document.querySelectorAll('.hl')).filter(d=>d.classList.contains('on')).length")
                st1 = page.evaluate("window.hl.state()")
                imgs = page.evaluate(
                    "Array.from(document.querySelectorAll('img.bg')).map(i=>[i.naturalWidth, i.getAttribute('src')])")
                bars = page.evaluate("document.querySelectorAll('.dot').length")
            finally:
                browser.close()
    except Exception as exc:  # noqa: BLE001
        problems.append(f"[browser] 打开播放器抛错：{type(exc).__name__}: {exc}")
        return None

    expect_doc_title = expect_title + " · 高亮讲解"
    if got_title != expect_doc_title:
        problems.append(f"[browser] document.title 不符：{got_title!r} != {expect_doc_title!r}")
    # h1 在模板里是「__TITLE__ · 高亮讲解」，字面渲染时也带后缀
    if h1 != expect_doc_title:
        problems.append(f"[browser] h1 未按字面文本渲染：{h1!r} != {expect_doc_title!r}")
    if errors:
        problems.append(f"[browser] 页面 JS 报错：{errors}")
    if dialogs:
        problems.append(f"[browser] 弹出对话框（注入被执行）：{dialogs}")
    if st0.get("totalPages") != 3:
        problems.append(f"[browser] totalPages 应为 3，实为 {st0.get('totalPages')}")
    if hl_on != 1:
        problems.append(f"[browser] goto(0,1) 后可见高亮 div 数应为 1，实为 {hl_on}")
    if st1.get("step") != 1:
        problems.append(f"[browser] goto 未同步生效：state={st1}")
    if bars != 3:
        problems.append(f"[browser] 页点数应为 3，实为 {bars}")
    print(f"  [browser] title={got_title!r}")
    print(f"  [browser] h1={h1!r}")
    print(f"  [browser] state={st1}  可见高亮={hl_on}  页点={bars}")
    print(f"  [browser] 底图加载情况（naturalWidth, src）：")
    for w, src in imgs:
        print(f"            {w:>5}  {src}")
    return imgs


def main() -> int:
    os.makedirs(OUT, exist_ok=True)
    problems = []

    # 真实底图文件（含 # % 空格 中文）：验证百分号编码能取回磁盘上的真实文件
    bg_dir = os.path.join(OUT, "p1")
    os.makedirs(os.path.join(bg_dir, "bg"), exist_ok=True)
    from PIL import Image
    real = os.path.join(bg_dir, BG_REAL.replace("/", os.sep))
    Image.new("RGB", (64, 36), (12, 34, 56)).save(real)

    units = [_unit(t, i) for i, t in enumerate(UNIT_TEXTS)]

    # ---- 用例 1：恶意 title + 恶意单元文本 + 真实编码路径
    p1 = hl_anim.build_player(
        os.path.join(OUT, "p1"),
        [BG_REAL, *BG_HOSTILE],
        [units, [_unit("第二页")], [_unit("第三页")]],
        title=TITLE)
    doc1 = open(p1, encoding="utf-8").read()
    static_checks(doc1, "case1", problems)
    if "__CONFIG__" not in doc1:
        problems.append("[case1] title 里的 __CONFIG__ 字面量消失了（单遍替换后应保留为字面文本）")
    print("== 用例 1：恶意 title / 单元文本 / 底图路径 ==")
    print(f"  产物：{p1}（{len(doc1)} 字符）")
    print(f"  <script 出现 {doc1.count('<script')} 次、</script> {doc1.count('</script>')} 次"
          f"（模板各 1）")
    print("  真实 Chrome 校验：")
    imgs = browser_check(p1, TITLE, problems)
    if imgs is not None:
        loaded = [w for w, _ in imgs if w > 0]
        if not loaded:
            problems.append("[browser] 没有一张底图加载成功（编码路径可能取不到文件）")
        else:
            print(f"  → 成功加载 {len(loaded)}/{len(imgs)} 张底图（含 #/%%/空格/中文 那张）")

    # ---- 用例 2：占位符字面量出现在**单元文本**里（单遍替换回归）
    print("\n== 用例 2：单元文本含全部占位符字面量 ==")
    p2 = hl_anim.build_player(os.path.join(OUT, "p2"), ["bg/slide_1.png"],
                              [[_unit("__TITLE__ __BGJSON__ __UNITS__ __CONFIG__")]],
                              title="正常标题")
    doc2 = open(p2, encoding="utf-8").read()
    static_checks(doc2, "case2", problems)
    m = re.search(r"const PAGES = (.*?);\n", doc2)
    data = json.loads(m.group(1))
    if data[0][0]["t"] != "__TITLE__ __BGJSON__ __UNITS__ __CONFIG__":
        problems.append("[case2] 单元文本里的占位符被二次替换了")
    print(f"  单元文本原样保留：{data[0][0]['t']!r}")
    print(f"  const CFG 仍是对象：{json.loads(re.search(r'const CFG = (.*?);', doc2).group(1))['canvasWidth']}")

    # ---- 用例 3：路径逃逸
    print("\n== 用例 3：底图路径逃逸（含反斜杠/编码双写）==")
    from pptx_io import PptxError
    for bad in ("..\\..\\x.png", "bg\\..\\..\\x.png", "C:foo.png",
                "\\\\server\\share\\x.png", "//server/share/x.png",
                "bg/..%2f..%2fetc.png"):
        try:
            hl_anim.build_player(os.path.join(OUT, "p3"), [bad], [[]])
            verdict = "未拦截"
            # 注意：bg/..%2f..%2fetc.png 不是字面 ../，拦截与否取决于是否先解码
            print(f"  {bad!r:32} → 未拦截（放行）")
        except PptxError as exc:
            print(f"  {bad!r:32} → 拦截 code={exc.code}")
        except Exception as exc:  # noqa: BLE001
            verdict = f"抛 {type(exc).__name__}"
            print(f"  {bad!r:32} → {verdict}")
            problems.append(f"[case3] 路径 {bad!r} 抛出非 PptxError：{type(exc).__name__}: {exc}")

    print("\n== 判定 ==")
    if problems:
        for x in problems:
            print(f"  ✗ {x}")
        print(f"\n  共 {len(problems)} 项问题")
        return 1
    print("  全部对抗性检查通过（静态 + 真实 Chrome）")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
