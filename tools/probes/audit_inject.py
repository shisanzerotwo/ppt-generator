"""审计复现 · 播放器 HTML 注入 + 底图路径穿越（只读，不落任何产物）。

用法：
    PYTHONIOENCODING=utf-8 ./.venv/Scripts/python.exe tools/probes/audit_inject.py

要点：`build_player` 会真的写 <out_dir>/index.html，所以这里把 `open`/`os.makedirs`
拦到内存里——既跑的是真代码路径，又不在磁盘上留任何文件。
"""
import builtins
import io
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

import hl_anim  # noqa: E402

CAP = {}

_real_open = builtins.open


class _Sink(io.StringIO):
    """像文件一样接收写入，但 with 退出时不丢内容。"""

    def close(self):        # noqa: D102
        pass

    def getvalue(self):     # noqa: D102
        return super().getvalue()


def _fake_open(path, mode="r", *a, **k):
    if "w" in mode:
        CAP["path"] = path
        CAP["buf"] = _Sink()
        return CAP["buf"]
    return _real_open(path, mode, *a, **k)


def build_in_memory(title, paths, pages_units=None):
    """调真 build_player，返回 (doc, path, None)；异常则返回 (None, None, 异常)。"""
    CAP.clear()
    real_makedirs = os.makedirs
    builtins.open = _fake_open
    os.makedirs = lambda *a, **k: None
    try:
        path = hl_anim.build_player(
            "OUTDIR", paths,
            pages_units if pages_units is not None else [[] for _ in paths],
            title=title)
        return CAP["buf"].getvalue(), path, None
    except Exception as exc:                                   # noqa: BLE001
        return None, None, exc
    finally:
        builtins.open = _real_open
        os.makedirs = real_makedirs


def analyze(doc):
    return {
        "script_开标签数": len(re.findall(r"<script", doc, re.I)),
        "script_闭标签数": len(re.findall(r"</script", doc, re.I)),
        "含裸'<!--'": "<!--" in doc,
        "含裸'-->'": "-->" in doc,
        "title_元素数": len(re.findall(r"<title", doc, re.I)),
        "h1_元素数": len(re.findall(r"<h1", doc, re.I)),
    }


def section(t):
    print("\n" + "=" * 78)
    print(t)
    print("=" * 78)


def main():
    section("一 · title 注入反例（模板自带 1 个 <script>/1 个 </script>/1 个 <title>）")
    cases = [
        ("A 经典闭合", "</title><script>alert(1)</script>", ["bg/1.png"]),
        ("B 占位符名", "__CONFIG__", ["bg/1.png"]),
        ("C 占位符+闭合", "__TITLE__<script>x</script>", ["bg/1.png"]),
        ("D 注释注入", "<!--<script>alert(1)</script>-->", ["bg/1.png"]),
        ("E 单独 -->", "-->", ["bg/1.png"]),
        ("F 混合占位符", "__BGJSON__ __UNITS__ __TITLE__", ["bg/1.png"]),
        ("G 零宽闭合", "</scr\tipt>", ["bg/1.png"]),
    ]
    for label, title, paths in cases:
        doc, path, exc = build_in_memory(title, paths)
        if exc is not None:
            print(f"\n[{label}] 抛 {type(exc).__name__}: {exc}")
            continue
        print(f"\n[{label}] title={title!r}")
        print(f"   {analyze(doc)}")
        m = re.search(r"<title>(.*?)</title>", doc, re.S)
        print(f"   <title> 实际内容 = {m.group(1)!r}" if m else "   <title> 未匹配")
        m = re.search(r"<h1>(.*?)</h1>", doc, re.S)
        print(f"   <h1>    实际内容 = {m.group(1)!r}" if m else "   <h1> 未匹配")
        m = re.search(r"const CFG = (.*?);\n", doc, re.S)
        if m:
            print(f"   const CFG = {m.group(1)!r}")
        m = re.search(r"const BGS = (.*?);\n", doc, re.S)
        if m:
            print(f"   const BGS = {m.group(1)!r}")

    section("二 · 单元文本注入反例（走 __UNITS__ / <script> 字符串上下文）")
    from hl_layout import Rect, Unit

    def unit(text):
        r = Rect(0, 0, 10, 10, 0.0, 0.0, 10.0, 10.0)
        return Unit(order=0, page_index=0, text=text, kind="body", rect=r, lines=[r],
                    size_pt=18.0, align="LEFT", shape_id=1, shape_name="s", is_estimated=False)

    for text in ["</script><script>alert(1)</script>", "<!--x", "a-->b", "\u2028\u2029"]:
        doc, path, exc = build_in_memory("t", ["bg/1.png"], [[unit(text)]])
        if exc is not None:
            print(f"\n[unit={text!r}] 抛 {type(exc).__name__}: {exc}")
            continue
        print(f"\n[unit={text!r}]")
        print(f"   {analyze(doc)}")

    section("三 · 底图路径：穿越 / 绝对路径 / 编码")
    path_cases = [
        "../outside.png", "bg/../../x.png", "..\\..\\x.png", "C:/x.png",
        "C:\\x.png", "//srv/share/x.png", "\\\\srv\\share\\x.png",
        "/etc/passwd", "/x.png", "bg/1.png", "bg/slide_#1%2.png",
        "bg/空格 名.png", "bg/</script><script>alert(1)</script>.png",
        "bg/....//x.png", "sub/../bg/1.png", "D:relative.png", "bg/a:b.png",
    ]
    for p in path_cases:
        doc, path, exc = build_in_memory("t", [p])
        if exc is not None:
            print(f"  {p!r:50s} -> 拒绝 {type(exc).__name__}/{getattr(exc, 'code', '')}")
            continue
        m = re.search(r"const BGS = (.*?);\n", doc, re.S)
        raw = m.group(1) if m else "?"
        # 解出 JS 数组里的字符串（够用即可：抓引号内内容）
        vals = re.findall(r'"((?:[^"\\]|\\.)*)"', raw)
        print(f"  {p!r:50s} -> 放行 src={vals[0] if vals else raw!r}")
        if p == "bg/</script><script>alert(1)</script>.png":
            print(f"      文档里是否出现裸 '</script><script'：{'</script><script' in doc}")


if __name__ == "__main__":
    main()
