"""M6 验收：导出的 pptx 可编辑、中文字体是**微软雅黑而非宋体**、qa.check_pptx 无 error。

"渲染成微软雅黑"这件事怎么证：
  · **静态**（机制层）：theme1.xml 的 `a:ea` 与 `<a:font script="Hans">` 都是微软雅黑，
    且每个 run 的 `a:ea` 排在 `a:latin` 之后。PowerPoint 选 CJK 字形就是查这两处。
  · **动态**（渲染层）：造一份**对照稿**（同样的文字，但不写 a:ea、也不改主题），
    两份都用 COM 导成 PNG，量整页墨迹密度 —— 微软雅黑是无衬线、笔画均匀，
    宋体是衬线、横细竖粗，同为 1.0em 宽度下**墨迹密度明显不同**。

跑法：.venv/Scripts/python.exe tools/probes/accept_m6.py
"""

import os
import re
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

import pptx_io  # noqa: E402
import pptx_out  # noqa: E402
import qa  # noqa: E402
from pptx.util import Pt  # noqa: E402

SRC = os.path.abspath("output/b_multislide.pptx")
OUT = os.path.abspath("output/spike/m6")
FONT = "微软雅黑"


def _font_facts(path):
    with zipfile.ZipFile(path) as z:
        theme = z.read("ppt/theme/theme1.xml").decode("utf-8")
        ea = re.findall(r'<a:ea typeface="([^"]*)"', theme)
        hans = re.findall(r'<a:font script="Hans" typeface="([^"]*)"', theme)
        runs_ea = 0
        runs_total = 0
        for name in z.namelist():
            if not re.fullmatch(r"ppt/slides/slide\d+\.xml", name):
                continue
            xml = z.read(name).decode("utf-8")
            for rPr in re.findall(r"<a:rPr[^>]*>.*?</a:rPr>|<a:rPr[^>]*/>", xml, re.S):
                runs_total += 1
                if "<a:ea" in rPr:
                    lat = rPr.find("<a:latin")
                    eai = rPr.find("<a:ea")
                    if lat != -1 and eai > lat:
                        runs_ea += 1
        return ea, hans, runs_ea, runs_total


def _ink_density(png):
    from PIL import Image
    with Image.open(png) as im:
        img = im.convert("RGB")
        px = img.load()
        dark = 0
        total = img.width * img.height
        for y in range(0, img.height, 2):
            for x in range(0, img.width, 2):
                p = px[x, y]
                if p[0] < 128 and p[1] < 128 and p[2] < 128:
                    dark += 1
        return dark / (total / 4)


def _export_many(jobs):
    """用**一个** PowerPoint 实例连导多份，返回 png 路径列表。

    一次探针只开关一次实例：`Quit()` 是异步的（进程要过一会儿才退），逐份开关会
    让下一次的安全闸门误判成"用户正开着 PowerPoint"。
    """
    import pythoncom
    import win32com.client
    if pptx_io._powerpoint_running():
        raise RuntimeError("检测到用户正在使用 PowerPoint，本探针中止（安全红线）")
    pythoncom.CoInitialize()
    app = pres = None
    pngs = []
    try:
        app = win32com.client.DispatchEx("PowerPoint.Application")
        for pptx_path, out_dir in jobs:
            os.makedirs(out_dir, exist_ok=True)
            pres = app.Presentations.Open(os.path.abspath(pptx_path), ReadOnly=True,
                                          Untitled=False, WithWindow=False)
            png = os.path.abspath(os.path.join(out_dir, "p1.png"))
            pres.Slides(1).Export(png, "PNG", 1600, 900)
            pngs.append(png)
            pres.Close()
            pres = None
        if app.Presentations.Count == 0:
            app.Quit()
    finally:
        try:
            if pres is not None:
                pres.Close()
        except Exception:
            pass
        app = pres = None
        pythoncom.CoUninitialize()
    return pngs


def main() -> int:
    os.makedirs(OUT, exist_ok=True)
    rc = 0
    deck, _ = pptx_io.read_pages(SRC, export_width_px=1920)
    print(f"素材：{SRC}（{len(deck.pages)} 页）")

    good = os.path.join(OUT, "deck_font_ok.pptx")
    pptx_out.build_deck_pptx(deck, good)
    print(f"\n[1] 导出：{good}（{os.path.getsize(good)} 字节）")

    ea, hans, runs_ea, runs_total = _font_facts(good)
    print(f"[2] 静态字体事实：theme a:ea = {ea}")
    print(f"                  theme Hans = {hans}")
    print(f"                  带 a:ea 且排在 a:latin 之后的 rPr = {runs_ea}/{runs_total}")
    ok_static = (ea and all(v == FONT for v in ea)
                 and hans and all(v == FONT for v in hans)
                 and runs_total > 0 and runs_ea == runs_total)
    print(f"    → {'PASS' if ok_static else 'FAIL'}")
    rc |= 0 if ok_static else 1

    print("\n[3] qa.check_pptx 自检")
    check = qa.check_pptx(good, len(deck.pages))
    print(f"    errors = {check['errors']}")
    print(f"    warnings = {len(check['warnings'])} 条，used_estimate = {check['used_estimate']}")
    print(f"    → {'PASS' if not check['errors'] else 'FAIL'}")
    rc |= 0 if not check["errors"] else 1

    print("\n[4] 可编辑性：改一个字再存回来")
    from pptx import Presentation
    prs = Presentation(good)
    target = None
    for shape in prs.slides[0].shapes:
        if shape.has_text_frame and shape.text_frame.text.strip():
            target = shape
            break
    before = target.text_frame.text
    target.text_frame.paragraphs[0].runs[0].text = "改过的标题"
    edited = os.path.join(OUT, "deck_edited.pptx")
    prs.save(edited)
    after = Presentation(edited).slides[0].shapes[0].text_frame.text
    print(f"    改前 {before!r} → 改后 {after!r}")
    ok_edit = "改过的标题" in after and len(Presentation(edited).slides) == len(deck.pages)
    print(f"    → {'PASS（改字不破版，页数不变）' if ok_edit else 'FAIL'}")
    rc |= 0 if ok_edit else 1

    print("\n[5] 渲染层对照（COM 导底图量墨迹密度）")
    if pptx_io._powerpoint_running():
        print("    [SKIP] 检测到用户 PowerPoint 在运行，按安全红线跳过 COM 对照")
        print(f"\n产物目录：{OUT}")
        print(f"判定：{'PASS' if rc == 0 else 'FAIL'}")
        return rc
    else:
        rc |= _render_compare(deck, good)


def _render_compare(deck, good):
    """三稿对照：ours(声明雅黑) / 显式宋体(阳性对照) / 不声明。

    ⚠️ 为什么要阳性对照：本机实测**不声明字体时也渲染成微软雅黑**（PowerPoint 的
    CJK 回退就是系统默认雅黑），所以"不声明"不能当宋体基线用。先证明"显式声明宋体
    确实能被量出来"，这次度量才有资格对 ours 下结论。
    """
    control = os.path.join(OUT, "deck_font_none.pptx")
    song = os.path.join(OUT, "deck_font_song.pptx")
    orig_set_run = pptx_out._set_run
    orig_patch = pptx_out._patch_theme_zip
    try:
        def _no_font(run, size_pt, bold, builtin):
            run.font.size = Pt(size_pt)
            run.font.bold = bold

        def _song(run, size_pt, bold, builtin):
            run.font.size = Pt(size_pt)
            run.font.bold = bold
            run.font.name = "宋体"
            pptx_out.builder._set_ea(run, "宋体")

        pptx_out._set_run = _no_font
        pptx_out._patch_theme_zip = lambda path, font=FONT: False
        pptx_out.build_deck_pptx(deck, control)

        pptx_out._set_run = _song
        pptx_out._patch_theme_zip = lambda path, font="宋体": False
        pptx_out.build_deck_pptx(deck, song)
    finally:
        pptx_out._set_run = orig_set_run
        pptx_out._patch_theme_zip = orig_patch

    p_ours, p_none, p_song = _export_many([
        (good, os.path.join(OUT, "png_ours")),
        (control, os.path.join(OUT, "png_none")),
        (song, os.path.join(OUT, "png_song")),
    ])
    d_ours, d_none, d_song = (_ink_density(p_ours), _ink_density(p_none),
                              _ink_density(p_song))
    print(f"    墨迹密度：ours(声明雅黑) = {d_ours:.5f}")
    print(f"              不声明       = {d_none:.5f}")
    print(f"              显式宋体     = {d_song:.5f}")
    diff = abs(d_song - d_ours) / max(d_ours, 1e-9)
    ratio = d_ours / max(d_song, 1e-9)
    ok_disc = diff > 0.05
    print(f"    宋体 vs 雅黑 密度差 = {diff * 100:.1f}% → "
          f"{'度量能区分（阳性对照成立）' if ok_disc else '度量区分不出来，本项结论不可用'}")
    print(f"    ours 与「不声明」的差 = "
          f"{abs(d_ours - d_none) / max(d_ours, 1e-9) * 100:.1f}%"
          f"（本机默认回退即雅黑，故两者应接近）")
    print(f"    ours / 宋体 = {ratio:.2f}× → ours 渲染的不是宋体")
    ok = ok_disc and ratio > 1.5
    print(f"    → {'PASS' if ok else 'FAIL'}")
    print("    边界（必须说清）：本机 PowerPoint 的 CJK 回退**本来就是雅黑**，")
    print("    所以「不声明字体」在这里也渲成雅黑 —— 本机无法证明*声明*是必需的。")
    print("    声明的价值在「默认 CJK 字体不是雅黑」的机器上，靠的是静态层证据")
    print("    （theme a:ea / <a:font script=\"Hans\"> / 每个 run 的 a:ea 全为微软雅黑）。")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
