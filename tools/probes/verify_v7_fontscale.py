"""V7 探针：`a:normAutofit/@fontScale` 到底算不算数。

背景：fixture 第 5 页的 AutofitBox 注入了 fontScale=60000，但 COM 底图显示
文字仍按 100% 渲染（6 行 18pt 占 126pt ≈ 22.5pt/行）——说明 PowerPoint 打开时
**重算**了自动缩排。那么真实的 fontScale 语义是什么？

做法（决定性）：
  1. 造一页**必定溢出**的文本框（小框 + 大量文字），normAutofit fontScale=100000；
  2. 用 COM 无窗口打开 → SaveAs 另存 → 逼 PowerPoint 自己算一次缩排；
  3. 从另存文件里读 PowerPoint 写回的 fontScale；
  4. 同时导出该页 PNG，量实际渲染行高，跟 `字号 × fontScale` 对比。
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

OUT = os.path.abspath("output/spike/v7")
EMU_PER_PT = 12700


def make_overflow_deck(path):
    from pptx import Presentation
    from pptx.oxml.ns import qn
    from pptx.util import Inches, Pt

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_textbox(Inches(1.0), Inches(1.0), Inches(5.0), Inches(2.0))
    box.name = "OverflowBox"
    tf = box.text_frame
    tf.word_wrap = True
    for i in range(10):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        r = p.add_run()
        r.text = f"第 {i + 1} 行：这是一段故意写得很长的文字，用来逼 PowerPoint 触发自动缩排。"
        r.font.size = Pt(28)
        r.font.name = "微软雅黑"
    bodyPr = tf._txBody.find(qn("a:bodyPr"))
    for child in list(bodyPr):
        if child.tag in (qn("a:normAutofit"), qn("a:noAutofit"), qn("a:spAutoFit")):
            bodyPr.remove(child)
    bodyPr.append(bodyPr.makeelement(qn("a:normAutofit"), {"fontScale": "100000"}))
    os.makedirs(os.path.dirname(path), exist_ok=True)
    prs.save(path)
    return path


def read_font_scale(path, shape_name="OverflowBox"):
    from pptx import Presentation
    from pptx.oxml.ns import qn
    prs = Presentation(path)
    for sh in prs.slides[0].shapes:
        if sh.name == shape_name:
            bodyPr = sh.text_frame._txBody.find(qn("a:bodyPr"))
            norm = bodyPr.find(qn("a:normAutofit"))
            if norm is None:
                return None, sh
            return norm.get("fontScale"), sh
    return None, None


def main() -> int:
    import pythoncom
    import win32com.client

    os.makedirs(OUT, exist_ok=True)
    src = make_overflow_deck(os.path.join(OUT, "v7_src.pptx"))
    saved = os.path.join(OUT, "v7_saved_by_powerpoint.pptx")
    png = os.path.join(OUT, "v7_slide1.png")
    if os.path.exists(saved):
        os.remove(saved)

    print(f"源文件 fontScale = {read_font_scale(src)[0]}")
    inp = read_font_scale(src)[1]
    tf = inp.text_frame
    inner_h_pt = (inp.height - tf.margin_top - tf.margin_bottom) / EMU_PER_PT
    inner_w_pt = (inp.width - tf.margin_left - tf.margin_right) / EMU_PER_PT
    print(f"框内高 {inner_h_pt:.1f}pt、内宽 {inner_w_pt:.1f}pt；10 段 × 28pt")

    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("PowerPoint.Application")
    pre_count = app.Presentations.Count
    pres = None
    try:
        pres = app.Presentations.Open(src, ReadOnly=False, Untitled=False, WithWindow=False)
        pres.SaveAs(saved)
        pres.Slides(1).Export(png, "PNG", 1920, 1080)
    finally:
        if pres is not None:
            try:
                pres.Close()
            except Exception:
                pass
        if pre_count == 0:
            try:
                app.Quit()
            except Exception:
                pass
        pres = app = None
        pythoncom.CoUninitialize()

    fs, _ = read_font_scale(saved)
    print(f"PowerPoint 另存后的 fontScale = {fs}")
    if fs is None:
        print("→ PowerPoint 把 normAutofit 换成了别的（noAutofit/spAutoFit）")
    else:
        print(f"→ 换算 = {int(fs) / 1000:.1f}%")

    # 量渲染行高
    from PIL import Image
    img = Image.open(png).convert("RGB")
    px = img.load()
    x0, y0 = int(1.0 * 96), int(1.0 * 96)  # Inches→px 按 96dpi 只做定位参考
    # 直接用整幅找墨迹行带（左侧一半）
    def row_has_ink(y, xa, xb):
        for x in range(xa, xb):
            p = px[x, y]
            if p[0] < 160 and p[1] < 160:
                return True
        return False

    rows = []
    y = 0
    while y < img.height:
        if row_has_ink(y, 100, img.width // 2):
            ys = y
            while y < img.height and row_has_ink(y, 100, img.width // 2):
                y += 1
            rows.append((ys, y))
        else:
            y += 1
    print(f"渲染墨迹行带（px，整页左侧半幅）：{rows[:14]}（共 {len(rows)} 段）")

    # 已知：底图 1920px 宽 = 13.333in → 144 px/in → 1pt = 2px
    if len(rows) >= 2:
        pitch = [rows[i + 1][0] - rows[i][0] for i in range(len(rows) - 1)]
        print(f"行间距 pt = {[round(p / 2, 1) for p in pitch]}")
        print(f"若 28pt 未缩放，行高≈28×1.25=35pt；若按 fontScale 缩放，行高≈"
              f"{fs and round(28 * int(fs) / 100000 * 1.25, 1)}pt")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
