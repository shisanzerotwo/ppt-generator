"""第三步：python-pptx 按固定版式组装 PPT。"""

import os

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Emu, Inches, Pt

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
FONT = "微软雅黑"
COLOR_TITLE = RGBColor(0x33, 0x33, 0x33)
COLOR_BODY = RGBColor(0x55, 0x55, 0x55)
COLOR_PLACEHOLDER = RGBColor(0xE0, 0xE0, 0xE0)


def _set_text(frame, text, size, color, bold=False, align=PP_ALIGN.LEFT):
    p = frame.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    return p


def _add_cover(slide, title, subtitle):
    box = slide.shapes.add_textbox(Inches(1), Inches(2.6), SLIDE_W - Inches(2), Inches(2))
    tf = box.text_frame
    _set_text(tf, title, 44, COLOR_TITLE, bold=True, align=PP_ALIGN.CENTER)
    if subtitle:
        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        run = p2.add_run()
        run.text = subtitle
        run.font.name = FONT
        run.font.size = Pt(20)
        run.font.color.rgb = COLOR_BODY


def _add_content(slide, title, points, image_path):
    # 左半文字
    text_box = slide.shapes.add_textbox(Inches(0.6), Inches(0.7), Inches(6.2), Inches(6.1))
    tf = text_box.text_frame
    tf.word_wrap = True
    _set_text(tf, title, 30, COLOR_TITLE, bold=True)
    for pt in points:
        p = tf.add_paragraph()
        p.space_before = Pt(14)
        run = p.add_run()
        run.text = f"• {pt}"
        run.font.name = FONT
        run.font.size = Pt(17)
        run.font.color.rgb = COLOR_BODY

    # 右半图片区：等比缩放居中，失败留灰框
    img_area_x, img_area_y = Inches(7.2), Inches(0.9)
    img_area_w, img_area_h = Inches(5.5), Inches(5.7)
    if image_path and os.path.exists(image_path):
        with Image.open(image_path) as im:
            w, h = im.size
        scale = min(img_area_w / w, img_area_h / h)
        disp_w, disp_h = int(w * scale), int(h * scale)
        left = img_area_x + int((img_area_w - disp_w) / 2)
        top = img_area_y + int((img_area_h - disp_h) / 2)
        slide.shapes.add_picture(image_path, left, top, width=Emu(disp_w), height=Emu(disp_h))
    else:
        rect = slide.shapes.add_shape(1, img_area_x, img_area_y, img_area_w, img_area_h)
        rect.fill.solid()
        rect.fill.fore_color.rgb = COLOR_PLACEHOLDER
        rect.line.fill.background()
        _set_text(rect.text_frame, "（图片生成失败）", 12, COLOR_BODY, align=PP_ALIGN.CENTER)


def build_ppt(slides: list[dict], image_paths: list[str | None], out_path: str) -> str:
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    blank = prs.slide_layouts[6]

    for i, s in enumerate(slides):
        slide = prs.slides.add_slide(blank)
        title = s.get("title", "")
        points = s.get("points", [])
        if i == 0 and not points:
            _add_cover(slide, title, s.get("image_prompt", ""))
        else:
            _add_content(slide, title, points, image_paths[i] if i < len(image_paths) else None)

    prs.save(out_path)
    return out_path
