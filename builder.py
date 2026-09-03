"""第三步：python-pptx 按主题色板组装 PPT（色板来自 ppt-maker skill）。"""

import os

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Emu, Inches, Pt

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
FONT = "微软雅黑"

# 主题色板来自 ppt-maker skill（~/.agents/skills/ppt-maker）
THEMES = {
    "blue":  {"bg": RGBColor(0x0F, 0x2A, 0x4A), "accent": RGBColor(0x3B, 0x82, 0xF6),
              "fg": RGBColor(0xFF, 0xFF, 0xFF), "muted": RGBColor(0xB6, 0xC7, 0xDC)},
    "dark":  {"bg": RGBColor(0x11, 0x18, 0x27), "accent": RGBColor(0xF5, 0x9E, 0x0B),
              "fg": RGBColor(0xF9, 0xFA, 0xFB), "muted": RGBColor(0x9C, 0xA3, 0xAF)},
    "green": {"bg": RGBColor(0x06, 0x2E, 0x21), "accent": RGBColor(0x22, 0xC5, 0x5E),
              "fg": RGBColor(0xEC, 0xFD, 0xF5), "muted": RGBColor(0xA7, 0xC4, 0xB6)},
}


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


def _add_bg(slide, prs, color):
    rect = slide.shapes.add_shape(1, 0, 0, SLIDE_W, SLIDE_H)
    rect.fill.solid()
    rect.fill.fore_color.rgb = color
    rect.line.fill.background()
    return rect


def _add_cover(slide, title, subtitle, t):
    _set_text(slide.shapes.add_textbox(Inches(1), Inches(2.6), SLIDE_W - Inches(2), Inches(2)).text_frame,
              title, 44, t["fg"], bold=True, align=PP_ALIGN.CENTER)
    if subtitle:
        box = slide.shapes.add_textbox(Inches(1), Inches(3.9), SLIDE_W - Inches(2), Inches(0.8))
        p = _set_text(box.text_frame, subtitle, 18, t["muted"], align=PP_ALIGN.CENTER)
        del p


def _add_content(slide, title, points, image_path, t):
    # 标题左侧强调色竖条（ppt-maker 视觉规范）
    bar = slide.shapes.add_shape(1, Inches(0.6), Inches(0.78), Inches(0.09), Inches(0.52))
    bar.fill.solid()
    bar.fill.fore_color.rgb = t["accent"]
    bar.line.fill.background()

    # 标题（竖条右侧）
    title_box = slide.shapes.add_textbox(Inches(0.85), Inches(0.7), Inches(6.0), Inches(0.9))
    _set_text(title_box.text_frame, title, 28, t["fg"], bold=True)

    # 要点
    text_box = slide.shapes.add_textbox(Inches(0.6), Inches(1.9), Inches(6.2), Inches(4.9))
    tf = text_box.text_frame
    tf.word_wrap = True
    for i, pt in enumerate(points):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_before = Pt(16)
        run = p.add_run()
        run.text = f"▸ {pt}"
        run.font.name = FONT
        run.font.size = Pt(17)
        run.font.color.rgb = t["muted"]

    # 右半图片区：等比缩放居中，失败留半透明占位框
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
        rect.fill.fore_color.rgb = t["muted"]
        rect.line.fill.background()
        _set_text(rect.text_frame, "（图片生成失败）", 12, t["fg"], align=PP_ALIGN.CENTER)


def build_ppt(slides: list[dict], image_paths: list[str | None], out_path: str,
              theme: str = "blue", subtitle: str = "") -> str:
    t = THEMES.get(theme, THEMES["blue"])
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    blank = prs.slide_layouts[6]

    for i, s in enumerate(slides):
        slide = prs.slides.add_slide(blank)
        title = s.get("title", "")
        points = s.get("points", [])
        _add_bg(slide, prs, t["bg"])  # 全页主题深色底
        if i == 0 and not points:
            _add_cover(slide, title, subtitle or s.get("image_prompt", ""), t)
        else:
            _add_content(slide, title, points, image_paths[i] if i < len(image_paths) else None, t)

    prs.save(out_path)
    return out_path
