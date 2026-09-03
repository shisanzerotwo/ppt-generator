"""第三步：python-pptx 按主题色板 + 多版式组装 PPT（色板来自 ppt-maker skill）。"""

import os

from PIL import Image
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
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


def _set_text(frame, text, size, color, bold=False, align=PP_ALIGN.LEFT, font=FONT):
    p = frame.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    return p


def _add_bg(slide, color):
    rect = slide.shapes.add_shape(1, 0, 0, SLIDE_W, SLIDE_H)
    rect.fill.solid()
    rect.fill.fore_color.rgb = color
    rect.line.fill.background()
    return rect


def _add_title(slide, title, t):
    """标题 + 左侧强调色竖条，返回 True。"""
    bar = slide.shapes.add_shape(1, Inches(0.6), Inches(0.78), Inches(0.09), Inches(0.52))
    bar.fill.solid()
    bar.fill.fore_color.rgb = t["accent"]
    bar.line.fill.background()
    _set_text(slide.shapes.add_textbox(Inches(0.85), Inches(0.7), Inches(6.5), Inches(0.9)).text_frame,
              title, 28, t["fg"], bold=True)


def _add_cover(slide, title, subtitle, t):
    _set_text(slide.shapes.add_textbox(Inches(1), Inches(2.6), SLIDE_W - Inches(2), Inches(2)).text_frame,
              title, 44, t["fg"], bold=True, align=PP_ALIGN.CENTER)
    if subtitle:
        _set_text(slide.shapes.add_textbox(Inches(1), Inches(3.9), SLIDE_W - Inches(2), Inches(0.8)).text_frame,
                  subtitle, 18, t["muted"], align=PP_ALIGN.CENTER)


def _add_toc(slide, title, items, t):
    _add_title(slide, title, t)
    box = slide.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(11.0), Inches(4.8))
    tf = box.text_frame
    tf.word_wrap = True
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_before = Pt(20)
        run = p.add_run()
        run.text = f"{i + 1:>2}    {it}"
        run.font.name = FONT
        run.font.size = Pt(22)
        run.font.color.rgb = t["fg"] if i == 0 else t["muted"]
        run.font.bold = (i == 0)


def _add_section(slide, title, points, t):
    box = slide.shapes.add_textbox(Inches(0.8), Inches(2.7), SLIDE_W - Inches(1.6), Inches(1.5))
    _set_text(box.text_frame, title, 42, t["fg"], bold=True, align=PP_ALIGN.CENTER)
    if points:
        _set_text(slide.shapes.add_textbox(Inches(0.8), Inches(4.2), SLIDE_W - Inches(1.6), Inches(0.8)).text_frame,
                  points[0], 18, t["muted"], align=PP_ALIGN.CENTER)
    line = slide.shapes.add_shape(1, Inches(3.5), Inches(4.0), Inches(6.3), Inches(0.05))
    line.fill.solid()
    line.fill.fore_color.rgb = t["accent"]
    line.line.fill.background()


def _add_content(slide, title, points, image_path, t):
    _add_title(slide, title, t)
    if image_path and os.path.exists(image_path):
        text_box = slide.shapes.add_textbox(Inches(0.6), Inches(1.9), Inches(6.2), Inches(4.9))
        img_area_x, img_area_y = Inches(7.2), Inches(0.9)
        img_area_w, img_area_h = Inches(5.5), Inches(5.7)
    else:
        text_box = slide.shapes.add_textbox(Inches(0.6), Inches(1.9), Inches(12.0), Inches(4.9))
        img_area_x = img_area_y = img_area_w = img_area_h = 0

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

    if image_path and os.path.exists(image_path):
        with Image.open(image_path) as im:
            w, h = im.size
        scale = min(img_area_w / w, img_area_h / h)
        disp_w, disp_h = int(w * scale), int(h * scale)
        slide.shapes.add_picture(image_path, img_area_x + int((img_area_w - disp_w) / 2),
                                 img_area_y + int((img_area_h - disp_h) / 2),
                                 width=Emu(disp_w), height=Emu(disp_h))
    elif not image_path:
        rect = slide.shapes.add_shape(1, Inches(7.2), Inches(0.9), Inches(5.5), Inches(5.7))
        rect.fill.solid()
        rect.fill.fore_color.rgb = t["muted"]
        rect.line.fill.background()
        _set_text(rect.text_frame, "（图片生成失败）", 12, t["fg"], align=PP_ALIGN.CENTER)


def _add_data(slide, title, chart, t):
    _add_title(slide, title, t)
    chart_data = CategoryChartData()
    chart_data.categories = chart.get("labels", [])
    chart_data.add_series("数值", chart.get("values", []))
    gframe = slide.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(0.6), Inches(2.0),
                                    Inches(12.0), Inches(4.9), chart_data)
    chart_obj = gframe.chart
    chart_obj.has_legend = False
    try:
        chart_obj.value_axis.has_major_gridlines = False
        chart_obj.category_axis.tick_labels.font.color.rgb = t["fg"]
        chart_obj.value_axis.tick_labels.font.color.rgb = t["muted"]
        chart_obj.category_axis.tick_labels.font.size = Pt(12)
        chart_obj.value_axis.tick_labels.font.size = Pt(10)
    except Exception:
        pass
    if chart_obj.plots and chart_obj.plots[0].series:
        series = chart_obj.plots[0].series[0]
        series.format.fill.solid()
        series.format.fill.fore_color.rgb = t["accent"]
        try:
            series.data_labels.font.color.rgb = t["fg"]
            series.data_labels.font.size = Pt(10)
            series.data_labels.number_format = "0"
            series.data_labels.number_format_is_linked = False
            chart_obj.plots[0].has_data_labels = True
        except Exception:
            pass


def _add_end(slide, title, t):
    box = slide.shapes.add_textbox(Inches(0.8), Inches(3.0), SLIDE_W - Inches(1.6), Inches(1.6))
    _set_text(box.text_frame, title, 36, t["fg"], bold=True, align=PP_ALIGN.CENTER)


def build_ppt(slides: list[dict], image_paths: list[str | None], out_path: str,
              theme: str = "blue", subtitle: str = "") -> str:
    t = THEMES.get(theme, THEMES["blue"])
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    blank = prs.slide_layouts[6]

    for i, s in enumerate(slides):
        slide = prs.slides.add_slide(blank)
        stype = s.get("type", "")
        title = s.get("title", "")
        points = s.get("points", [])
        _add_bg(slide, t["bg"])

        if stype == "cover" or (i == 0 and not points):
            _add_cover(slide, title, subtitle or s.get("image_prompt", ""), t)
        elif stype == "toc":
            _add_toc(slide, title, points, t)
        elif stype == "section":
            _add_section(slide, title, points, t)
        elif stype == "data" and s.get("chart") and s["chart"].get("values"):
            _add_data(slide, title, s["chart"], t)
        elif stype == "end":
            _add_end(slide, title, t)
        else:  # content / 未知 type
            _add_content(slide, title, points, image_paths[i] if i < len(image_paths) else None, t)

    prs.save(out_path)
    return out_path
