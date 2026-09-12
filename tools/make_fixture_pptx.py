"""造契约 §V12 要求的自测素材稿（真实 PPT 缺失时的自缓解）。

产物落在 output/fixtures/。覆盖 builder.py 规整稿覆盖不到的六类形状/排版：

  fixtures_cover.pptx       一稿六页，逐页覆盖：
    1 组合形状（含 chOff/chExt 缩放，非恒等映射）
    2 表格 + 合并单元格
    3 图片
    4 图表（column）
    5 显式行距(Length) + 首段 space_before + normAutofit(fontScale)
    6 中英混排 + 超长无空格英文词

**覆盖不了的**（写进 IMPL_REPORT）：真实世界的字体替换（稿里声明宋体但机器
无该字体 → PowerPoint 回退）、SmartArt、嵌入视频、PowerPoint 自己重排后的
版式（本稿由 python-pptx 直写 XML，未经 PowerPoint 打开保存过）。

跑法：.venv/Scripts/python.exe tools/make_fixture_pptx.py
"""

import os

from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

OUT_DIR = os.path.abspath("output/fixtures")
ASSET_DIR = os.path.join(OUT_DIR, "assets")

FONT = "微软雅黑"


def _style(run, size_pt: float, bold: bool = False, color=(0x1F, 0x2A, 0x44)):
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.name = FONT
    run.font.color.rgb = RGBColor(*color)


def _add_title(slide, text, size_pt=32.0):
    box = slide.shapes.add_textbox(Inches(0.6), Inches(0.4), Inches(12.1), Inches(1.0))
    box.name = "TitleBox"
    tf = box.text_frame
    tf.word_wrap = True
    tf.text = text
    _style(tf.paragraphs[0].runs[0], size_pt, bold=True)
    return box


# ---------------------------------------------------------------- 1 组合形状

def _page_group(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_title(slide, "组合形状页")

    group = slide.shapes.add_group_shape()
    group.name = "SampleGroup"
    # 子形状写在组合的**子坐标系**里（原点 0,0 起）
    t1 = group.shapes.add_textbox(Emu(0), Emu(0), Emu(2000000), Emu(600000))
    t1.text_frame.text = "组合内第一块"
    _style(t1.text_frame.paragraphs[0].runs[0], 20.0)

    t2 = group.shapes.add_textbox(Emu(0), Emu(800000), Emu(2000000), Emu(600000))
    t2.text_frame.text = "组合内第二块"
    _style(t2.text_frame.paragraphs[0].runs[0], 20.0)

    # 显式写 xfrm：off/ext 是"落在页面上的位置与大小"，chOff/chExt 是子坐标系，
    # 两者不同 → 触发 §2.2 的仿射缩放（真实 PowerPoint 组合也会这么写）。
    xfrm = group._element.find(qn("p:grpSpPr")).find(qn("a:xfrm"))
    off = xfrm.find(qn("a:off"))
    ext = xfrm.find(qn("a:ext"))
    ch_off = xfrm.find(qn("a:chOff"))
    ch_ext = xfrm.find(qn("a:chExt"))
    off.set("x", str(Inches(1.5))); off.set("y", str(Inches(2.0)))
    ext.set("cx", str(Inches(6.0))); ext.set("cy", str(Inches(3.0)))
    ch_off.set("x", "0"); ch_off.set("y", "0")
    ch_ext.set("cx", str(Inches(3.0))); ch_ext.set("cy", str(Inches(3.0)))
    # → 水平放大 2x、垂直不变，子形状矩形应落在 (1.5in,2in)+ 的 2 倍宽处
    return slide


# ---------------------------------------------------------------- 2 表格合并

def _page_table(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_title(slide, "表格与合并单元格页")

    rows, cols = 4, 3
    shape = slide.shapes.add_table(rows, cols, Inches(1.0), Inches(1.8),
                                   Inches(11.0), Inches(4.0))
    shape.name = "SampleTable"
    table = shape.table
    data = [["指标", "2025", "2026"],
            ["营收", "1280 万元", "1960 万元"],
            ["用户数", "42 万", "78 万"],
            ["合并说明：以上数据为示例", "", ""]]
    for r in range(rows):
        for c in range(cols):
            cell = table.cell(r, c)
            cell.text_frame.text = data[r][c]
            if data[r][c]:
                _style(cell.text_frame.paragraphs[0].runs[0], 16.0)
    table.cell(3, 0).merge(table.cell(3, 2))  # 末行三列合并
    return slide


# ---------------------------------------------------------------- 3 图片

def _page_picture(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_title(slide, "图片页")
    img_path = os.path.join(ASSET_DIR, "fixture_image.png")
    if not os.path.isfile(img_path):
        raise RuntimeError(f"缺少素材图：{img_path}（先跑 _make_asset_image）")
    pic = slide.shapes.add_picture(img_path, Inches(3.6), Inches(2.2),
                                   Inches(6.0), Inches(3.6))
    pic.name = "SamplePicture"
    cap = slide.shapes.add_textbox(Inches(3.6), Inches(6.0), Inches(6.0), Inches(0.6))
    cap.text_frame.text = "图 1 · 示例配图"
    _style(cap.text_frame.paragraphs[0].runs[0], 14.0)
    return slide


# ---------------------------------------------------------------- 4 图表

def _page_chart(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_title(slide, "图表页")
    chart_data = CategoryChartData()
    chart_data.categories = ["一季度", "二季度", "三季度", "四季度"]
    chart_data.add_series("营收", (120.0, 180.0, 240.0, 310.0))
    gf = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1.2), Inches(1.8),
                                Inches(6.4), Inches(4.4), chart_data)
    gf.name = "SampleChart"
    note = slide.shapes.add_textbox(Inches(8.0), Inches(2.4), Inches(4.4), Inches(2.4))
    note.text_frame.word_wrap = True
    for i, line in enumerate(["数据说明：", "• 全年营收增长 158%", "• 四季度环比最高"]):
        p = note.text_frame.paragraphs[0] if i == 0 else note.text_frame.add_paragraph()
        p.add_run().text = line
        _style(p.runs[0], 18.0)
    return slide


# ---------------------------------------------------------------- 5 排版

def _page_typography(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_title(slide, "行距 / 段前距 / 自动缩排页")

    box = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(7.0), Inches(4.6))
    box.name = "TypographyBox"
    tf = box.text_frame
    tf.word_wrap = True
    lines = ["首段带段前距 20pt（验证 PowerPoint 是否渲染首段段前距）",
             "第二段使用 1.8 倍行距",
             "第三段使用固定 30pt 行距",
             "第四段结束"]
    for i, text in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.add_run().text = text
        _style(p.runs[0], 16.0)
    tf.paragraphs[0].space_before = Pt(20)
    tf.paragraphs[1].line_spacing = 1.8
    tf.paragraphs[2].line_spacing = Pt(30)

    # 注入 normAutofit fontScale=60000（60%），验证 python-pptx 可读、PowerPoint 是否照办
    autofit_box = slide.shapes.add_textbox(Inches(8.4), Inches(1.8), Inches(4.0), Inches(2.0))
    autofit_box.name = "AutofitBox"
    atf = autofit_box.text_frame
    atf.word_wrap = True
    for i in range(6):
        p = atf.paragraphs[0] if i == 0 else atf.add_paragraph()
        p.add_run().text = f"自动缩排测试行 {i + 1}：这一行故意写长一点让它溢出文本框高度"
        _style(p.runs[0], 18.0)
    bodyPr = atf._txBody.find(qn("a:bodyPr"))
    for child in list(bodyPr):
        if child.tag in (qn("a:normAutofit"), qn("a:noAutofit"), qn("a:spAutoFit")):
            bodyPr.remove(child)
    norm = bodyPr.makeelement(qn("a:normAutofit"), {"fontScale": "60000", "lnSpcReduction": "20000"})
    bodyPr.append(norm)
    return slide


# ---------------------------------------------------------------- 6 中英混排

def _page_mixed(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_title(slide, "中英混排与超长英文词页")

    box = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(7.4), Inches(4.6))
    box.name = "MixedBox"
    tf = box.text_frame
    tf.word_wrap = True
    texts = [
        "中英混排：AI 与 Machine Learning 正在改变 Education 行业",
        "超长无空格词：Pneumonoultramicroscopicsilicovolcanoconiosis 会如何断行",
        "全角标点混排：你好，世界！这是“全角”标点（测试）——是否断行正确？",
        "行首行尾空格：   前导空格与尾随空格   ",
    ]
    for i, t in enumerate(texts):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.add_run().text = t
        _style(p.runs[0], 16.0)

    narrow = slide.shapes.add_textbox(Inches(8.6), Inches(1.8), Inches(1.2), Inches(4.6))
    narrow.name = "NarrowBox"
    narrow.text_frame.word_wrap = True
    p = narrow.text_frame.paragraphs[0]
    p.add_run().text = "窄框中的超长英文词 Pneumonoultramicroscopicsilicovolcanoconiosis"
    _style(p.runs[0], 16.0)
    return slide


# ---------------------------------------------------------------- 素材图

def _make_asset_image():
    os.makedirs(ASSET_DIR, exist_ok=True)
    path = os.path.join(ASSET_DIR, "fixture_image.png")
    img = Image.new("RGB", (800, 480), (0xF3, 0xF6, 0xFB))
    d = ImageDraw.Draw(img)
    d.rectangle([40, 40, 760, 440], outline=(0x2B, 0x5C, 0xE0), width=6)
    d.line([40, 440, 400, 120], fill=(0x2B, 0x5C, 0xE0), width=8)
    d.line([400, 120, 760, 320], fill=(0xE0, 0x6B, 0x2B), width=8)
    d.ellipse([360, 80, 440, 160], fill=(0xE0, 0x6B, 0x2B))
    img.save(path)
    return path


def build(out_path: str) -> str:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    _page_group(prs)
    _page_table(prs)
    _page_picture(prs)
    _page_chart(prs)
    _page_typography(prs)
    _page_mixed(prs)
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    prs.save(out_path)
    return out_path


def main() -> int:
    os.makedirs(OUT_DIR, exist_ok=True)
    _make_asset_image()
    out = build(os.path.join(OUT_DIR, "fixtures_cover.pptx"))
    print(f"已生成：{out}（{os.path.getsize(out)} 字节）")

    # 立刻回读核验：每页形状类型齐全、组合 xfrm 非退化
    prs = Presentation(out)
    print(f"页数={len(prs.slides)}，画布={prs.slide_width}x{prs.slide_height} EMU")
    for i, slide in enumerate(prs.slides, 1):
        kinds = []
        for sh in slide.shapes:
            kinds.append(_kind(sh))
        print(f"  第{i}页：{kinds}")
    return 0


def _kind(sh):
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
        return f"group({len(sh.shapes)})"
    if getattr(sh, "has_table", False) and sh.has_table:
        return "table"
    if getattr(sh, "has_chart", False) and sh.has_chart:
        return "chart"
    if sh.shape_type == MSO_SHAPE_TYPE.PICTURE:
        return "picture"
    if sh.has_text_frame:
        return f"text:{sh.name}"
    return "other"


if __name__ == "__main__":
    raise SystemExit(main())
