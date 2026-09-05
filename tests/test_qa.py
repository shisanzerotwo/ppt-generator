"""qa 模块测试：换行模拟 + check_pptx 回读校验 + a:ea 中文字体回写。"""

import pytest
from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

import qa
from builder import FONT, build_ppt


# ---------- measure_text_lines：换行模拟 ----------

def test_measure_cjk_breaks_per_char():
    # CJK 逐字可断：每字 10pt，box 95pt 放 9 字 → 2 行；box 105pt 一行放下 → 1 行
    assert qa.measure_text_lines("一" * 10, 10, 95) == 2
    assert qa.measure_text_lines("一" * 10, 10, 105) == 1


def test_measure_60_cjk_title_three_lines():
    # 用例①的度量基础：60 字 18pt 在 6.5in 标题框有效宽（453.6pt）内 → 3 行
    assert qa.measure_text_lines("长" * 60, 18, 453.6) == 3


def test_measure_latin_word_breaks_at_space():
    # 拉丁词整体在空格处断行："hello world hello" 每行一词 → 3 行
    # （若错误地逐字强拆，box 40pt 每行可塞 7 字符 → 2 行）
    assert qa.measure_text_lines("hello world hello", 10, 40) == 3


def test_measure_long_word_force_break():
    # 超长词（整词宽超过行宽）逐字强拆：box 恰放 9 个 'a'，20 个 'a' → 9+9+2 = 3 行
    char_w = qa._char_em_width("a") * 10
    assert qa.measure_text_lines("a" * 20, 10, char_w * 9.5) == 3


def test_measure_leading_space_dropped():
    # 行首空格丢弃：带前导空格与不带空格的行数一致
    # box 31.5pt：丢弃空格每行 3 字（3 行）；若不丢弃每行只放 2 字（4 行）
    assert qa.measure_text_lines("一二三四五六七", 10, 31.5) == 3
    assert qa.measure_text_lines("  一二三四五六七", 10, 31.5) == 3


def test_measure_empty_and_spaces_only():
    assert qa.measure_text_lines("", 18, 100) == 1
    assert qa.measure_text_lines("   ", 18, 100) == 1


def test_measure_never_raises_without_font(monkeypatch):
    # 字体不可用时降级估算，不抛异常
    monkeypatch.setattr(qa, "_char_em_width", lambda ch: None)
    assert qa.measure_text_lines("中文 mixed 123", 18, 200) >= 1


# ---------- check_pptx ----------

def _base_slides():
    """覆盖全部版式的正常小稿（8 页）。"""
    return [
        {"type": "cover", "title": "主题", "points": [], "image_prompt": "封面插画"},
        {"type": "toc", "title": "目录", "points": ["章节一", "章节二"], "image_prompt": ""},
        {"type": "section", "title": "第一章", "points": ["一句话简介"], "image_prompt": ""},
        {"type": "content", "title": "要点页", "points": ["要点一：说明", "要点二：说明", "要点三：说明"],
         "image_prompt": ""},
        {"type": "data", "title": "数据页", "points": [], "image_prompt": "",
         "chart": {"type": "column", "labels": ["A", "B"], "values": [30, 70]}},
        {"type": "timeline", "title": "发展历程", "points": ["2015年：成立", "2020年：扩张"], "image_prompt": ""},
        {"type": "compare", "title": "对比", "points": ["左一", "左二", "右一", "右二"], "image_prompt": ""},
        {"type": "end", "title": "感谢观看", "points": [], "image_prompt": ""},
    ]


def test_check_pptx_long_title_warns_without_errors(tmp_path):
    """用例①：60+ 字超长标题 → 文本溢出 warning，且 0 error。"""
    slides = [
        {"type": "cover", "title": "封面", "points": [], "image_prompt": ""},
        {"type": "content", "title": "这是一个特别特别长的标题用来验证溢出检测逻辑是否正常工作" * 3,
         "points": ["要点一：说明", "要点二：说明", "要点三：说明"], "image_prompt": ""},
    ]
    out = str(tmp_path / "long.pptx")
    build_ppt(slides, [None, None], out)

    result = qa.check_pptx(out, expected_pages=2)
    assert result["errors"] == []
    assert len(result["warnings"]) == 1
    assert "溢出" in result["warnings"][0]
    assert "第2页" in result["warnings"][0]


def test_check_pptx_clean_deck_no_problems(tmp_path):
    """用例②：正常小稿 → 0 error 0 warning。"""
    out = str(tmp_path / "clean.pptx")
    build_ppt(_base_slides(), [None] * 8, out)

    result = qa.check_pptx(out, expected_pages=8)
    assert result["errors"] == []
    assert result["warnings"] == []


def test_check_pptx_page_count_mismatch(tmp_path):
    out = str(tmp_path / "count.pptx")
    build_ppt(_base_slides(), [None] * 8, out)

    result = qa.check_pptx(out, expected_pages=99)
    assert len(result["errors"]) == 1
    assert "页数不符" in result["errors"][0]


def test_check_pptx_out_of_bound_shape(tmp_path):
    """用例③：手工造一个形状越界的文件 → error。"""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_textbox(Inches(12), Inches(6), Inches(3), Inches(3))
    box.text_frame.text = "越界形状"
    out = str(tmp_path / "oob.pptx")
    prs.save(out)

    result = qa.check_pptx(out, expected_pages=1)
    assert len(result["errors"]) == 1
    assert "越界" in result["errors"][0] or "超出画布" in result["errors"][0]
    assert result["warnings"] == []


def test_check_pptx_unreadable_file(tmp_path):
    result = qa.check_pptx(str(tmp_path / "not_exist.pptx"), expected_pages=1)
    assert len(result["errors"]) == 1
    assert "无法打开" in result["errors"][0]


def test_check_pptx_used_estimate_flag(monkeypatch, tmp_path):
    # 字体不可用时 used_estimate=True 且仍能完成校验
    out = str(tmp_path / "est.pptx")
    build_ppt(_base_slides(), [None] * 8, out)
    monkeypatch.setattr(qa, "_load_font", lambda: None)
    result = qa.check_pptx(out, expected_pages=8)
    assert result["used_estimate"] is True
    assert result["errors"] == []


# ---------- a:ea 中文字体回写 ----------

def test_all_runs_have_ea_typeface(tmp_path):
    """用例④：导出 pptx 中任意 run 的 rPr 均含 a:ea 且 typeface=微软雅黑。"""
    out = str(tmp_path / "ea.pptx")
    build_ppt(_base_slides(), [None] * 8, out)

    prs = Presentation(out)
    run_count = 0
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    run_count += 1
                    rPr = run._r.rPr
                    assert rPr is not None, f"run '{run.text}' 缺 rPr"
                    ea = rPr.find(qn("a:ea"))
                    assert ea is not None, f"run '{run.text}' 缺 a:ea"
                    assert ea.get("typeface") == FONT
    assert run_count > 0  # 确认真的遍历到了 run，避免空转假绿


def test_textbox_autosize_is_none(tmp_path):
    # 补充：所有文本框均显式禁用自动调整（防止 PowerPoint 打开时长高）
    out = str(tmp_path / "autosize.pptx")
    build_ppt(_base_slides(), [None] * 8, out)

    prs = Presentation(out)
    checked = 0
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            bodyPr = shape.text_frame._txBody.find(qn("a:bodyPr"))
            assert bodyPr is not None
            assert bodyPr.find(qn("a:spAutoFit")) is None, f"{shape.name} 仍带 spAutoFit"
            checked += 1
    assert checked > 0
