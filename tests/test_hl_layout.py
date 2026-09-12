"""M2 前半：hl_layout.wrap_lines 与 qa.measure_text_lines 的**等价测试**。

这是防两套度量算法漂移的强制闸门：高亮矩形按 wrap_lines 的断点画，PPTX 溢出
QA 按 qa 的行数算，一旦两边不一致，同一段文字会"高亮在这行、溢出告警算那行"。
"""

import pytest

import hl_layout
import qa

# 契约 §4.2 点名的 8 类语料
CORPUS = {
    "纯中文": "人工智能正在改变教育行业的每一个环节与角落",
    "纯英文": "Artificial intelligence is reshaping education industry",
    "中英混排": "AI 与 Machine Learning 正在改变 Education 行业",
    "超长无空格英文词": "Pneumonoultramicroscopicsilicovolcanoconiosis 之后还有正文",
    "全角标点": "你好，世界！这是“全角”标点（测试）——是否断行正确？",
    "行首行尾空格": "   前导空格与尾随空格   ",
    "单字宽超行宽": "中",
    "空串": "",
}

WIDTHS = [1.0, 5.0, 12.0, 30.0, 80.0, 160.0, 300.0, 640.0]
SIZES = [12.0, 18.0, 28.0]


@pytest.mark.parametrize("name", list(CORPUS))
@pytest.mark.parametrize("size_pt", SIZES)
@pytest.mark.parametrize("box_width_pt", WIDTHS)
def test_wrap_lines_matches_qa_line_count(name, size_pt, box_width_pt):
    """行数必须逐条相等——这是两套算法同源的证明。"""
    text = CORPUS[name]
    assert len(hl_layout.wrap_lines(text, size_pt, box_width_pt)) == \
        qa.measure_text_lines(text, size_pt, box_width_pt)


def test_wrap_lines_fuzz_matches_qa():
    """伪随机语料 fuzz：混合中英、空格、标点、长词，逐条对行数。"""
    import random
    rng = random.Random(20260912)
    alphabet = list("人工智能教育行业发展") + list("abcdefgXYZ .,-_123") + ["，", "。", "！", "—"]
    for _ in range(300):
        n = rng.randint(0, 40)
        text = "".join(rng.choice(alphabet) for _ in range(n))
        size = rng.choice([10.0, 14.0, 18.0, 24.0, 36.0])
        box = rng.choice([8.0, 20.0, 45.0, 90.0, 200.0, 500.0])
        assert len(hl_layout.wrap_lines(text, size, box)) == \
            qa.measure_text_lines(text, size, box), \
            f"漂移：{text!r} size={size} box={box}"


def test_empty_text_returns_single_empty_line():
    """空文本与 qa 的 lines=1 对齐；纯空格也只剩一行空行。"""
    assert hl_layout.wrap_lines("", 18.0, 100.0) == [hl_layout.Line("", 0.0, 18.0)]
    lines = hl_layout.wrap_lines("   ", 18.0, 100.0)
    assert len(lines) == 1 and lines[0].text == ""


def test_cjk_wraps_char_by_char():
    """逐字可断：18pt 中文每字 18pt，100pt 宽一行放 5 字。"""
    lines = hl_layout.wrap_lines("人工智能教育行业", 18.0, 100.0)
    assert len(lines) == 2
    assert lines[0].text == "人工智能教"
    assert lines[1].text == "育行业"


def test_latin_breaks_at_space_not_midword():
    """拉丁整词不拆：断点只能落在空格处，每行里都是完整单词。

    （行**内**的空格是合法的——"alpha beta" 能整行放下时两词同行。）
    """
    words = "alpha beta gamma delta".split()
    lines = hl_layout.wrap_lines("alpha beta gamma delta", 18.0, 120.0)
    assert len(lines) > 1, "该宽度下应当换行"
    for line in lines:
        assert line.text.split() and all(w in words for w in line.text.split())
    assert [w for line in lines for w in line.text.split()] == words


def test_leading_spaces_dropped_on_wrap():
    """行首空格丢弃：换行后行首不留空格（与 qa 同规则）。"""
    lines = hl_layout.wrap_lines("人工智能 教育行业", 18.0, 100.0)
    assert lines[1].text.startswith("教") or lines[1].text[0] != " "


def test_overlong_word_force_split():
    """整词宽 > 行宽 → 逐字强拆，且每行宽度不超过行宽（除非单字本身就超）。"""
    word = "Pneumonoultramicroscopicsilicovolcanoconiosis"
    lines = hl_layout.wrap_lines(word, 18.0, 100.0)
    assert len(lines) > 1
    assert "".join(line.text for line in lines) == word
    assert all(line.width_pt <= 100.0 or len(line.text) == 1 for line in lines)


def test_trailing_spaces_not_counted_in_width():
    """行尾空格的宽度不计入该行宽度（否则高亮条会拖出一截空白）。"""
    lines = hl_layout.wrap_lines("人工智能    ", 18.0, 200.0)
    assert lines[0].text == "人工智能"
    assert lines[0].width_pt == pytest.approx(4 * 18.0, rel=0.02)


def test_line_width_is_sum_of_char_widths():
    """行宽自洽：等于该行字符宽度之和。"""
    lines = hl_layout.wrap_lines("人工智能教育", 20.0, 200.0)
    for line in lines:
        expect = sum(qa._char_width_pt(c, 20.0) for c in line.text)
        assert line.width_pt == pytest.approx(expect, rel=1e-9)


def test_single_char_wider_than_box_occupies_own_line():
    """单字宽 > 行宽：独占一行（强行放置），不进入死循环。"""
    lines = hl_layout.wrap_lines("中中中", 200.0, 50.0)
    assert len(lines) == 3
    assert all(line.text == "中" for line in lines)


def test_wrap_lines_preserves_all_non_space_chars():
    """不丢字：去掉空格后拼接内容与原文一致（8 类语料全跑）。"""
    for name, text in CORPUS.items():
        for box in (30.0, 120.0, 400.0):
            lines = hl_layout.wrap_lines(text, 18.0, box)
            joined = "".join(line.text for line in lines).replace(" ", "")
            assert joined == text.replace(" ", ""), f"{name} @ {box}pt 丢字"
