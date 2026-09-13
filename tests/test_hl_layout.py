"""M2 前半：hl_layout.wrap_lines 与 qa.measure_text_lines 的**等价测试**。

这是防两套度量算法漂移的强制闸门：高亮矩形按 wrap_lines 的断点画，PPTX 溢出
QA 按 qa 的行数算，一旦两边不一致，同一段文字会"高亮在这行、溢出告警算那行"。
"""

import os

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


# ---------------------------------------------------------------- 定位层

def _shape(**kw):
    from pptx_io import ShapeInfo
    base = dict(shape_id=1, name="S", kind="text", left_emu=12700 * 100, top_emu=12700 * 50,
                width_emu=12700 * 200, height_emu=12700 * 100)
    base.update(kw)
    return ShapeInfo(**base)


def _para(text, size=18.0, **kw):
    from pptx_io import ParaInfo, RunInfo
    return ParaInfo(text=text, runs=[RunInfo(text=text, size_pt=size, bold=None,
                                             italic=None, font_name=None)], **kw)


def _page(shapes, w=12700 * 1920, h=12700 * 1080):
    from pptx_io import PageShapes
    return PageShapes(index=0, width_emu=w, height_emu=h, shapes=shapes)


def test_build_units_positions_first_line_inside_margins():
    """行矩形从框内边距处起算，并退回 pad 补偿量。"""
    sh = _shape(paragraphs=[_para("人工智能教育")])
    units = hl_layout.build_units(_page([sh]), export_width_px=1920, pad_x_pt=2.0, pad_y_pt=1.0)
    assert len(units) == 1
    line = units[0].lines[0]
    # inner_left = 100pt + 7.2pt = 107.2pt；减 pad 2pt → 105.2pt
    assert line.left_emu == round(105.2 * 12700)
    assert line.top_emu == round((50 + 3.6 - 1.0) * 12700)


def test_build_units_px_uses_export_width():
    """px 是派生量：导出宽减半，px 也减半，EMU 不变。"""
    sh = _shape(paragraphs=[_para("一二三四")])
    a = hl_layout.build_units(_page([sh]), export_width_px=1920)[0]
    b = hl_layout.build_units(_page([sh]), export_width_px=960)[0]
    assert a.lines[0].left_emu == b.lines[0].left_emu
    assert a.lines[0].left_px == pytest.approx(b.lines[0].left_px * 2)


def test_build_units_multiline_advances_cursor():
    """多行：每行一个 rect，top 按 line_h = 字号×1.25 递增。"""
    sh = _shape(width_emu=12700 * 80, paragraphs=[_para("人工智能教育行业")])
    lines = hl_layout.build_units(_page([sh]))[0].lines
    assert len(lines) > 1
    pitch = lines[1].top_emu - lines[0].top_emu
    assert pitch == pytest.approx(18.0 * qa.LINE_HEIGHT_FACTOR * 12700, rel=1e-6)


def test_build_units_middle_anchor_shifts_block_down():
    """MIDDLE 锚点：整块按剩余高度一半下移（builder 稿大量使用）。"""
    mid = _shape(height_emu=12700 * 300, vertical_anchor="MIDDLE", paragraphs=[_para("短")])
    top = hl_layout.build_units(_page([mid]))[0].lines[0].top_px
    plain_sh = _shape(height_emu=12700 * 300, paragraphs=[_para("短")])
    plain = hl_layout.build_units(_page([plain_sh]))[0].lines[0].top_px
    assert top > plain


def test_build_units_bottom_anchor_shifts_more_than_middle():
    mid = _shape(height_emu=12700 * 300, vertical_anchor="MIDDLE", paragraphs=[_para("短")])
    bot = _shape(height_emu=12700 * 300, vertical_anchor="BOTTOM", paragraphs=[_para("短")])
    y_mid = hl_layout.build_units(_page([mid]))[0].lines[0].top_px
    y_bot = hl_layout.build_units(_page([bot]))[0].lines[0].top_px
    assert y_bot > y_mid


def test_build_units_center_and_right_alignment():
    x_left = hl_layout.build_units(_page([
        _shape(width_emu=12700 * 400, paragraphs=[_para("短")])]))[0].lines[0].left_px
    x_center = hl_layout.build_units(_page([
        _shape(width_emu=12700 * 400, paragraphs=[_para("短", align="CENTER")])]))[0].lines[0].left_px
    x_right = hl_layout.build_units(_page([
        _shape(width_emu=12700 * 400, paragraphs=[_para("短", align="RIGHT")])]))[0].lines[0].left_px
    assert x_left < x_center < x_right


def test_build_units_justify_falls_back_to_left_with_warning():
    sh = _shape(paragraphs=[_para("两端对齐", align="JUSTIFY")])
    u = hl_layout.build_units(_page([sh]))[0]
    assert u.align == "LEFT" and "justify_approximated" in u.warnings


def test_build_units_left_align_none_is_normalized():
    u = hl_layout.build_units(_page([_shape(paragraphs=[_para("继承对齐")])]))[0]
    assert u.align == "LEFT"


def test_build_units_font_scale_scales_size_and_warns():
    sh = _shape(font_scale=60.0, paragraphs=[_para("自动缩排", size=20.0)])
    u = hl_layout.build_units(_page([sh]))[0]
    assert u.size_pt == pytest.approx(12.0)
    assert "autofit_scaled" in u.warnings


def test_build_units_vertical_anchor_none_warns_inherited():
    u = hl_layout.build_units(_page([_shape(paragraphs=[_para("继承锚点")])]))[0]
    assert "vertical_anchor_inherited" in u.warnings


def test_build_units_skips_shape_with_no_inner_width():
    """内宽 <= 0（边距吃掉整宽）→ 不出 Unit，不抛异常。"""
    sh = _shape(width_emu=12700 * 10, margin_left_emu=91440, margin_right_emu=91440,
                paragraphs=[_para("放不下")])
    assert hl_layout.build_units(_page([sh])) == []


def test_build_units_skips_blank_paragraph():
    sh = _shape(paragraphs=[_para("   "), _para("有内容")])
    units = hl_layout.build_units(_page([sh]))
    assert len(units) == 1 and units[0].text == "有内容"


def test_build_units_first_paragraph_space_before_ignored():
    """首段段前距不计入（V5 实测 PowerPoint 忽略文本框首段段前距），后续段落计入。"""
    sh = _shape(paragraphs=[_para("一", space_before_pt=20.0),
                            _para("二", space_before_pt=20.0)])
    units = hl_layout.build_units(_page([sh]))
    assert len(units) == 2
    gap = units[1].lines[0].top_emu - units[0].lines[0].top_emu
    # 第一段行高 18×1.25 = 22.5pt，加第二段段前距 20pt
    assert gap == pytest.approx((18 * qa.LINE_HEIGHT_FACTOR + 20) * 12700, rel=1e-6)


def test_build_units_order_is_document_order():
    shapes = [_shape(shape_id=i, name=f"S{i}", paragraphs=[_para(f"第{i}段")])
              for i in (1, 2, 3)]
    units = hl_layout.build_units(_page(shapes))
    assert [u.order for u in units] == [0, 1, 2]
    assert [u.shape_id for u in units] == [1, 2, 3]


def test_build_units_picture_is_one_block_unit():
    sh = _shape(kind="picture", shape_id=7, name="Pic",
                width_emu=12700 * 200, height_emu=12700 * 100)
    units = hl_layout.build_units(_page([sh]))
    assert len(units) == 1 and units[0].kind == "picture"
    assert units[0].lines == [units[0].rect]
    assert units[0].rect.width_emu == 12700 * 200


def test_build_units_chart_is_single_unit_not_split():
    """R4：图表整块一个单元，不拆数据点。"""
    units = hl_layout.build_units(_page([_shape(kind="chart", shape_id=9, name="Chart")]))
    assert len(units) == 1 and units[0].kind == "chart"


def test_build_units_table_cells_use_prefix_sums():
    from pptx_io import ShapeInfo
    sh = ShapeInfo(shape_id=3, name="T", kind="table", left_emu=0, top_emu=0,
                   width_emu=12700 * 300, height_emu=12700 * 200,
                   table_cells=[[[_para("A1")], [_para("")]],
                                [[_para("")], [_para("B2")]]],
                   table_col_widths_emu=[12700 * 100, 12700 * 200],
                   table_row_heights_emu=[12700 * 100, 12700 * 100])
    units = hl_layout.build_units(_page([sh]))
    assert [u.kind for u in units] == ["cell", "cell"]
    a1, b2 = units
    assert b2.lines[0].top_emu - a1.lines[0].top_emu == pytest.approx(12700 * 100)
    assert b2.lines[0].left_emu - a1.lines[0].left_emu == pytest.approx(12700 * 100)


def test_build_units_walks_group_children():
    from pptx_io import ShapeInfo
    child = _shape(shape_id=11, name="Child", paragraphs=[_para("组内")])
    group = ShapeInfo(shape_id=10, name="G", kind="group", left_emu=0, top_emu=0,
                      width_emu=12700 * 400, height_emu=12700 * 200, children=[child])
    units = hl_layout.build_units(_page([group]))
    assert len(units) == 1 and units[0].shape_id == 11


def test_build_units_other_kind_produces_nothing():
    assert hl_layout.build_units(_page([_shape(kind="other", shape_id=5)])) == []


def test_build_units_rejects_dict_without_deck():
    """deck.json 的 page dict 不带画布尺寸 → 必须显式过 page_shapes()。"""
    with pytest.raises(ValueError, match="page_shapes"):
        hl_layout.build_units({"index": 0, "shapes": []})


def test_page_shapes_bridges_deck_dict():
    from types import SimpleNamespace
    deck = SimpleNamespace(width_emu=100, height_emu=50)
    ps = hl_layout.page_shapes({"index": 2, "bg": "bg/slide_1.png", "shapes": []}, deck)
    assert (ps.index, ps.width_emu, ps.height_emu) == (2, 100, 50)


def test_page_shapes_passthrough():
    p = _page([])
    assert hl_layout.page_shapes(p) is p


def test_title_classified_by_relative_font_size():
    """标题判定：字号 >= 1.3× 本页中位字号（builder 的框名是 TextBox N，名字判不出）。"""
    units = hl_layout.build_units(_page([
        _shape(shape_id=1, name="TextBox 1", paragraphs=[_para("大标题", size=40.0)]),
        _shape(shape_id=2, name="TextBox 2", paragraphs=[_para("正文一", size=18.0)]),
        _shape(shape_id=3, name="TextBox 3", paragraphs=[_para("正文二", size=18.0)]),
    ]))
    kinds = {u.text: u.kind for u in units}
    assert kinds["大标题"] == "title"
    assert kinds["正文一"] in ("body", "bullet")


def test_title_classified_by_shape_name():
    u = hl_layout.build_units(_page([
        _shape(name="标题 1", paragraphs=[_para("小字", size=10.0)])]))[0]
    assert u.kind == "title"


# ---------------------------------------------------------------- 覆盖率度量

def _png(tmp_path, name, size=(60, 40), bg=(255, 255, 255), box=None):
    from PIL import Image, ImageDraw
    img = Image.new("RGB", size, bg)
    if box:
        ImageDraw.Draw(img).rectangle(box, fill=(0, 0, 0))
    p = tmp_path / name
    img.save(p)
    return str(p)


def test_measure_coverage_half_black(tmp_path):
    p = _png(tmp_path, "half.png", size=(40, 40), box=[0, 0, 19, 39])
    assert hl_layout.measure_coverage(p, (0, 0, 40, 40)) == pytest.approx(0.5, abs=0.03)


def test_measure_coverage_all_ink_with_explicit_bg(tmp_path):
    """整块墨迹：必须显式给底色——环取底色时整块同色矩形没有对比，测不出墨迹
    （这是契约默认 bg_rgb=None 的固有盲点，验收脚本因此显式传局部底色）。"""
    p = _png(tmp_path, "ink.png", size=(20, 20), bg=(0, 0, 0))
    assert hl_layout.measure_coverage(p, (0, 0, 20, 20), bg_rgb=(255, 255, 255)) == 1.0


def test_measure_coverage_ring_bg_blind_on_uniform_block(tmp_path):
    """锁住上面那条盲点：同色整块 + 默认环取底色 → 0.0（不是 bug，是口径使然）。"""
    p = _png(tmp_path, "blk.png", size=(20, 20), bg=(0, 0, 0))
    assert hl_layout.measure_coverage(p, (0, 0, 20, 20)) == 0.0


def test_measure_coverage_blank_is_zero(tmp_path):
    assert hl_layout.measure_coverage(_png(tmp_path, "blank.png"), (0, 0, 60, 40)) == 0.0


def test_measure_coverage_empty_rect_is_zero(tmp_path):
    p = _png(tmp_path, "e.png")
    assert hl_layout.measure_coverage(p, (5, 5, 0, 10)) == 0.0
    assert hl_layout.measure_coverage(p, (5, 5, 10, 0)) == 0.0


def test_measure_coverage_explicit_bg_overrides_ring(tmp_path):
    """显式底色优先：深底与浅底各自算得对（环取底色在紧贴墨迹时会失准）。"""
    from PIL import Image, ImageDraw
    img = Image.new("RGB", (40, 40), (200, 200, 200))
    ImageDraw.Draw(img).rectangle([0, 0, 39, 19], fill=(30, 30, 30))
    p = tmp_path / "two.png"
    img.save(p)
    dark = hl_layout.measure_coverage(str(p), (0, 0, 40, 20), bg_rgb=(30, 30, 30))
    light = hl_layout.measure_coverage(str(p), (0, 20, 40, 20), bg_rgb=(200, 200, 200))
    assert dark == pytest.approx(0.0, abs=0.02)
    assert light == pytest.approx(0.0, abs=0.02)


def test_measure_coverage_clamps_out_of_bounds_rect(tmp_path):
    p = _png(tmp_path, "small.png", size=(10, 10), bg=(0, 0, 0))
    assert hl_layout.measure_coverage(p, (5, 5, 100, 100), bg_rgb=(255, 255, 255)) == 1.0


def test_measure_coverage_respects_diff_threshold(tmp_path):
    from PIL import Image
    img = Image.new("RGB", (10, 10), (255, 255, 255))
    for x in range(10):
        img.putpixel((x, 0), (245, 245, 245))  # 差 10，低于默认阈值
    p = tmp_path / "thr.png"
    img.save(p)
    assert hl_layout.measure_coverage(str(p), (0, 0, 10, 10)) == 0.0
    assert hl_layout.measure_coverage(str(p), (0, 0, 10, 10), diff_thresh=5) > 0.0


# ---------------------------------------------------------------- 端到端验收（需底图）

_BG_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                       "output", "spike", "m1")
_DECK = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                     "output", "b_multislide.pptx")

needs_material = pytest.mark.skipif(
    not (os.path.isfile(_DECK) and os.path.isfile(os.path.join(_BG_DIR, "slide_1.png"))),
    reason="缺 output/b_multislide.pptx 或 COM 底图（先跑 tools/probes/accept_m1.py）")


def _band_bg(img, box, band=3):
    """矩形外侧 3px 环带的底色中位值（比内侧环稳，比幻灯片主色正确）。"""
    l, t, w, h = (int(round(v)) for v in box)
    la, ta = max(0, l - band), max(0, t - band)
    ra, ba = min(img.width, l + w + band), min(img.height, t + h + band)
    ri, bi = min(img.width, l + w), min(img.height, t + h)
    px = img.load()
    vals = []
    for y in range(ta, ba):
        for x in range(la, ra):
            if l <= x < ri and t <= y < bi:
                continue
            vals.append(px[x, y])
    if not vals:
        return None
    return tuple(sorted(v[c] for v in vals)[len(vals) // 2] for c in range(3))


@needs_material
def test_line_coverage_beats_shape_level_by_2_5x():
    """M2 端到端验收：行级覆盖率相对**形状级**提升 >= 2.5 倍（同一函数、同一底色）。

    为什么是相对提升而不是绝对阈值：绝对覆盖率的天花板由两个不可控量决定——
    行高框（line_h=字号×1.25）比 CJK 墨迹高约 27%，且字形墨迹在 em 盒里的密度
    只有 ~0.41（实测分解 0.731×0.960×0.414=0.286）。绝对目标 0.35 在本契约公式下
    不可达，详见 docs/IMPL_REPORT.md 的契约缺陷节。相对提升则只衡量"行级定位是否
    真的比形状级贴合"，与尺度的绝对值无关。
    """
    from PIL import Image

    import pptx_io

    deck, _ = pptx_io.read_pages(_DECK)
    shape_cov, line_cov = [], []
    for page in deck.pages:
        bg = os.path.join(_BG_DIR, f"slide_{page['index'] + 1}.png")
        img = Image.open(bg).convert("RGB")
        for shape in page["shapes"]:
            if shape.kind != "text":
                continue
            scale = 1920.0 / deck.width_emu
            box = (shape.left_emu * scale, shape.top_emu * scale,
                   shape.width_emu * scale, shape.height_emu * scale)
            bb = _band_bg(img, box)
            shape_cov.append(hl_layout.measure_coverage(bg, box, bg_rgb=bb))
        for u in hl_layout.build_units(hl_layout.page_shapes(page, deck)):
            for r in u.lines:
                box = (r.left_px, r.top_px, r.width_px, r.height_px)
                bb = _band_bg(img, box)
                line_cov.append(hl_layout.measure_coverage(bg, box, bg_rgb=bb))

    import statistics
    shape_med = statistics.median(shape_cov)
    line_med = statistics.median(line_cov)
    assert shape_med > 0, "形状级覆盖率不该为 0"
    assert line_med / shape_med >= 2.5, (
        f"行级/形状级 = {line_med:.3f}/{shape_med:.3f} = {line_med / shape_med:.2f}× < 2.5×")


@needs_material
def test_no_line_rect_escapes_its_shape_box():
    """框溢出率 <= 5%：行 rect 不该跑到所属形状外框之外（高亮跑出框外是观感事故）。"""
    from PIL import Image  # noqa: F401

    import pptx_io

    deck, _ = pptx_io.read_pages(_DECK)
    scale = 1920.0 / deck.width_emu
    total = escaped = 0
    for page in deck.pages:
        by_id = {}
        _index_shapes(by_id, page["shapes"])
        for u in hl_layout.build_units(hl_layout.page_shapes(page, deck)):
            shape = by_id.get(u.shape_id)
            if shape is None:
                continue
            sl, st = shape.left_emu * scale, shape.top_emu * scale
            sr = sl + shape.width_emu * scale
            sb = st + shape.height_emu * scale
            for r in u.lines:
                total += 1
                if (r.left_px < sl - 1 or r.top_px < st - 1
                        or r.left_px + r.width_px > sr + 1
                        or r.top_px + r.height_px > sb + 1):
                    escaped += 1
    assert total > 0
    assert escaped / total <= 0.05, f"溢出 {escaped}/{total} = {escaped / total:.1%}"


def _index_shapes(store, shapes):
    for shape in shapes:
        store[shape.shape_id] = shape
        _index_shapes(store, shape.children)


# ---------------------------------------------------------------- M3 丢弃出口

def test_dropped_shape_is_reported_via_collector():
    """契约 §4.3：inner_w_pt <= 0 要"记 warning"，不能无声无息地少一块高亮。"""
    narrow = _shape(shape_id=7, name="TooNarrow", width_emu=12700 * 10,
                    margin_left_emu=91440, margin_right_emu=91440,
                    paragraphs=[_para("放不下")])
    skipped: list = []
    units = hl_layout.build_units(_page([narrow]), skipped=skipped)
    assert units == []
    assert skipped == [{"page_index": 0, "shape_id": 7,
                        "shape_name": "TooNarrow", "reason": "no_inner_width"}]


def test_no_collector_means_no_crash_and_no_record():
    """不传收集器时行为与从前一致（只是没有出口），绝不因新增参数而抛。"""
    narrow = _shape(width_emu=12700 * 10, margin_left_emu=91440,
                    margin_right_emu=91440, paragraphs=[_para("放不下")])
    assert hl_layout.build_units(_page([narrow])) == []


def test_dropped_shape_reason_distinguishes_height():
    tall = _shape(height_emu=12700 * 5, margin_top_emu=45720,
                  margin_bottom_emu=45720, paragraphs=[_para("太扁")])
    skipped: list = []
    assert hl_layout.build_units(_page([tall]), skipped=skipped) == []
    assert skipped[0]["reason"] == "no_inner_height"


def test_healthy_shapes_are_not_reported():
    ok = _shape(paragraphs=[_para("正常")])
    skipped: list = []
    assert len(hl_layout.build_units(_page([ok]), skipped=skipped)) == 1
    assert skipped == []


def test_zero_size_table_cell_is_reported():
    from pptx_io import ShapeInfo
    sh = ShapeInfo(shape_id=9, name="T", kind="table", left_emu=0, top_emu=0,
                   width_emu=12700 * 100, height_emu=12700 * 100,
                   margin_left_emu=91440, margin_right_emu=91440,
                   table_cells=[[[_para("A1")]]],
                   table_col_widths_emu=[12700 * 10],
                   table_row_heights_emu=[12700 * 100])
    skipped: list = []
    assert hl_layout.build_units(_page([sh]), skipped=skipped) == []
    assert skipped[0]["reason"] == "zero_cell"


def test_short_column_array_is_reported():
    """同族静默丢弃（审计 L2）：列宽数组短于单元格矩阵时剩下的列会无声消失。"""
    from pptx_io import ShapeInfo
    sh = ShapeInfo(shape_id=10, name="T", kind="table", left_emu=0, top_emu=0,
                   width_emu=12700 * 300, height_emu=12700 * 100,
                   table_cells=[[[_para("A1")], [_para("B1")]]],
                   table_col_widths_emu=[12700 * 100],   # 只有 1 列，却给了 2 列单元格
                   table_row_heights_emu=[12700 * 100])
    skipped: list = []
    hl_layout.build_units(_page([sh]), skipped=skipped)
    assert "col_array_short" in [s["reason"] for s in skipped]


# ---------------------------------------------------------------- M1 硬换行

HARD_BREAK_CASES = [
    ("两行", "第一行\n第二行", 2),
    ("段尾换了行", "只有一行\n", 2),
    ("连续两个硬换行", "甲\n\n乙", 3),
    ("开头就是硬换行", "\n\nabc", 3),
    ("硬换行 + 软折行", "前缀字" * 20 + "\n后缀", None),
    ("行首带空格", "甲\n 乙", 2),
]


@pytest.mark.parametrize("label,text,expect", HARD_BREAK_CASES)
def test_hard_break_makes_qa_and_geometry_agree(label, text, expect):
    """修 M1：`a:br`（读成 "\\n"）是**硬换行**，三套度量必须给同一行数。

    契约 §4.2 的防漂移断言原本只盖 `wrap_lines`，而**真正产出几何**的是
    `_paragraph_lines` —— 这条用例把它也盖上了。
    """
    p_lines = hl_layout._paragraph_lines(text, 18.0, 300.0, True)
    w_lines = hl_layout.wrap_lines(text, 18.0, 300.0)
    q_lines = qa.measure_text_lines(text, 18.0, 300.0)
    assert len(p_lines) == len(w_lines) == q_lines, f"{label}: 三套度量不一致"
    if expect is not None:
        assert q_lines == expect, f"{label}: 期望 {expect} 行，实得 {q_lines}"


def test_qa_counts_hard_break_as_a_line():
    """`qa` 自己也必须认硬换行：它算的是"PowerPoint 会排出几行"，软换行算一个字形就错了。"""
    assert qa.measure_text_lines("甲\n乙", 18.0, 300.0) == 2
    assert qa.measure_text_lines("甲", 18.0, 300.0) == 1
    # 与"把两段分开算再相加"一致（这正是硬换行的语义）
    assert qa.measure_text_lines("甲\n乙", 18.0, 300.0) == (
        qa.measure_text_lines("甲", 18.0, 300.0)
        + qa.measure_text_lines("乙", 18.0, 300.0))


def test_hard_break_line_texts_are_split_not_glued():
    """硬换行把文本拆到不同行，而不是塞进同一行的文本里。"""
    lines = hl_layout.wrap_lines("甲\n乙", 18.0, 300.0)
    assert [ln.text for ln in lines] == ["甲", "乙"]


def test_hard_break_trailing_produces_empty_line():
    """PowerPoint 里 `a:br` 之后另起一行：段尾的硬换行会留下一个空行。"""
    lines = hl_layout.wrap_lines("甲\n", 18.0, 300.0)
    assert [ln.text for ln in lines] == ["甲", ""]
    assert lines[1].width_pt == 0.0
