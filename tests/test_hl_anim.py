"""M3：hl_anim.build_player 单测 —— 重点锁转义纪律与程序驱动 JS API。

背景（契约 §6.3）：title / 底图路径 / JSON 各自处在不同上下文（HTML 文本、src 属性、
<script> 字符串），转义方式不同；模板替换必须**单遍**，否则后一次 replace 会扫到
前一次插入的文本，title 里写 `__CONFIG__` 就能绕过转义注入活标签。
"""

import json
import os
import re

import pytest

import hl_anim
from hl_layout import Rect, Unit


def _rect(l, t, w, h):
    return Rect(int(l * 12700), int(t * 12700), int(w * 12700), int(h * 12700),
                float(l * 2), float(t * 2), float(w * 2), float(h * 2))


def _unit(text="要点一", kind="body", lines=None):
    lines = lines or [_rect(100, 50, 200, 40)]
    return Unit(order=0, page_index=0, text=text, kind=kind, rect=lines[0],
                lines=lines, size_pt=18.0, align="LEFT", shape_id=1,
                shape_name="S", is_estimated=False, warnings=[])


def _build(tmp_path, bg=None, pages=None, **kw):
    bg = bg if bg is not None else ["bg/slide_1.png"]
    pages = pages if pages is not None else [[_unit()]]
    path = hl_anim.build_player(str(tmp_path / "player"), bg, pages, **kw)
    return path, open(path, encoding="utf-8").read()


# ---------------------------------------------------------------- 基本结构

def test_writes_index_html(tmp_path):
    path, doc = _build(tmp_path, title="演示")
    assert path == os.path.join(str(tmp_path / "player"), "index.html")
    assert os.path.isfile(path)
    assert "演示" in doc


def test_zero_iframe(tmp_path):
    """契约 §6.2：底图 + 绝对定位 div，同源限制毫无收益 → 零 iframe。"""
    _, doc = _build(tmp_path)
    assert "<iframe" not in doc
    assert "sandbox" not in doc


def test_exposes_program_driven_api(tmp_path):
    """video --mode step 依赖这套 JS API。"""
    _, doc = _build(tmp_path)
    for name in ("goto", "next", "back", "state", "ready"):
        assert re.search(r"\b%s\b" % name, doc), f"缺 {name}"
    assert "window.hl" in doc


def test_goto_is_synchronous_no_transition(tmp_path):
    """goto 不许依赖 rAF/过渡：截图侧只靠 screenshot(animations=disabled) 兜底。

    底图靠预加载 + display 切换（不是换 src），高亮靠直接写 style。
    """
    _, doc = _build(tmp_path)
    goto_body = doc.split("goto(p, s)")[1].split("next()")[0]
    assert "requestAnimationFrame" not in goto_body
    assert "classList.toggle('on'" in doc  # 底图切页
    assert ".hl{position:absolute;pointer-events:none" in doc and "transition" not in doc.split(".hl{")[1].split("}")[0]


def test_bg_images_are_preloaded_and_stacked(tmp_path):
    """多张底图叠放 + display 切换，避免换 src 带来的异步加载。"""
    _, doc = _build(tmp_path, bg=["bg/slide_1.png", "bg/slide_2.png"],
                    pages=[[_unit()], [_unit()]])
    assert "BGS.map" in doc and "imgs.forEach" in doc


def test_units_are_embedded_as_json(tmp_path):
    _, doc = _build(tmp_path, pages=[[_unit(text="第一条"), _unit(text="第二条")]])
    m = re.search(r"const PAGES = (.*?);\n", doc)
    data = json.loads(m.group(1))
    assert data[0][0]["t"] == "第一条"
    assert len(data[0]) == 2
    assert data[0][0]["l"][0] == [200.0, 100.0, 400.0, 80.0]


def test_bg_paths_are_embedded(tmp_path):
    _, doc = _build(tmp_path, bg=["bg/slide_1.png"])
    m = re.search(r"const BGS = (.*?);\n", doc)
    assert json.loads(m.group(1)) == ["bg/slide_1.png"]


# ---------------------------------------------------------------- 转义

def test_title_with_script_tag_is_escaped(tmp_path):
    """title 含 </title><script> 不能变成活标签（审计 M2 的原型攻击）。"""
    payload = "</title><script>alert(1)</script>"
    _, doc = _build(tmp_path, title=payload)
    assert "<script>alert(1)</script>" not in doc
    assert "&lt;/title&gt;&lt;script&gt;alert(1)&lt;/script&gt;" in doc


def test_title_in_both_contexts_escaped(tmp_path):
    """title 同时进 <title> 与正文 h1 两处，都必须转义。"""
    _, doc = _build(tmp_path, title="<b>x</b>")
    assert doc.count("&lt;b&gt;x&lt;/b&gt;") == 2
    assert "<title>&lt;b&gt;x&lt;/b&gt; · 高亮讲解</title>" in doc


def test_double_underscore_placeholder_in_title_not_resubstituted(tmp_path):
    """单遍替换：title 里写 __CONFIG__ 不能被后续替换扫到（审计 M2）。"""
    _, doc = _build(tmp_path, title="__CONFIG__")
    m = re.search(r"const CFG = (.*?);\n", doc)
    cfg = json.loads(m.group(1))
    assert isinstance(cfg, dict) and cfg["canvasWidth"] == 1920
    assert "<title>__CONFIG__ · 高亮讲解</title>" in doc


def test_placeholder_in_unit_text_not_resubstituted(tmp_path):
    """单元文本里出现占位符字样也不能触发二次替换。"""
    _, doc = _build(tmp_path, pages=[[_unit(text="__UNITS__ 与 __TITLE__")]])
    m = re.search(r"const PAGES = (.*?);\n", doc)
    assert json.loads(m.group(1))[0][0]["t"] == "__UNITS__ 与 __TITLE__"


def test_bg_path_hash_and_percent_are_url_encoded(tmp_path):
    """文件名含 # 或 % ：裸拼 URL 会请求到错文件（# 被当 fragment，% 被当转义）。"""
    _, doc = _build(tmp_path, bg=["bg/slide_#1%2.png"])
    m = re.search(r"const BGS = (.*?);\n", doc)
    assert json.loads(m.group(1)) == ["bg/slide_%231%252.png"]
    assert 'src="bg/slide_#1%2.png"' not in doc


def test_bg_path_with_space_and_cjk_encoded(tmp_path):
    _, doc = _build(tmp_path, bg=["bg/第一页 稿.png"])
    m = re.search(r"const BGS = (.*?);\n", doc)
    assert json.loads(m.group(1))[0] == "bg/%E7%AC%AC%E4%B8%80%E9%A1%B5%20%E7%A8%BF.png"


def test_bg_path_with_quote_is_not_breakable(tmp_path):
    """路径含引号/尖括号也不能破出上下文：危险字符全部百分号编码。"""
    _, doc = _build(tmp_path, bg=['bg/a"><script>x</script>.png'])
    assert '<script>x</script>' not in doc
    encoded = json.loads(re.search(r"const BGS = (.*?);\n", doc).group(1))[0]
    assert encoded == "bg/a%22%3E%3Cscript%3Ex%3C/script%3E.png"
    assert not (set('<>"# ') & set(encoded))
    assert re.fullmatch(r"([^%]|%[0-9A-Fa-f]{2})+", encoded), "裸 % 会被当转义引导符"


def test_script_terminator_in_text_is_escaped(tmp_path):
    """单元文本含 </script> 不能提前闭合脚本块（JSON 上下文转义）。"""
    _, doc = _build(tmp_path, pages=[[_unit(text="危险</script><script>alert(2)")]])
    assert "<\\/script>" in doc
    assert "</script><script>alert(2)" not in doc


def test_html_comment_in_text_is_escaped(tmp_path):
    """<!-- 会让解析器进注释态使末尾 </script> 失效（审计 L4）。"""
    _, doc = _build(tmp_path, pages=[[_unit(text="<!-- 注释")]])
    assert "<\\!--" in doc


# ---------------------------------------------------------------- 参数校验

def test_page_count_mismatch_raises(tmp_path):
    from pptx_io import PptxError
    with pytest.raises(PptxError) as ei:
        hl_anim.build_player(str(tmp_path / "p"), ["bg/a.png"], [[], []])
    assert ei.value.code == "IR_MISMATCH"


@pytest.mark.parametrize("bad", [0, 0.6, -0.1, 1.0])
def test_dim_out_of_range_raises(tmp_path, bad):
    from pptx_io import PptxError
    with pytest.raises(PptxError) as ei:
        hl_anim.build_player(str(tmp_path / "p"), ["bg/a.png"], [[]], dim=bad)
    assert ei.value.code == "BAD_ARGS"


def test_dim_non_numeric_raises(tmp_path):
    from pptx_io import PptxError
    with pytest.raises(PptxError) as ei:
        hl_anim.build_player(str(tmp_path / "p"), ["bg/a.png"], [[]], dim="深一点")
    assert ei.value.code == "BAD_ARGS"


def test_auto_step_ms_must_be_positive(tmp_path):
    from pptx_io import PptxError
    with pytest.raises(PptxError) as ei:
        hl_anim.build_player(str(tmp_path / "p"), ["bg/a.png"], [[]], auto_step_ms=0)
    assert ei.value.code == "BAD_ARGS"


@pytest.mark.parametrize("bad", ["/etc/passwd.png", "../outside.png", "bg/../../x.png",
                                 "C:/tmp/x.png"])
def test_bg_path_escaping_out_dir_raises(tmp_path, bad):
    """deck.json 被改过时不能让它指向任意文件。"""
    from pptx_io import PptxError
    with pytest.raises(PptxError) as ei:
        hl_anim.build_player(str(tmp_path / "p"), [bad], [[]])
    assert ei.value.code == "IR_MISMATCH"


def test_absolute_bg_path_rejected(tmp_path):
    from pptx_io import PptxError
    with pytest.raises(PptxError) as ei:
        hl_anim.build_player(str(tmp_path / "p"), [os.path.abspath("x.png")], [[]])
    assert ei.value.code == "IR_MISMATCH"


def test_creates_out_dir(tmp_path):
    target = tmp_path / "deep" / "nested"
    path = hl_anim.build_player(str(target), ["bg/slide_1.png"], [[]])
    assert os.path.isfile(path)


def test_empty_deck_is_allowed(tmp_path):
    """零页不该崩（CLI 前置会拦，但纯函数不该抛）。"""
    path, doc = _build(tmp_path, bg=[], pages=[])
    assert os.path.isfile(path)
    assert json.loads(re.search(r"const PAGES = (.*?);\n", doc).group(1)) == []


# ---------------------------------------------------------------- 步进截图

def _mkbg(path, color):
    from PIL import Image
    os.makedirs(os.path.dirname(path), exist_ok=True)
    Image.new("RGB", (64, 36), color).save(path)


def _shot(player, out, steps, **kw):
    import hl_anim as m
    try:
        return m.shot_player(player, out, steps, **kw)
    except RuntimeError as e:  # 本机没有 Chrome/Edge
        pytest.skip(f"无可用浏览器：{e}")


def _canvas_player(tmp_path, colors=((200, 30, 30), (30, 200, 30))):
    """2 页、每页 1 个单元、底图为纯色的迷你播放器（画布 640×360，截图快）。"""
    out = tmp_path / "pl"
    bg = []
    for i, c in enumerate(colors, 1):
        rel = f"bg/slide_{i}.png"
        _mkbg(str(out / rel), c)
        bg.append(rel)
    pages = [[_unit(text=f"第{i}页要点")] for i in range(len(colors))]
    player = hl_anim.build_player(str(out), bg, pages, canvas_width_px=640,
                                  canvas_height_px=360)
    return player, out


def test_shot_player_rejects_empty_steps(tmp_path):
    from pptx_io import PptxError
    with pytest.raises(PptxError) as ei:
        hl_anim.shot_player(str(tmp_path / "x.html"), str(tmp_path / "o"), [])
    assert ei.value.code == "BAD_ARGS"


def test_shot_player_rejects_missing_player_html(tmp_path):
    from pptx_io import PptxError
    with pytest.raises(PptxError) as ei:
        hl_anim.shot_player(str(tmp_path / "nope.html"), str(tmp_path / "o"), [(0, 1)])
    assert ei.value.code == "IR_MISMATCH"


def test_shot_player_one_png_per_step(tmp_path):
    """验收：步进序列截图数 == len(steps)，尺寸 == 画布，文件非空白。"""
    from PIL import Image
    player, out = _canvas_player(tmp_path)
    shots = _shot(player, str(tmp_path / "frames"),
                  [(0, 1), (0, 2), (1, 1), (1, 2), (1, 3)])
    assert len(shots) == 5
    assert [os.path.basename(s) for s in shots] == [
        "step_0001.png", "step_0002.png", "step_0003.png", "step_0004.png", "step_0005.png"]
    for s in shots:
        assert os.path.getsize(s) > 0
        with Image.open(s) as im:
            assert im.size == (640, 360)
            # 非空白：画面主色应当是该页底图色（纯色块）
            assert sum(im.convert("RGB").getpixel((20, 20))) > 60


def test_shot_player_frames_match_requested_page(tmp_path):
    """第 0 页与第 1 页截出来的底色不同 —— 证明 goto(page) 真的切了底图。"""
    from PIL import Image
    player, out = _canvas_player(tmp_path)
    shots = _shot(player, str(tmp_path / "f2"), [(0, 1), (1, 1)])
    with Image.open(shots[0]) as a, Image.open(shots[1]) as b:
        pa = a.convert("RGB").getpixel((600, 20))
        pb = b.convert("RGB").getpixel((600, 20))
    assert pa[0] > pa[1] and pb[1] > pb[0], f"两页底色没区分开：{pa} vs {pb}"


def test_shot_player_missing_bg_raises_not_black_frame(tmp_path):
    """审计 L4：缺底图时播放器会安静地出黑帧；shot_player 必须先报错。"""
    from pptx_io import PptxError
    out = tmp_path / "pl"
    _mkbg(str(out / "bg" / "slide_1.png"), (10, 10, 200))
    player = hl_anim.build_player(str(out), ["bg/slide_1.png", "bg/nope.png"],
                                  [[_unit()], [_unit()]],
                                  canvas_width_px=320, canvas_height_px=180)
    with pytest.raises(PptxError) as ei:
        hl_anim.shot_player(player, str(tmp_path / "f3"), [(0, 1)])
    assert ei.value.code == "IR_MISMATCH"
    assert "底图" in ei.value.message
