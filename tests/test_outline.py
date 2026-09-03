"""outline 模块纯逻辑单元测试（离线，不调真实 API）。"""

import pytest

from outline import _extract_json, _normalize, VALID_TYPES


# ---------- _extract_json ----------

def test_extract_json_plain_array():
    text = '[{"type": "content", "title": "标题", "points": ["a"], "image_prompt": "p"}]'
    assert _extract_json(text) == [
        {"type": "content", "title": "标题", "points": ["a"], "image_prompt": "p"}
    ]


def test_extract_json_code_block():
    text = '```json\n[{"type": "cover", "title": "T"}]\n```'
    assert _extract_json(text) == [{"type": "cover", "title": "T"}]


def test_extract_json_missing_brackets_raises():
    with pytest.raises(ValueError):
        _extract_json("这是纯文本，没有 JSON 数组")


def test_extract_json_empty_string_raises():
    with pytest.raises(ValueError):
        _extract_json("")


# ---------- _normalize ----------

def test_normalize_invalid_type_falls_back_to_content():
    slides = [
        {"type": "cover", "title": "T", "points": [], "image_prompt": "p"},
        {"type": "weird", "title": "X", "points": ["a"], "image_prompt": ""},
    ]
    out = _normalize(slides)
    assert out[0]["type"] == "cover"
    assert out[1]["type"] == "content"


def test_normalize_first_page_no_points_forced_cover():
    # 首页是 toc 但无 points → 强制 cover
    out = _normalize([{"type": "toc", "title": "目录", "points": [], "image_prompt": ""}])
    assert out[0]["type"] == "cover"


def test_normalize_first_page_content_forced_cover():
    # 首页 type=content 即使有 points 也强制 cover
    out = _normalize([{"type": "content", "title": "主题", "points": ["a"], "image_prompt": ""}])
    assert out[0]["type"] == "cover"


def test_normalize_cover_content_missing_prompt_fallback():
    out = _normalize([
        {"type": "cover", "title": "封面标题", "points": [], "image_prompt": ""},
        {"type": "content", "title": "内容标题", "points": ["a"], "image_prompt": ""},
    ])
    assert out[0]["image_prompt"] == "封面标题，扁平插画风格"
    # content 页无图时保持空（走无图布局，如卡片/两栏/居中）
    assert out[1]["image_prompt"] == ""


def test_normalize_keeps_present_prompt():
    out = _normalize([{"type": "content", "title": "T", "points": ["a"], "image_prompt": "已有画面"}])
    assert out[0]["image_prompt"] == "已有画面"


def test_normalize_chart_valid_kept():
    chart = {"type": "pie", "labels": ["A", "B"], "values": [60, 40]}
    out = _normalize([
        {"type": "cover", "title": "T", "points": [], "image_prompt": "p"},
        {"type": "data", "title": "D", "points": [], "image_prompt": "", "chart": chart},
    ])
    assert out[1]["chart"] == chart


def test_normalize_chart_labels_not_list_dropped():
    out = _normalize([
        {"type": "cover", "title": "T", "points": [], "image_prompt": "p"},
        {"type": "data", "title": "D", "points": [], "image_prompt": "",
         "chart": {"labels": "A", "values": [1]}},
    ])
    assert out[1]["chart"] is None


def test_normalize_chart_values_not_list_dropped():
    out = _normalize([
        {"type": "cover", "title": "T", "points": [], "image_prompt": "p"},
        {"type": "data", "title": "D", "points": [], "image_prompt": "",
         "chart": {"labels": ["A"], "values": "1"}},
    ])
    assert out[1]["chart"] is None


def test_normalize_chart_not_dict_dropped():
    out = _normalize([
        {"type": "cover", "title": "T", "points": [], "image_prompt": "p"},
        {"type": "data", "title": "D", "points": [], "image_prompt": "", "chart": "pie"},
    ])
    assert out[1]["chart"] is None


def test_normalize_no_chart_field_is_none():
    out = _normalize([
        {"type": "cover", "title": "T", "points": [], "image_prompt": "p"},
        {"type": "content", "title": "C", "points": ["a"], "image_prompt": "p"},
    ])
    assert out[1]["chart"] is None


def test_normalize_skips_non_dict_entries():
    out = _normalize(["不是字典", {"type": "content", "title": "C", "points": ["a"], "image_prompt": "p"}])
    assert len(out) == 1
    assert out[0]["title"] == "C"


def test_valid_types_set():
    assert VALID_TYPES == {"cover", "toc", "section", "content", "data", "timeline", "compare", "end"}
