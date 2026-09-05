"""定点改写 critic.revise_target 的单元测试（mock LLM，不发真实请求）。"""

import copy
import json

import pytest

import critic


def _slides():
    return [
        {"type": "cover", "title": "封面", "points": [], "image_prompt": "封面画面",
         "chart": None, "layout": None, "image": "/images/slide_0.png",
         "imageStatus": "done", "review": {"ok": True, "reason": "", "tries": 0}},
        {"type": "content", "title": "第二页", "points": ["要点甲：说明一", "要点乙：说明二"],
         "image_prompt": "原画面", "chart": None, "layout": "image-right",
         "image": "/images/slide_1.png", "imageStatus": "done",
         "review": {"ok": True, "reason": "", "tries": 0}},
        {"type": "data", "title": "数据页", "points": [], "image_prompt": "",
         "chart": {"type": "pie", "labels": ["A", "B"], "values": [60, 40]},
         "layout": None, "image": None, "imageStatus": "skipped",
         "review": {"ok": True, "reason": "", "tries": 0}},
    ]


# ===== 要点级 =====

def test_point_level_replaces_only_matched_quote(monkeypatch):
    monkeypatch.setattr(critic, "_ask", lambda p: "要点甲：改写后的说明")
    out, info = critic.revise_target(_slides(), {"slide": 1, "quote": "要点甲"}, "更简洁")
    assert info == {"level": "point", "slide": 1, "regen_image": False}
    assert out[1]["points"] == ["要点甲：改写后的说明", "要点乙：说明二"]  # 另一条不动


def test_point_level_keeps_other_slides_untouched(monkeypatch):
    before = _slides()
    monkeypatch.setattr(critic, "_ask", lambda p: "新文本：x")
    out, _ = critic.revise_target(before, {"slide": 1, "quote": "要点乙"}, "改")
    assert out[0] == before[0]
    assert out[2] == before[2]
    # 状态字段（图/校验）原样保留，未被清空
    assert out[1]["image"] == "/images/slide_1.png"
    assert out[1]["imageStatus"] == "done"


def test_point_level_does_not_mutate_input(monkeypatch):
    src = _slides()
    snapshot = copy.deepcopy(src)
    monkeypatch.setattr(critic, "_ask", lambda p: "改过了：y")
    critic.revise_target(src, {"slide": 1, "quote": "要点甲"}, "改")
    assert src == snapshot


def test_point_level_strips_extra_lines_and_quotes(monkeypatch):
    """LLM 常附加解释或引号，只取首行并剥壳。"""
    monkeypatch.setattr(critic, "_ask",
                        lambda p: '  "要点甲：正确内容"\n\n以上是改写结果。')
    out, _ = critic.revise_target(_slides(), {"slide": 1, "quote": "要点甲"}, "改")
    assert out[1]["points"][0] == "要点甲：正确内容"


def test_point_level_quote_not_found_raises(monkeypatch):
    monkeypatch.setattr(critic, "_ask", lambda p: "x")
    with pytest.raises(ValueError):
        critic.revise_target(_slides(), {"slide": 1, "quote": "根本不存在的文本"}, "改")


def test_point_level_empty_llm_output_raises(monkeypatch):
    monkeypatch.setattr(critic, "_ask", lambda p: "   ")
    with pytest.raises(ValueError):
        critic.revise_target(_slides(), {"slide": 1, "quote": "要点甲"}, "改")


# ===== 页级 =====

def test_slide_level_replaces_content_fields(monkeypatch):
    payload = [{"type": "content", "title": "新标题", "points": ["新要点：说明"],
                "image_prompt": "新画面"}]
    monkeypatch.setattr(critic, "_ask", lambda p: json.dumps(payload, ensure_ascii=False))
    out, info = critic.revise_target(_slides(), {"slide": 1}, "整页重写")
    assert info["level"] == "slide"
    assert out[1]["title"] == "新标题"
    assert out[1]["points"] == ["新要点：说明"]
    assert out[1]["image_prompt"] == "新画面"


def test_slide_level_prompt_change_flags_regen(monkeypatch):
    payload = [{"type": "content", "title": "T", "points": ["p"], "image_prompt": "完全不同的画面"}]
    monkeypatch.setattr(critic, "_ask", lambda p: json.dumps(payload, ensure_ascii=False))
    _, info = critic.revise_target(_slides(), {"slide": 1}, "改")
    assert info["regen_image"] is True


def test_slide_level_unchanged_prompt_no_regen(monkeypatch):
    payload = [{"type": "content", "title": "T", "points": ["p"], "image_prompt": "原画面"}]
    monkeypatch.setattr(critic, "_ask", lambda p: json.dumps(payload, ensure_ascii=False))
    _, info = critic.revise_target(_slides(), {"slide": 1}, "改")
    assert info["regen_image"] is False


def test_slide_level_forces_type_unchanged(monkeypatch):
    """LLM 若擅自改 type，服务端强制回写为原版式。"""
    payload = [{"type": "end", "title": "T", "points": ["p"], "image_prompt": "原画面"}]
    monkeypatch.setattr(critic, "_ask", lambda p: json.dumps(payload, ensure_ascii=False))
    out, _ = critic.revise_target(_slides(), {"slide": 1}, "改")
    assert out[1]["type"] == "content"


def test_slide_level_keeps_chart_when_llm_drops_it(monkeypatch):
    payload = [{"type": "data", "title": "数据页", "points": ["说明"], "image_prompt": ""}]
    monkeypatch.setattr(critic, "_ask", lambda p: json.dumps(payload, ensure_ascii=False))
    out, _ = critic.revise_target(_slides(), {"slide": 2}, "加点说明")
    assert out[2]["chart"] == {"type": "pie", "labels": ["A", "B"], "values": [60, 40]}


def test_slide_level_blank_prompt_falls_back_to_old(monkeypatch):
    payload = [{"type": "content", "title": "T", "points": ["p"], "image_prompt": ""}]
    monkeypatch.setattr(critic, "_ask", lambda p: json.dumps(payload, ensure_ascii=False))
    out, info = critic.revise_target(_slides(), {"slide": 1}, "改")
    assert out[1]["image_prompt"] == "原画面"      # 不被 _normalize 的兜底污染
    assert info["regen_image"] is False           # 且不该白白重生图


def test_slide_level_adds_image_when_originally_none(monkeypatch):
    """原本无图的页（toc/section 等），改写后 LLM 给了画面描述 → 应重生配图。"""
    payload = [{"type": "section", "title": "新章节", "points": ["p"], "image_prompt": "新配图描述"}]
    monkeypatch.setattr(critic, "_ask", lambda p: json.dumps(payload, ensure_ascii=False))
    out, info = critic.revise_target(_slides(), {"slide": 2}, "改")
    assert info["regen_image"] is True
    assert out[2]["image_prompt"] == "新配图描述"


# ===== 参数校验 =====

@pytest.mark.parametrize("bad", [-1, 99, "1", None])
def test_invalid_slide_index_raises(bad):
    with pytest.raises(ValueError):
        critic.revise_target(_slides(), {"slide": bad}, "改")


def test_llm_garbage_raises(monkeypatch):
    monkeypatch.setattr(critic, "_ask", lambda p: "这不是 JSON")
    with pytest.raises(Exception):
        critic.revise_target(_slides(), {"slide": 1}, "改")


def test_prompt_carries_only_one_point_constraint(monkeypatch):
    """要点级 prompt 必须带「只输出这一条要点」约束，否则 LLM 易返回整页。"""
    seen = {}
    monkeypatch.setattr(critic, "_ask", lambda p: seen.update(p=p) or "新要点：y")
    critic.revise_target(_slides(), {"slide": 1, "quote": "要点甲"}, "改")
    assert "这一条要点" in seen["p"]
    assert "只输出" in seen["p"]
