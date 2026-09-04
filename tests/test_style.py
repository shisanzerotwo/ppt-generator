"""style 单测：AI 选风格的解析、回退与 prompt 注入（不调真实 LLM）。"""

import html_gen
import style


SLIDES = [
    {"type": "cover", "title": "人工智能入门", "points": []},
    {"type": "content", "title": "什么是机器学习", "points": ["a", "b"]},
]


def test_decide_style_valid_key(monkeypatch):
    monkeypatch.setattr(style, "_call_llm", lambda p: '{"key": "tech-deep", "reason": "AI 主题适合科技风"}')
    s = style.decide_style("人工智能入门", SLIDES)
    assert s["key"] == "tech-deep"
    assert s["name"] == "深空科技"
    assert "科技" in s["reason"]


def test_decide_style_invalid_key_falls_back(monkeypatch):
    monkeypatch.setattr(style, "_call_llm", lambda p: '{"key": "不存在的", "reason": "x"}')
    s = style.decide_style("随便", SLIDES)
    assert s["key"] == "fresh-light"


def test_decide_style_llm_error_falls_back(monkeypatch):
    def boom(p):
        raise RuntimeError("api down")

    monkeypatch.setattr(style, "_call_llm", boom)
    s = style.decide_style("随便", SLIDES)
    assert s["key"] == "fresh-light"


def test_decide_style_garbage_output_falls_back(monkeypatch):
    monkeypatch.setattr(style, "_call_llm", lambda p: "我不会 JSON")
    s = style.decide_style("随便", SLIDES)
    assert s["key"] == "fresh-light"


def test_build_prompt_contains_library_and_topic():
    p = style.build_prompt("青少年如何学习", SLIDES)
    for key in style.STYLE_LIBRARY:
        assert key in p
    assert "青少年如何学习" in p


def test_style_guidance_injected_into_html_prompt():
    chosen = dict(style.STYLE_LIBRARY["tech-deep"], key="tech-deep", reason="r")
    prompt = html_gen._build_prompt("AI 主题", [], {}, chosen)
    assert "深空科技" in prompt
    assert "#38bdf8" in prompt
    assert "【大纲数据】" in prompt  # 注入不破坏原结构


def test_html_prompt_without_style_unchanged():
    prompt = html_gen._build_prompt("AI 主题", [], {}, None)
    assert "设计基调" not in prompt
