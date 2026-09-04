"""html_gen 单测：不调真实 LLM，monkeypatch _call_llm。"""

import json

import html_gen
import pytest


SLIDES = [
    {"type": "cover", "title": "测试主题", "points": [], "image_prompt": "", "chart": None, "layout": None},
    {"type": "data", "title": "数据页", "points": [], "image_prompt": "",
     "chart": {"type": "bar", "labels": ["甲", "乙"], "values": [3, 5]}, "layout": None},
    {"type": "content", "title": "要点页", "points": ["要点一", "要点二"], "image_prompt": "p",
     "chart": None, "layout": "image-right"},
]

GOOD_HTML = '<!DOCTYPE html><html lang="zh"><head><style>.slide{}</style></head><body><section class="slide"></section></body></html>'


def test_extract_html_normal():
    text = f"说明文字\n{GOOD_HTML}\n尾部杂文"
    assert html_gen._extract_html(text) == GOOD_HTML


def test_extract_html_missing_doctype():
    with pytest.raises(ValueError, match="DOCTYPE"):
        html_gen._extract_html("<html></html>")


def test_extract_html_truncated():
    text = '<!DOCTYPE html><html><body><section class="slide">未闭合'
    with pytest.raises(ValueError, match="截断"):
        html_gen._extract_html(text)


def test_build_prompt_contains_key_info():
    prompt = html_gen._build_prompt("测试主题", SLIDES, {0: "../images/slide_0.png"})
    assert "测试主题" in prompt
    assert "../images/slide_0.png" in prompt
    # 无图页不写 image 字段
    data = json.loads(prompt[prompt.find("["):prompt.rfind("]") + 1])
    assert data[0]["image"] == "../images/slide_0.png"
    assert "image" not in data[1]
    assert data[1]["chart"]["values"] == [3, 5]


def test_generate_success(monkeypatch):
    calls = []

    def fake_call(prompt):
        calls.append(prompt)
        return GOOD_HTML

    monkeypatch.setattr(html_gen, "_call_llm", fake_call)
    out = html_gen.generate_html_deck("测试主题", SLIDES, {0: "../images/slide_0.png"})
    assert out == GOOD_HTML
    assert len(calls) == 1


def test_generate_retry_on_truncation(monkeypatch):
    truncated = '<!DOCTYPE html><html><body>被截断'
    attempts = []

    def fake_call(prompt):
        attempts.append(1)
        return truncated if len(attempts) == 1 else GOOD_HTML

    monkeypatch.setattr(html_gen, "_call_llm", fake_call)
    out = html_gen.generate_html_deck("测试主题", SLIDES, {})
    assert out == GOOD_HTML
    assert len(attempts) == 2  # 截断后重试了一次


def test_generate_fail_after_retry(monkeypatch):
    def fake_call(prompt):
        return "完全不是 HTML 的回复"

    monkeypatch.setattr(html_gen, "_call_llm", fake_call)
    with pytest.raises(RuntimeError, match="重试"):
        html_gen.generate_html_deck("测试主题", SLIDES, {})
