"""模板模块（内置库 + 参考稿识别）与路由的单元测试。mock LLM，不发真实请求。"""

import io
import json

import pytest

import app as app_mod
import style as style_mod
import template as tmpl


# ===== 内置模板库 =====

def test_builtin_templates_cover_all_styles():
    bt = tmpl.builtin_templates()
    assert len(bt) == len(style_mod.STYLE_LIBRARY)
    keys = {t["key"] for t in bt}
    assert keys == set(style_mod.STYLE_LIBRARY)
    for t in bt:
        assert t["accent"].startswith("#") and t["bg"].startswith("#")
        assert t["name"] and "layout_hint" in t


def test_get_template_style_merges_layout_hint():
    st = tmpl.get_template_style("business-minimal")
    assert st["key"] == "business-minimal"
    assert "版式骨架" in st["guidance"]
    assert st["accent"] == style_mod.STYLE_LIBRARY["business-minimal"]["accent"]


def test_get_template_style_unknown_returns_none():
    assert tmpl.get_template_style("no-such") is None


# ===== _sanitize 容错 =====

def test_sanitize_keeps_valid_hex():
    out = tmpl._sanitize({"bg": "#101820", "accent": "#FF8800", "fg": "#ffffff",
                          "muted": "#888888", "name": "测试", "guidance": "g"})
    assert out["accent"] == "#FF8800"
    assert out["key"] == "custom"


def test_sanitize_rejects_bad_hex_falls_back():
    out = tmpl._sanitize({"bg": "blue", "accent": "rgb(1,2,3)", "name": ""})
    fresh = style_mod.STYLE_LIBRARY["fresh-light"]
    assert out["accent"] == fresh["accent"]      # 非法 → 回退
    assert out["bg"] == fresh["bg"]
    assert out["name"] == "参考稿风格"           # 空名 → 默认


def test_sanitize_truncates_long_fields():
    out = tmpl._sanitize({"name": "名" * 50, "guidance": "很长" * 200})
    assert len(out["name"]) <= 20
    assert len(out["guidance"]) <= 200


# ===== _extract_json_obj =====

def test_extract_json_obj_normal():
    assert tmpl._extract_json_obj('前言{"accent":"#ff0000"}尾巴')["accent"] == "#ff0000"


def test_extract_json_obj_hex_fallback():
    obj = tmpl._extract_json_obj("主色 #00ff88 背景 #111111")
    assert obj["accent"] == "#00ff88" and obj["bg"] == "#111111"


def test_extract_json_obj_nothing():
    assert tmpl._extract_json_obj("毫无颜色信息") is None


# ===== analyze_reference 失败降级 =====

def test_analyze_reference_returns_none_on_error(monkeypatch):
    def boom(*a, **k):
        raise RuntimeError("网络挂了")
    monkeypatch.setattr(tmpl, "_client", boom)
    assert tmpl.analyze_reference("image", b"\x89PNG") is None


def test_analyze_reference_html_parses(monkeypatch):
    class FakeResp:
        class choices:
            pass
    fake = FakeResp()
    msg = type("M", (), {"content": '{"name":"参考","accent":"#123456","bg":"#000000","fg":"#ffffff","muted":"#999999"}'})
    fake.choices = [type("C", (), {"message": msg})]
    monkeypatch.setattr(tmpl, "_client",
                        lambda: type("Cli", (), {"chat": type("Ch", (), {
                            "completions": type("Co", (), {
                                "create": staticmethod(lambda **k: fake)})()})()})())
    out = tmpl.analyze_reference("html", "<style>body{color:#123456}</style>" + "x" * 40)
    assert out and out["accent"] == "#123456" and out["key"] == "custom"


# ===== 路由 =====

@pytest.fixture
def client():
    app_mod.app.config["TESTING"] = True
    return app_mod.app.test_client()


@pytest.fixture(autouse=True)
def restore():
    before = {k: v for k, v in app_mod.state.items()}
    yield
    app_mod.state.clear()
    app_mod.state.update(before)


def test_api_templates_lists(client):
    d = client.get("/api/templates").get_json()
    assert len(d["templates"]) == len(style_mod.STYLE_LIBRARY)


def test_api_template_select_sets_and_clears(client):
    app_mod.state["phase"] = "ready"
    r = client.post("/api/template/select", json={"key": "tech-deep"})
    assert r.status_code == 200
    assert app_mod.state["tpl_style"]["key"] == "tech-deep"
    assert app_mod.state["tpl_style_name"] == r.get_json()["tpl_style_name"]
    # 清除
    r2 = client.post("/api/template/select", json={"key": None})
    assert r2.status_code == 200
    assert app_mod.state["tpl_style"] is None


def test_api_template_select_unknown_404(client):
    app_mod.state["phase"] = "ready"
    assert client.post("/api/template/select", json={"key": "xxx"}).status_code == 404


def test_api_template_select_busy_409(client):
    app_mod.state["phase"] = "images"
    assert client.post("/api/template/select", json={"key": "tech-deep"}).status_code == 409


def test_api_template_analyze_html(client, monkeypatch):
    app_mod.state["phase"] = "ready"
    monkeypatch.setattr(tmpl, "analyze_reference",
                        lambda kind, data: {"key": "custom", "name": "参考稿风格",
                                            "bg": "#101820", "accent": "#22d3ee",
                                            "fg": "#ffffff", "muted": "#94a3b8",
                                            "mood": "科技", "guidance": "g"})
    r = client.post("/api/template/analyze", json={"html": "<style>x</style>" + "y" * 40})
    assert r.status_code == 200
    assert app_mod.state["tpl_style"]["accent"] == "#22d3ee"
    assert r.get_json()["accent"] == "#22d3ee"


def test_api_template_analyze_short_html_400(client):
    app_mod.state["phase"] = "ready"
    assert client.post("/api/template/analyze", json={"html": "太短"}).status_code == 400


def test_api_template_analyze_unrecognized_422(client, monkeypatch):
    app_mod.state["phase"] = "ready"
    monkeypatch.setattr(tmpl, "analyze_reference", lambda kind, data: None)
    r = client.post("/api/template/analyze", json={"html": "<html>" + "x" * 40 + "</html>"})
    assert r.status_code == 422


def test_status_exposes_tpl_style_name(client):
    app_mod.state["phase"] = "ready"
    client.post("/api/template/select", json={"key": "nature-green"})
    s = client.get("/api/status").get_json()
    assert "tpl_style_name" in s and s["tpl_style_name"]
