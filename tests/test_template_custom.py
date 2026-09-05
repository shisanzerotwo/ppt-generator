"""自定义模板库：保存 / 列表合并 / 选中可用 / 删除 / :root 风格提取 / 路由接线。"""

import pytest

import app as app_mod
import template as template_mod


@pytest.fixture
def client():
    return app_mod.app.test_client()


@pytest.fixture
def custom_dir(tmp_path, monkeypatch):
    monkeypatch.setattr(template_mod, "CUSTOM_DIR", str(tmp_path))
    return tmp_path


SAMPLE_HTML = ('<style>:root{--bg:#0f2a4a;--fg:#ffffff;--accent:#ff5500;'
               '--muted:#b6c7dc;}.slide{color:var(--fg)}</style>')


def test_extract_style_from_root_vars():
    style = template_mod.extract_style_from_html(SAMPLE_HTML, "我的模板")
    assert style["accent"] == "#ff5500" and style["bg"] == "#0f2a4a"
    assert style["name"] == "我的模板" and style["key"] == "custom"  # _sanitize 产物


def test_extract_rejects_html_without_vars():
    assert template_mod.extract_style_from_html("<style>body{color:red}</style>") is None


def test_save_list_delete_roundtrip(custom_dir):
    style = template_mod.extract_style_from_html(SAMPLE_HTML, "我的模板")
    entry = template_mod.save_custom_template(style)
    assert entry["key"].startswith("custom:") and entry["source"] == "custom"
    assert any(t["key"] == entry["key"] for t in template_mod.custom_templates())
    # 选中可用：get_template_style 认得自定义 key（app.py /api/template/select 无需改动的原因）
    loaded = template_mod.get_template_style(entry["key"])
    assert loaded and loaded["accent"] == "#ff5500"
    assert template_mod.delete_custom_template(entry["key"])
    assert template_mod.custom_templates() == []
    assert not template_mod.delete_custom_template(entry["key"])  # 重复删除 False


def test_custom_key_blocks_path_escape(custom_dir):
    # key 只放行 8 位十六进制，路径拼接无法逃逸
    assert template_mod.get_template_style("custom:../../../app") is None
    assert template_mod.delete_custom_template("custom:..\\..\\x") is False


def test_routes_save_list_delete(custom_dir, tmp_path, monkeypatch, client):
    decks = tmp_path / "decks"
    decks.mkdir()
    (decks / "t.html").write_text(SAMPLE_HTML, encoding="utf-8")
    monkeypatch.setattr(app_mod, "DECKS_DIR", str(decks))
    with app_mod.lock:
        app_mod.state.update({"phase": "ready", "html_path": "/decks/t.html",
                              "tpl_style": None, "tpl_style_name": ""})
    try:
        resp = client.post("/api/templates/save", json={"name": "我的模板"})
        entry = resp.get_json()["template"]
        assert resp.status_code == 200 and entry["source"] == "custom"

        listed = client.get("/api/templates").get_json()["templates"]
        assert any(t["key"] == entry["key"] for t in listed)          # 与内置库合并展示
        builtin_keys = {t["key"] for t in template_mod.builtin_templates()}
        assert all(t["key"] in builtin_keys or t["source"] == "custom" for t in listed)

        # 选中 → 删除时正在选用 → 应一并取消
        client.post("/api/template/select", json={"key": entry["key"]})
        assert client.post("/api/templates/delete", json={"key": entry["key"]}).status_code == 200
        with app_mod.lock:
            assert app_mod.state["tpl_style"] is None
    finally:
        with app_mod.lock:
            app_mod.state.update({"phase": "idle", "html_path": None})


def test_save_requires_deck(client):
    with app_mod.lock:
        app_mod.state.update({"phase": "idle", "html_path": None})
    resp = client.post("/api/templates/save", json={"name": "x"})
    assert resp.status_code == 409
