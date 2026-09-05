"""/api/theme 即时换色：变量替换、非法输入拒收、旧稿 409。"""

import pytest

import app as app_mod


@pytest.fixture
def client():
    return app_mod.app.test_client()


@pytest.fixture
def deck_with_vars(tmp_path, monkeypatch):
    deck = tmp_path / "t.html"
    deck.write_text(
        "<style>:root{--bg:#0f2a4a;--fg:#ffffff;--accent:#3b82f6;--muted:#b6c7dc;}</style>",
        encoding="utf-8",
    )
    monkeypatch.setattr(app_mod, "DECKS_DIR", str(tmp_path))
    with app_mod.lock:
        app_mod.state["phase"] = "ready"
        app_mod.state["html_path"] = "/decks/t.html"
    yield deck
    with app_mod.lock:
        app_mod.state["phase"] = "idle"
        app_mod.state["html_path"] = None


def test_theme_rewrites_root_vars(deck_with_vars, client):
    resp = client.post("/api/theme", json={"accent": "#ff0000"})
    assert resp.status_code == 200
    text = deck_with_vars.read_text(encoding="utf-8")
    assert "--accent:#ff0000" in text
    # 其它变量不受影响
    assert "--bg:#0f2a4a" in text


def test_theme_rejects_non_hex(deck_with_vars, client):
    resp = client.post("/api/theme", json={"accent": "<script>alert(1)</script>"})
    assert resp.status_code == 400
    assert deck_with_vars.read_text(encoding="utf-8").count("#3b82f6") == 1  # 文件未被触碰


def test_theme_deck_without_vars_409(tmp_path, monkeypatch, client):
    deck = tmp_path / "old.html"
    deck.write_text("<style>body{background:#123456}</style>", encoding="utf-8")
    monkeypatch.setattr(app_mod, "DECKS_DIR", str(tmp_path))
    with app_mod.lock:
        app_mod.state["phase"] = "ready"
        app_mod.state["html_path"] = "/decks/old.html"
    try:
        resp = client.post("/api/theme", json={"accent": "#ff0000"})
        assert resp.status_code == 409
    finally:
        with app_mod.lock:
            app_mod.state["phase"] = "idle"
            app_mod.state["html_path"] = None


def test_theme_requires_ready_phase(deck_with_vars, client):
    with app_mod.lock:
        app_mod.state["phase"] = "images"
    try:
        resp = client.post("/api/theme", json={"accent": "#ff0000"})
        assert resp.status_code == 409
    finally:
        with app_mod.lock:
            app_mod.state["phase"] = "ready"
