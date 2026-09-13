"""可视化制作路由：在 PowerPoint 中打开 / 版式切换 / 逐页缩略图。"""
import os
import time

import pytest

import app as app_mod


@pytest.fixture
def client():
    return app_mod.app.test_client()


@pytest.fixture
def vis_env(tmp_path, monkeypatch):
    """ready 态 + tmp 隔离的 OUTPUT_DIR / DECKS_DIR / THUMBNAILS_DIR，结束后恢复 state。"""
    decks = tmp_path / "decks"
    decks.mkdir()
    monkeypatch.setattr(app_mod, "OUTPUT_DIR", str(tmp_path))
    monkeypatch.setattr(app_mod, "DECKS_DIR", str(decks))
    monkeypatch.setattr(app_mod, "THUMBNAILS_DIR", str(tmp_path / "thumbnails"))
    with app_mod.lock:
        app_mod.state["phase"] = "ready"
        app_mod.state["slides"] = [
            {"type": "content", "title": "页一", "points": ["a"], "layout": None,
             "image": None, "imageStatus": "skipped"}]
        app_mod.state["html_path"] = None
    yield tmp_path
    with app_mod.lock:
        app_mod.state["phase"] = "idle"
        app_mod.state["slides"] = []
        app_mod.state["html_path"] = None


# ---------------- 在 PowerPoint 中打开 ----------------

def test_open_pptx_calls_startfile(vis_env, client, monkeypatch):
    (vis_env / "deck_a.pptx").write_bytes(b"PK")
    calls = []
    monkeypatch.setattr(os, "startfile", lambda p: calls.append(p))
    r = client.post("/api/open_in_powerpoint", json={"name": "deck_a.pptx"})
    assert r.status_code == 200 and r.get_json()["ok"] is True
    assert calls and calls[0].endswith("deck_a.pptx")


def test_open_rejects_non_pptx(vis_env, client):
    r = client.post("/api/open_in_powerpoint", json={"name": "deck_a.pdf"})
    assert r.status_code == 400


def test_open_missing_file(vis_env, client):
    r = client.post("/api/open_in_powerpoint", json={"name": "nope.pptx"})
    assert r.status_code == 404


def test_open_path_traversal_is_baselined(vis_env, client):
    r = client.post("/api/open_in_powerpoint", json={"name": "../../evil.pptx"})
    # basename 化后只剩 evil.pptx，文件不存在 → 404，绝逃不出 OUTPUT_DIR
    assert r.status_code == 404


def test_open_startfile_oserror(vis_env, client, monkeypatch):
    (vis_env / "deck_b.pptx").write_bytes(b"PK")

    def boom(_):
        raise OSError("no association")

    monkeypatch.setattr(os, "startfile", boom)
    r = client.post("/api/open_in_powerpoint", json={"name": "deck_b.pptx"})
    assert r.status_code == 500 and "PowerPoint" in r.get_json()["error"]


# ---------------- 版式切换 ----------------

def test_layout_updates_slide(vis_env, client):
    r = client.post("/api/slide/0/layout", json={"layout": "cards"})
    assert r.status_code == 200 and r.get_json()["layout"] == "cards"
    with app_mod.lock:
        assert app_mod.state["slides"][0]["layout"] == "cards"


def test_layout_reset_to_auto(vis_env, client):
    with app_mod.lock:
        app_mod.state["slides"][0]["layout"] = "cards"
    r = client.post("/api/slide/0/layout", json={"layout": ""})
    assert r.status_code == 200 and r.get_json()["layout"] is None
    with app_mod.lock:
        assert app_mod.state["slides"][0]["layout"] is None


def test_layout_invalid_value(vis_env, client):
    r = client.post("/api/slide/0/layout", json={"layout": "bogus"})
    assert r.status_code == 400


def test_layout_wrong_phase(vis_env, client):
    with app_mod.lock:
        app_mod.state["phase"] = "designing"
    r = client.post("/api/slide/0/layout", json={"layout": "cards"})
    assert r.status_code == 409


def test_layout_missing_page(vis_env, client):
    r = client.post("/api/slide/999/layout", json={"layout": "cards"})
    assert r.status_code == 404


# ---------------- 逐页缩略图 ----------------

def test_thumbnails_requires_deck(vis_env, client):
    r = client.post("/api/thumbnails", json={})
    assert r.status_code == 409


def test_thumbnails_queues_worker(vis_env, client, monkeypatch):
    deck = vis_env / "decks" / "deck_x_20260913_000000.html"
    deck.write_text("<html></html>", encoding="utf-8")
    with app_mod.lock:
        app_mod.state["html_path"] = "/decks/deck_x_20260913_000000.html"

    calls = []
    monkeypatch.setattr(app_mod.shot_mod, "shot_deck",
                        lambda html, out: calls.append((html, out)) or [out + "/slide_1.png"])
    r = client.post("/api/thumbnails", json={})
    assert r.status_code == 200 and r.get_json()["queued"] is True
    for _ in range(100):   # 等 worker 线程收尾
        if not app_mod._thumbnail_jobs:
            break
        time.sleep(0.05)
    assert calls and calls[0][0].endswith("deck_x_20260913_000000.html")
    assert "deck_x_20260913_000000" in calls[0][1]


def test_thumbnails_busy_guard(vis_env, client, monkeypatch):
    deck = vis_env / "decks" / "deck_y_20260913_000000.html"
    deck.write_text("<html></html>", encoding="utf-8")
    with app_mod.lock:
        app_mod.state["html_path"] = "/decks/deck_y_20260913_000000.html"
        app_mod._thumbnail_jobs.add("deck_y_20260913_000000")
    try:
        r = client.post("/api/thumbnails", json={})
        assert r.status_code == 409
    finally:
        with app_mod.lock:
            app_mod._thumbnail_jobs.discard("deck_y_20260913_000000")
