"""删除功能：历史设计稿（联动产物）/ 历史项目 / 导出产物 / pptx 导入记录 / 回收站。"""
import os
import pathlib
import time

import pytest

import app as app_mod


@pytest.fixture
def client():
    app_mod.app.config["TESTING"] = True
    return app_mod.app.test_client()


@pytest.fixture
def del_env(tmp_path, monkeypatch):
    """tmp 隔离全部产物目录 + ready 态。目录名与测试用到的字面名严格一致。"""
    dirs = {"OUTPUT_DIR": "output_dir", "DECKS_DIR": "decks", "PROJECTS_DIR": "projects",
            "THUMBNAILS_DIR": "thumbnails", "ANIMATION_DIR": "animation",
            "VIDEOS_DIR": "videos", "PPTX_UPLOAD_DIR": "pptx_uploads", "PPTX_SRC_DIR": "pptx_src"}
    for attr, sub in dirs.items():
        monkeypatch.setattr(app_mod, attr, str(tmp_path / sub))
        os.makedirs(tmp_path / sub, exist_ok=True)
    monkeypatch.setattr(app_mod, "TRASH_DIR", str(tmp_path / "trash"))
    with app_mod.lock:
        app_mod.state["phase"] = "ready"
        app_mod.state["html_path"] = "/decks/keep.html"
    yield tmp_path
    with app_mod.lock:
        app_mod.state["phase"] = "idle"
        app_mod.state["html_path"] = None


def _mk_deck(env, stem="deck_a", with_side=True):
    (env / "decks" / f"{stem}.html").write_text("<html></html>", encoding="utf-8")
    if with_side:
        t = env / "thumbnails" / stem
        t.mkdir(parents=True, exist_ok=True)
        (t / "slide_1.png").write_bytes(b"png")
        a = env / "animation" / stem
        a.mkdir(parents=True, exist_ok=True)
        (a / "index.html").write_text("<html></html>", encoding="utf-8")
        (env / "videos" / f"{stem}_20260913_000000_000.mp4").write_bytes(b"mp4")
        out = env / "output_dir"          # 与 OUTPUT_DIR 的映射一致
        out.mkdir(exist_ok=True)
        (out / f"{stem}.pptx").write_bytes(b"PK")
        (out / f"{stem}.pdf").write_bytes(b"%PDF")
    return env / "decks" / f"{stem}.html"


# ---------------- 历史设计稿（联动删除）----------------

def test_deck_delete_moves_bundle_to_trash(del_env, client):
    _mk_deck(del_env)
    r = client.post("/api/decks/delete", json={"name": "deck_a.html"})
    assert r.status_code == 200 and r.get_json()["ok"] is True
    assert len(r.get_json()["removed"]) >= 5   # html+缩略图+动画+视频+pptx+pdf
    assert not (del_env / "decks" / "deck_a.html").exists()
    trash = del_env / "trash"
    assert list(trash.rglob("*deck_a.html"))
    assert list(trash.rglob("*slide_1.png"))
    assert list(trash.rglob("*deck_a_20260913_000000_000.mp4"))
    assert list(trash.rglob("*deck_a.pptx"))
    # 原位全部清空
    assert not list((env for env in [del_env / "thumbnails" / "deck_a"] if env.exists()))


def test_deck_delete_blocked_when_loaded_in_workarea(del_env, client):
    _mk_deck(del_env, "keep")
    with app_mod.lock:
        app_mod.state["html_path"] = "/decks/keep.html"
    r = client.post("/api/decks/delete", json={"name": "keep.html"})
    assert r.status_code == 409
    assert (del_env / "decks" / "keep.html").exists()


def test_deck_delete_rejects_non_html_and_missing(del_env, client):
    assert client.post("/api/decks/delete", json={"name": "x.pptx"}).status_code == 400
    assert client.post("/api/decks/delete", json={"name": "nope.html"}).status_code == 404
    # 路径穿越：basename 化后不存在
    assert client.post("/api/decks/delete", json={"name": "../../x.html"}).status_code == 404


# ---------------- 历史项目 ----------------

def test_project_delete(del_env, client):
    pj = del_env / "projects" / "proj_20260913_000000.json"
    pj.write_text("{}", encoding="utf-8")
    r = client.post("/api/projects/delete", json={"name": "proj_20260913_000000.json"})
    assert r.status_code == 200
    assert not pj.exists()
    assert list((del_env / "trash").rglob("*proj_20260913_000000.json"))
    assert client.post("/api/projects/delete",
                       json={"name": "proj_20260913_000000.json"}).status_code == 404
    assert client.post("/api/projects/delete", json={"name": "x.txt"}).status_code == 400


# ---------------- 导出产物 ----------------

def test_artifact_delete_pptx_and_mp4(del_env, client):
    out = del_env / "output_dir"          # 与 OUTPUT_DIR 的映射一致
    out.mkdir(exist_ok=True)
    (out / "report_20260913_000000.pptx").write_bytes(b"PK")
    (del_env / "videos" / "v_20260913_000000_000.mp4").write_bytes(b"mp4")
    assert client.post("/api/artifacts/delete",
                       json={"name": "report_20260913_000000.pptx"}).status_code == 200
    assert client.post("/api/artifacts/delete",
                       json={"name": "v_20260913_000000_000.mp4"}).status_code == 200
    assert not (del_env / "report_20260913_000000.pptx").exists()
    assert list((del_env / "trash").rglob("*report_20260913_000000.pptx"))
    assert client.post("/api/artifacts/delete", json={"name": "x.exe"}).status_code == 400
    assert client.post("/api/artifacts/delete", json={"name": "nope.pptx"}).status_code == 404


# ---------------- pptx 导入记录 ----------------

def test_pptx_import_record_delete(del_env, client):
    up = del_env / "pptx_uploads" / "demo.src.pptx"
    up.write_bytes(b"PK")
    src = del_env / "pptx_src" / "demo"
    src.mkdir()
    (src / "deck.json").write_text("{}", encoding="utf-8")
    r = client.post("/api/pptx/delete", json={"name": "demo"})
    assert r.status_code == 200
    assert not up.exists() and not src.exists()
    assert list((del_env / "trash").rglob("*demo.src.pptx"))
    assert client.post("/api/pptx/delete", json={"name": "demo"}).status_code == 404
    assert client.post("/api/pptx/delete", json={"name": ""}).status_code == 400


# ---------------- 回收站过期清理 ----------------

def test_trash_expires_after_keep_days(del_env, client):
    old = del_env / "trash" / "decks"
    old.mkdir(parents=True, exist_ok=True)
    stale = old / "20260101_000000_old.html"
    stale.write_text("<html></html>", encoding="utf-8")
    ancient = time.time() - (app_mod.TRASH_KEEP_DAYS + 2) * 86400
    os.utime(stale, (ancient, ancient))
    _mk_deck(del_env, "fresh")
    client.post("/api/decks/delete", json={"name": "fresh.html"})
    assert not stale.exists()          # 过期项被惰性清理
    assert list((del_env / "trash").rglob("*fresh.html"))   # 新删项在
