"""阶段三新功能单测：页面排序/增删、品牌模板、历史项目库（路线图 #6/#7/#8）。"""

import json
import os
import time

import pytest

import app as app_mod


@pytest.fixture
def client(tmp_path, monkeypatch):
    """隔离 DECKS_DIR/PROJECTS_DIR，state 置为 ready 且带 3 页样例。"""
    monkeypatch.setattr(app_mod, "DECKS_DIR", str(tmp_path / "decks"))
    monkeypatch.setattr(app_mod, "PROJECTS_DIR", str(tmp_path / "projects"))
    os.makedirs(str(tmp_path / "decks"))
    app_mod.app.config["TESTING"] = True
    _seed_ready_state()
    yield app_mod.app.test_client()
    _reset_state()


def _seed_ready_state():
    slides = [
        {"type": "cover", "title": f"页{i}", "points": [], "image_prompt": "",
         "chart": None, "layout": None, "image": None, "imageStatus": "skipped",
         "review": {"ok": True, "reason": "", "tries": 0}}
        for i in range(3)
    ]
    with app_mod.lock:
        app_mod.state.update({"topic": "测试主题", "phase": "ready", "slides": slides,
                              "style": "tech-deep", "style_name": "深空科技",
                              "brand": None, "html_path": None, "log": []})


def _reset_state():
    with app_mod.lock:
        app_mod.state.update({"topic": "", "phase": "idle", "slides": [],
                              "style": None, "style_name": "", "brand": None,
                              "html_path": None, "log": []})


# ---------- #7 排序 / 加页 / 删页 ----------

def test_reorder_moves_last_to_first(client):
    r = client.post("/api/slide/reorder", json={"order": [2, 0, 1]})
    assert r.status_code == 200
    assert [s["title"] for s in app_mod.state["slides"]] == ["页2", "页0", "页1"]


def test_reorder_rejects_invalid_order(client):
    r = client.post("/api/slide/reorder", json={"order": [0, 0, 1]})
    assert r.status_code == 400


def test_add_inserts_after_index(client):
    r = client.post("/api/slide/add", json={"after": 0})
    assert r.status_code == 200
    slides = app_mod.state["slides"]
    assert len(slides) == 4 and slides[1]["title"] == "新页面"
    assert slides[1]["type"] == "content"


def test_delete_removes_slide(client):
    r = client.delete("/api/slide/1/delete")
    assert r.status_code == 200
    assert [s["title"] for s in app_mod.state["slides"]] == ["页0", "页2"]


def test_delete_last_slide_rejected(client):
    with app_mod.lock:
        app_mod.state["slides"] = app_mod.state["slides"][:1]
    r = client.delete("/api/slide/0/delete")
    assert r.status_code == 400


def test_slide_ops_rejected_when_not_ready(client):
    with app_mod.lock:
        app_mod.state["phase"] = "images"
    assert client.post("/api/slide/reorder", json={"order": [1, 0, 2]}).status_code == 409
    assert client.post("/api/slide/add", json={}).status_code == 409
    assert client.delete("/api/slide/0/delete").status_code == 409


# ---------- #6 品牌模板 ----------

def test_brand_set_and_applied(client):
    r = client.post("/api/brand", json={"name": "星尘科技", "color": "#FF8800"})
    assert r.status_code == 200
    assert app_mod.state["brand"] == {"name": "星尘科技", "color": "#FF8800"}
    applied = app_mod._apply_brand({"name": "深空科技", "guidance": "深色底"},
                                   app_mod.state["brand"])
    assert applied["accent"] == "#FF8800"
    assert "星尘科技" in applied["name"] and "星尘科技" in applied["guidance"]


def test_brand_invalid_color_rejected(client):
    assert client.post("/api/brand", json={"name": "x", "color": "orange"}).status_code == 400
    assert app_mod.state["brand"] is None


def test_brand_clear(client):
    client.post("/api/brand", json={"name": "x", "color": "#123456"})
    r = client.post("/api/brand", json={})
    assert r.status_code == 200 and app_mod.state["brand"] is None


# ---------- #8 历史项目库 ----------

def test_snapshot_saved_and_listed(client, tmp_path):
    app_mod._save_project_snapshot()
    files = os.listdir(str(tmp_path / "projects"))
    assert len(files) == 1 and files[0].endswith(".json")
    r = client.get("/api/projects")
    decks = r.get_json()["projects"]
    assert len(decks) == 1 and decks[0]["title"] == "测试主题"


def test_load_restores_state(client, tmp_path):
    app_mod._save_project_snapshot()
    name = os.listdir(str(tmp_path / "projects"))[0]
    # 模拟切换走：清空当前会话
    _reset_state()
    r = client.post("/api/projects/load", json={"name": name})
    assert r.status_code == 200
    assert app_mod.state["phase"] == "ready"
    assert app_mod.state["topic"] == "测试主题"
    assert len(app_mod.state["slides"]) == 3
    assert app_mod.state["style"] == "tech-deep"


def test_load_missing_design_html_set_to_none(client, tmp_path):
    app_mod._save_project_snapshot()
    name = os.listdir(str(tmp_path / "projects"))[0]
    # 快照里 html_path 指向不存在的文件
    p = os.path.join(str(tmp_path / "projects"), name)
    payload = json.load(open(p, encoding="utf-8"))
    payload["html_path"] = "/decks/不存在的文件.html"
    json.dump(payload, open(p, "w", encoding="utf-8"), ensure_ascii=False)
    _reset_state()
    r = client.post("/api/projects/load", json={"name": name})
    assert r.status_code == 200
    assert app_mod.state["html_path"] is None


def test_load_rejects_path_traversal(client):
    assert client.post("/api/projects/load",
                       json={"name": "../secrets.json"}).status_code == 400
    assert client.post("/api/projects/load",
                       json={"name": "a/b.json"}).status_code == 400


def test_load_missing_project_404(client):
    assert client.post("/api/projects/load",
                       json={"name": "不存在_20260101_000000.json"}).status_code == 404
