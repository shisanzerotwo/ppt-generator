"""/api/refine 定点修改的路由层测试（mock LLM 与设计/生图，不发真实请求）。"""

import time

import pytest

import app as app_mod
import critic


def _slides():
    return [
        {"type": "cover", "title": "封面", "points": [], "image_prompt": "封面画面",
         "chart": None, "layout": None, "image": "/images/slide_0.png",
         "imageStatus": "done", "review": {"ok": True, "reason": "", "tries": 0}},
        {"type": "content", "title": "第二页", "points": ["要点甲：原始说明", "要点乙：另一条"],
         "image_prompt": "原画面", "chart": None, "layout": "image-right",
         "image": "/images/slide_1.png", "imageStatus": "done",
         "review": {"ok": True, "reason": "", "tries": 0}},
    ]


@pytest.fixture
def client(monkeypatch):
    # 屏蔽真实网络：定点改写只桩 _ask，设计与生图整段跳过
    monkeypatch.setattr(critic, "_ask", lambda p: "要点甲：改短后的说明")
    monkeypatch.setattr(app_mod, "_design_and_save", lambda: True)
    monkeypatch.setattr(app_mod, "_gen_images", lambda indices=None: None)
    app_mod.app.config["TESTING"] = True
    return app_mod.app.test_client()


@pytest.fixture(autouse=True)
def restore_state():
    before = dict(app_mod.state)
    before["slides"] = [dict(s) for s in before["slides"]]
    yield
    app_mod.state.clear()
    app_mod.state.update(before)


def _ready():
    app_mod.state.update({"topic": "测试", "slides": _slides(), "phase": "ready",
                          "html_path": None, "log": [], "style": None,
                          "style_name": "", "brand": None})


# ===== 定点：只改圈定处 =====

def test_point_level_edits_only_matched_quote(client):
    _ready()
    r = client.post("/api/refine", json={"instruction": "缩到12字",
                                        "target": {"slide": 1, "quote": "要点甲"}})
    assert r.status_code == 200
    time.sleep(0.5)
    pts = app_mod.state["slides"][1]["points"]
    assert pts == ["要点甲：改短后的说明", "要点乙：另一条"]  # 另一条不动


def test_point_level_leaves_other_slides_and_images_intact(client):
    _ready()
    client.post("/api/refine", json={"instruction": "改", "target": {"slide": 1, "quote": "要点甲"}})
    time.sleep(0.5)
    s = app_mod.state["slides"]
    assert s[0]["title"] == "封面" and s[0]["image"] == "/images/slide_0.png"
    # 定点修改不得把该页配图打回重做
    assert s[1]["image"] == "/images/slide_1.png"
    assert s[1]["imageStatus"] == "done"
    assert app_mod.state["phase"] == "ready"


# ===== 输入校验（回归：曾返回 200 让前端误报成功） =====

@pytest.mark.parametrize("bad", [{"slide": 99}, {"slide": -1}, {"slide": "1"},
                                {"slide": True}, {"slide": None}, {}])
def test_invalid_slide_rejected_at_http_layer(client, bad):
    _ready()
    r = client.post("/api/refine", json={"instruction": "改", "target": bad})
    assert r.status_code == 400
    assert app_mod.state["phase"] == "ready"       # 不被推到 refining 卡住


def test_quote_not_found_rejected(client):
    _ready()
    r = client.post("/api/refine", json={"instruction": "改",
                                         "target": {"slide": 1, "quote": "不存在的选择内容"}})
    assert r.status_code == 400
    assert "未找到" in r.get_json()["error"]
    assert app_mod.state["phase"] == "ready"


def test_target_wrong_type_rejected(client):
    _ready()
    assert client.post("/api/refine", json={"instruction": "改", "target": ["a"]}).status_code == 400


# ===== 整篇路径行为不变 =====

def test_no_target_still_rebuilds_everything(client, monkeypatch):
    _ready()
    monkeypatch.setattr(critic, "refine_outline",
                        lambda sl, ins: [dict(sl[0]), {**sl[1], "title": "整篇改过"}])
    r = client.post("/api/refine", json={"instruction": "整篇换商务风"})
    assert r.status_code == 200
    time.sleep(0.6)
    s = app_mod.state["slides"]
    assert s[1]["title"] == "整篇改过"
    assert s[1]["image"] is None and s[1]["imageStatus"] == "pending"  # 整篇需打回重生图


def test_empty_instruction_rejected(client):
    _ready()
    assert client.post("/api/refine", json={"instruction": "  "}).status_code == 400


@pytest.mark.parametrize("phase", ["images", "outline", "designing", "idle"])
def test_refine_only_allowed_when_ready(client, phase):
    _ready()
    app_mod.state["phase"] = phase
    assert client.post("/api/refine", json={"instruction": "改"}).status_code == 409
