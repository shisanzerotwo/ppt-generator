"""分步确认（human-in-the-loop）全链路测试：暂停 → 放行 → 完成。

mock 掉大纲/生图/视觉校验/HTML 设计四处 LLM 调用，真实跑 Flask 路由与 worker 线程。
"""

import time

import pytest

import app as app_mod
import critic
import html_gen
import image_gen
import outline
import style as style_mod


@pytest.fixture(autouse=True)
def isolate(tmp_path, monkeypatch):
    """隔离输出目录 + 桩掉四处外部调用，并在用例后还原全局 state。"""
    before = {k: (list(v) if isinstance(v, list) else v)
              for k, v in app_mod.state.items()}

    monkeypatch.setattr(app_mod, "IMAGES_DIR", str(tmp_path / "images"))
    monkeypatch.setattr(app_mod, "DECKS_DIR", str(tmp_path / "decks"))
    monkeypatch.setattr(app_mod, "PROJECTS_DIR", str(tmp_path / "projects"))
    monkeypatch.setattr(app_mod, "OUTPUT_DIR", str(tmp_path))

    def fake_outline(topic, density="balanced"):
        return [{"type": "cover", "title": "封面", "points": [], "image_prompt": "p",
                 "chart": None, "layout": None},
                {"type": "content", "title": "内容页", "points": ["要点：说明"],
                 "image_prompt": "p2", "chart": None, "layout": "image-right"}]

    monkeypatch.setattr(outline, "generate_outline", fake_outline)
    monkeypatch.setattr(style_mod, "decide_style",
                        lambda topic, slides: dict(style_mod.STYLE_LIBRARY["tech-deep"],
                                                   key="tech-deep", reason="测试"))
    # 生图：写一个占位文件，避免真实网络与图片依赖
    def fake_gen(prompt, path):
        import os
        os.makedirs(os.path.dirname(path), exist_ok=True)
        open(path, "wb").write(b"\x89PNG\r\n\x1a\n")
        return path
    monkeypatch.setattr(image_gen, "generate_image", fake_gen)
    monkeypatch.setattr(critic, "review_image",
                        lambda title, points, path: {"ok": True, "reason": "桩", "advice": ""})
    monkeypatch.setattr(html_gen, "generate_html_deck",
                        lambda topic, slides, image_map, style=None:
                        "<!DOCTYPE html><html><head></head><body>x</body></html>")
    yield
    app_mod.state.clear()
    app_mod.state.update(before)


@pytest.fixture
def client():
    app_mod.app.config["TESTING"] = True
    return app_mod.app.test_client()


def _wait_for(client, predicate, timeout=8.0):
    deadline = time.time() + timeout
    while time.time() < deadline:
        s = client.get("/api/status").get_json()
        if predicate(s):
            return s
        time.sleep(0.1)
    raise AssertionError(f"等待超时，最后状态: {client.get('/api/status').get_json()['phase']}")


def test_stepwise_pauses_after_outline(client):
    client.post("/api/stepwise", json={"enabled": True})
    assert client.post("/api/generate", json={"topic": "测试主题"}).status_code == 200
    s = _wait_for(client, lambda x: x["phase"] == "review")
    assert s["await_step"] == "outline"
    # 暂停时大纲已就绪，可审阅
    assert len(s["slides"]) == 2


def test_continue_finishes_pipeline(client):
    client.post("/api/stepwise", json={"enabled": True})
    client.post("/api/generate", json={"topic": "测试主题"})
    _wait_for(client, lambda x: x["phase"] == "review" and x["await_step"] == "outline")
    r = client.post("/api/continue")
    assert r.status_code == 200 and r.get_json()["resumed"] == "outline"
    # 第二次暂停在配图后、设计前
    s = _wait_for(client, lambda x: x["phase"] == "review" and x["await_step"] == "design")
    assert s["await_step"] == "design"
    client.post("/api/continue")
    s = _wait_for(client, lambda x: x["phase"] == "ready")
    assert s["await_step"] is None
    assert s["html_path"] and s["html_path"].startswith("/decks/")


def test_stepwise_off_runs_through(client):
    client.post("/api/stepwise", json={"enabled": False})
    client.post("/api/generate", json={"topic": "测试主题"})
    s = _wait_for(client, lambda x: x["phase"] in ("ready", "idle"))
    assert s["phase"] == "ready"
    assert s["await_step"] is None
    assert s["html_path"]


def test_continue_when_not_paused_is_409(client):
    client.post("/api/stepwise", json={"enabled": False})
    assert client.post("/api/continue").status_code == 409


def test_await_step_reset_between_runs(client):
    """连续两次生成不得被上一轮残留的放行信号瞬间跳过。"""
    client.post("/api/stepwise", json={"enabled": True})
    client.post("/api/generate", json={"topic": "第一次"})
    _wait_for(client, lambda x: x["phase"] == "review")
    client.post("/api/continue")
    _wait_for(client, lambda x: x["phase"] == "review" and x["await_step"] == "design")
    client.post("/api/continue")
    _wait_for(client, lambda x: x["phase"] == "ready")
    # 第二轮：应再次停在 outline，而不是带着 set 状态的 event 直接冲过
    client.post("/api/generate", json={"topic": "第二次"})
    s = _wait_for(client, lambda x: x["phase"] == "review")
    assert s["await_step"] == "outline"


def test_refine_rejected_during_review(client):
    """暂停期不允许 refine（会另起 worker 与原闸门竞争），改文字请用单页保存。"""
    client.post("/api/stepwise", json={"enabled": True})
    client.post("/api/generate", json={"topic": "测试"})
    _wait_for(client, lambda x: x["phase"] == "review")
    assert client.post("/api/refine", json={"instruction": "改"}).status_code == 409
    client.post("/api/continue")


def test_edit_text_works_during_review(client):
    """暂停期直接改卡片文字应可用，且放行后按新内容继续。"""
    client.post("/api/stepwise", json={"enabled": True})
    client.post("/api/generate", json={"topic": "测试"})
    _wait_for(client, lambda x: x["phase"] == "review")
    r = client.post("/api/slide/1/text", json={"title": "我改过的标题", "points": ["新要点"]})
    assert r.status_code == 200
    # 流程有 outline / design 两个暂停点，需各放行一次
    client.post("/api/continue")
    _wait_for(client, lambda x: x["phase"] == "review" and x["await_step"] == "design")
    client.post("/api/continue")
    s = _wait_for(client, lambda x: x["phase"] == "ready")
    assert s["slides"][1]["title"] == "我改过的标题"
