"""速度优化：生图与设计并行（非分步模式）+ 大纲指纹缓存。

并行断言方式：fake 生图阻塞等一个事件，fake 设计负责 set 事件——
并行实现下"设计开始"必然早于"图片完成"；若退回串行实现本用例超时失败。
"""

import json
import os
import threading
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
    before = {k: (list(v) if isinstance(v, list) else v)
              for k, v in app_mod.state.items()}
    monkeypatch.setattr(app_mod, "IMAGES_DIR", str(tmp_path / "images"))
    monkeypatch.setattr(app_mod, "DECKS_DIR", str(tmp_path / "decks"))
    monkeypatch.setattr(app_mod, "PROJECTS_DIR", str(tmp_path / "projects"))
    monkeypatch.setattr(app_mod, "OUTPUT_DIR", str(tmp_path))

    monkeypatch.setattr(style_mod, "decide_style",
                        lambda topic, slides: dict(style_mod.STYLE_LIBRARY["tech-deep"],
                                                   key="tech-deep", reason="测试"))

    def fake_gen(prompt, path):
        os.makedirs(os.path.dirname(path), exist_ok=True)
        open(path, "wb").write(b"\x89PNG\r\n\x1a\n")
        return path

    monkeypatch.setattr(image_gen, "generate_image", fake_gen)
    monkeypatch.setattr(critic, "review_image",
                        lambda title, points, path: {"ok": True, "reason": "", "advice": ""})
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


def _wait_for(client, predicate, timeout=10.0):
    deadline = time.time() + timeout
    while time.time() < deadline:
        s = client.get("/api/status").get_json()
        if predicate(s):
            return s
        time.sleep(0.05)
    raise AssertionError(f"等待超时，最后状态: {client.get('/api/status').get_json()}")


def test_design_runs_parallel_with_images(client, monkeypatch):
    ev = threading.Event()
    marks = {"design_start": None, "img_done": []}

    def slow_gen(prompt, path):
        ev.wait(timeout=8)          # 等设计开始才"完成"——并行时设计会先到
        marks["img_done"].append(time.time())
        os.makedirs(os.path.dirname(path), exist_ok=True)
        open(path, "wb").write(b"\x89PNG\r\n\x1a\n")
        return path

    def fake_design(topic, slides, image_map, style=None):
        marks["design_start"] = time.time()
        ev.set()                     # 设计一开始就放行生图
        return "<!DOCTYPE html><html><body>x</body></html>"

    monkeypatch.setattr(image_gen, "generate_image", slow_gen)
    monkeypatch.setattr(html_gen, "generate_html_deck", fake_design)
    monkeypatch.setattr(outline, "generate_outline",
                        lambda topic: [{"type": "cover", "title": "封面", "points": [],
                                        "image_prompt": "p", "chart": None, "layout": None}])

    client.post("/api/stepwise", json={"enabled": False})
    client.post("/api/generate", json={"topic": "并行测试"})
    s = _wait_for(client, lambda x: x["phase"] == "ready")
    assert s["html_path"]
    assert marks["design_start"] is not None and marks["img_done"]
    assert marks["design_start"] < max(marks["img_done"])  # 设计开始早于图片完成 → 真并行
    assert os.path.isfile(os.path.join(app_mod.IMAGES_DIR, "slide_0.png"))


def test_outline_cache_second_run_skips_llm(client, monkeypatch):
    calls = []

    def counting_outline(topic):
        calls.append(topic)
        return [{"type": "cover", "title": "封面", "points": [], "image_prompt": "",
                 "chart": None, "layout": None}]

    monkeypatch.setattr(outline, "generate_outline", counting_outline)
    client.post("/api/stepwise", json={"enabled": False})
    client.post("/api/generate", json={"topic": "缓存测试主题"})
    _wait_for(client, lambda x: x["phase"] == "ready")
    assert len(calls) == 1

    # 同主题第二次：大纲直接命中缓存，不再调 LLM
    client.post("/api/generate", json={"topic": "缓存测试主题"})
    s = _wait_for(client, lambda x: x["phase"] == "ready")
    assert len(calls) == 1
    assert any("命中大纲缓存" in line["msg"] for line in s["log"])
    # 缓存文件落在（被隔离的）OUTPUT_DIR/outline_cache
    cache_dir = os.path.join(app_mod.OUTPUT_DIR, "outline_cache")
    assert os.path.isdir(cache_dir) and os.listdir(cache_dir)


def test_cache_key_differs_by_mode_and_content():
    p1 = app_mod._outline_cache_path("主题A", False)
    p2 = app_mod._outline_cache_path("主题B", False)
    p3 = app_mod._outline_cache_path("主题A", True)
    assert p1 != p2 and p1 != p3


def test_cache_roundtrip(tmp_path, monkeypatch):
    monkeypatch.setattr(app_mod, "OUTPUT_DIR", str(tmp_path))
    slides = [{"type": "cover", "title": "t", "points": [], "image_prompt": "p",
               "chart": None, "layout": None}]
    app_mod._save_outline_cache("内容X", True, slides)
    assert app_mod._load_outline_cache("内容X", True) == slides
    assert app_mod._load_outline_cache("内容X", False) is None      # 模式不同不串
    assert app_mod._load_outline_cache("内容Y", True) is None       # 内容不同不串
