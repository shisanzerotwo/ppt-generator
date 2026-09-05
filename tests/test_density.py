"""内容密度档位：prompt 注入 / 参数校验 / 缓存键区分 / 透传 outline。"""

import pytest

import app as app_mod
import outline


@pytest.fixture
def client():
    app_mod.app.config["TESTING"] = True
    return app_mod.app.test_client()


@pytest.fixture(autouse=True)
def isolate(tmp_path, monkeypatch):
    before = {k: (list(v) if isinstance(v, list) else v)
              for k, v in app_mod.state.items()}
    monkeypatch.setattr(app_mod, "IMAGES_DIR", str(tmp_path / "images"))
    monkeypatch.setattr(app_mod, "DECKS_DIR", str(tmp_path / "decks"))
    monkeypatch.setattr(app_mod, "OUTPUT_DIR", str(tmp_path))

    def fake_gen(prompt, path):
        import os
        os.makedirs(os.path.dirname(path), exist_ok=True)
        open(path, "wb").write(b"\x89PNG\r\n\x1a\n")
        return path

    monkeypatch.setattr(app_mod.image_gen, "generate_image", fake_gen)
    monkeypatch.setattr(app_mod.critic, "review_image",
                        lambda t, p, path: {"ok": True, "reason": "", "advice": ""})
    monkeypatch.setattr(app_mod.html_gen, "generate_html_deck",
                        lambda topic, slides, image_map, style=None:
                        "<!DOCTYPE html><html><body>x</body></html>")
    monkeypatch.setattr(app_mod.style_mod, "decide_style",
                        lambda topic, slides: dict(app_mod.style_mod.STYLE_LIBRARY["tech-deep"],
                                                   key="tech-deep", reason="测试"))
    yield
    app_mod.state.clear()
    app_mod.state.update(before)


def test_prompt_carries_density_hint():
    dense = outline.PROMPT_TEMPLATE.format(topic="t", density_hint=outline.DENSITY_HINTS["dense"])
    assert "3~4 条" in dense and "30~50 字" in dense
    sparse = outline.PROMPT_TEMPLATE.format(topic="t", density_hint=outline.DENSITY_HINTS["sparse"])
    assert "2 条" in sparse
    ft = outline.FROM_TEXT_PROMPT.format(text="x", density_hint=outline.DENSITY_HINTS["balanced"])
    assert "内容密度：标准" in ft


def test_generate_passes_density_to_outline(client, monkeypatch):
    seen = {}

    def fake_outline(topic, density="balanced"):
        seen["density"] = density
        return [{"type": "cover", "title": "封面", "points": [], "image_prompt": "",
                 "chart": None, "layout": None}]

    monkeypatch.setattr(outline, "generate_outline", fake_outline)
    client.post("/api/stepwise", json={"enabled": False})
    client.post("/api/generate", json={"topic": "密度透传", "density": "dense"})
    # 等 worker 跑完（ready），断言 outline 收到了 dense
    import time
    deadline = time.time() + 8
    while time.time() < deadline and app_mod.state["phase"] != "ready":
        time.sleep(0.1)
    assert seen["density"] == "dense"


def test_invalid_density_rejected(client):
    resp = client.post("/api/generate", json={"topic": "t", "density": "超厚"})
    assert resp.status_code == 400


def test_cache_key_differs_by_density(tmp_path, monkeypatch):
    monkeypatch.setattr(app_mod, "OUTPUT_DIR", str(tmp_path))
    k1 = app_mod._outline_cache_path("主题", False, "sparse")
    k2 = app_mod._outline_cache_path("主题", False, "dense")
    assert k1 != k2
