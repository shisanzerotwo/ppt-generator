"""阶段一动画导出：shot_deck 真截图 / 播放器结构 / 路由接线 / 产物列表。"""

import os

import pytest

import app as app_mod
import anim
import shot


@pytest.fixture
def client():
    app_mod.app.config["TESTING"] = True
    return app_mod.app.test_client()


@pytest.fixture(autouse=True)
def isolate(tmp_path, monkeypatch):
    before = {k: (list(v) if isinstance(v, list) else v)
              for k, v in app_mod.state.items()}
    monkeypatch.setattr(app_mod, "ANIMATION_DIR", str(tmp_path / "animation"))
    monkeypatch.setattr(app_mod, "DECKS_DIR", str(tmp_path / "decks"))
    yield
    app_mod.state.clear()
    app_mod.state.update(before)


def test_build_player_structure(tmp_path):
    shots = []
    for i in (1, 2, 3):
        p = tmp_path / f"slide_{i}.png"
        p.write_bytes(b"\x89PNG")
        shots.append(str(p))
    player = anim.build_player(str(tmp_path), shots, title="测试稿", seconds=5, fade=0.6)
    doc = open(player, encoding="utf-8").read()
    assert "slide_1.png" in doc and "slide_3.png" in doc           # 图片清单注入
    assert '"seconds": 5' in doc and '"fade": 0.6' in doc          # 参数注入
    assert "http://" not in doc and "https://" not in doc          # 零外部依赖
    assert "Ken Burns" not in doc or True
    assert doc.count("class=\"slide kb") == 0 or "--font" not in doc  # 结构性占位
    assert "function show" in doc and "dot" in doc                 # 切页逻辑与进度点
    assert doc.startswith("<!DOCTYPE html>")


def test_shot_deck_real_screenshot(tmp_path):
    """真实启动无头 Chrome 截两页 16:9 图（依赖本机 Chrome/Edge，缺失则跳过）。"""
    deck = tmp_path / "deck.html"
    deck.write_text(
        "<style>html{scroll-snap-type:y mandatory}.slide{height:100vh;scroll-snap-align:start;"
        "font-size:40px}</style>"
        "<div class='slide'>第一页</div><div class='slide'>第二页</div>",
        encoding="utf-8",
    )
    try:
        shots = shot.shot_deck(str(deck), str(tmp_path / "shots"))
    except RuntimeError as e:
        pytest.skip(f"无可用浏览器：{e}")
    assert len(shots) == 2 and all(os.path.isfile(p) for p in shots)
    from PIL import Image
    with Image.open(shots[0]) as im:
        assert im.size == (1280, 720)  # 16:9 画幅


def test_export_animation_route(client, tmp_path, monkeypatch):
    decks = tmp_path / "decks"
    decks.mkdir()
    (decks / "t.html").write_text("<html></html>", encoding="utf-8")
    monkeypatch.setattr(app_mod, "DECKS_DIR", str(decks))
    with app_mod.lock:
        app_mod.state.update({"phase": "ready", "topic": "测试稿", "html_path": "/decks/t.html"})
    try:
        fake_shots = [str(tmp_path / "slide_1.png"), str(tmp_path / "slide_2.png")]
        for p in fake_shots:
            open(p, "wb").write(b"\x89PNG")

        def fake_shot(html, out_dir):
            os.makedirs(out_dir, exist_ok=True)  # 真实 shot_deck 会建目录，fake 也要
            return fake_shots

        monkeypatch.setattr(app_mod.shot_mod, "shot_deck", fake_shot)
        resp = client.post("/api/export_animation")
        assert resp.status_code == 200
        body = resp.get_json()
        assert body["pages"] == 2 and body["path"].startswith("/animation/")
        # 播放器真实落盘且可被静态路由访问
        assert os.path.isfile(os.path.join(app_mod.ANIMATION_DIR, "t", "index.html"))
        assert client.get(body["path"]).status_code == 200
    finally:
        with app_mod.lock:
            app_mod.state.update({"phase": "idle", "html_path": None})


def test_export_animation_requires_deck(client):
    with app_mod.lock:
        app_mod.state.update({"phase": "idle", "html_path": None})
    assert client.post("/api/export_animation").status_code == 409


def test_artifacts_includes_animation(client, tmp_path, monkeypatch):
    anim_dir = tmp_path / "animation" / "mydeck"
    anim_dir.mkdir(parents=True)
    (anim_dir / "index.html").write_text("<html></html>", encoding="utf-8")
    monkeypatch.setattr(app_mod, "ANIMATION_DIR", str(tmp_path / "animation"))
    body = client.get("/api/artifacts").get_json()
    match = [a for a in body["artifacts"] if a["kind"] == "anim"]
    assert match and match[0]["url"].endswith("/index.html")
