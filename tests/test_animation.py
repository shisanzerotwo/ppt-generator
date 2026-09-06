"""教学动画（元素级动效）：播放器结构 / 路由接线 / 产物列表。"""

import os

import pytest

import app as app_mod
import anim


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
    monkeypatch.setattr(app_mod, "VIDEOS_DIR", str(tmp_path / "videos"))
    yield
    app_mod.state.clear()
    app_mod.state.update(before)


def test_build_player_element_engine(tmp_path):
    player = anim.build_player(str(tmp_path), "/decks/t.html", title="测试稿",
                               auto_step_ms=1500)
    doc = open(player, encoding="utf-8").read()
    assert 'src="/decks/t.html"' in doc                       # iframe 同源加载设计稿
    assert ".ae{opacity:0" in doc and ".ae.on{opacity:1" in doc  # 元素级入场 CSS
    assert "const SEL = 'h1,h2,h3,h4,p,li,img,svg,video,table'" in doc  # 元素收集选择器
    assert "function stepFwd" in doc and "function stepBack" in doc  # 步进/回退
    assert "revealAll" in doc and "toggleAuto" in doc          # 全显 / 自动讲解
    assert '"autoStepMs": 1500' in doc                         # 参数注入
    assert "aspect-ratio:16/9" in doc                          # 16:9 舞台
    assert "https://" not in doc                               # 零外部依赖
    assert doc.startswith("<!DOCTYPE html>")


def test_build_player_escapes_malicious_title(tmp_path):
    """审计修复回归：主题含 </title><script> 不得在可执行上下文注入。"""
    player = anim.build_player(str(tmp_path), "/decks/t.html",
                               title='x</title><script>alert(1)</script>')
    doc = open(player, encoding="utf-8").read()
    # title/h1 上下文：整体转义
    assert "&lt;/title&gt;&lt;script&gt;" in doc
    # script 字符串上下文：危险的是闭合序列 </script>，开标签惰性无害
    assert "</script>alert" not in doc
    assert "<\\/script>" in doc


def test_export_animation_route(client, tmp_path, monkeypatch):
    decks = tmp_path / "decks"
    decks.mkdir()
    (decks / "t.html").write_text("<html></html>", encoding="utf-8")
    monkeypatch.setattr(app_mod, "DECKS_DIR", str(decks))
    with app_mod.lock:
        app_mod.state.update({"phase": "ready", "topic": "测试稿", "html_path": "/decks/t.html"})
    try:
        resp = client.post("/api/export_animation")
        assert resp.status_code == 200
        body = resp.get_json()
        assert body["deck"] == "/decks/t.html"
        player = os.path.join(app_mod.ANIMATION_DIR, "t", "index.html")
        assert os.path.isfile(player)
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


def test_player_placeholder_collision_escaped(tmp_path):
    """审计 M2：title 含 __CONFIG__ 时不得经二次替换绕过转义。"""
    title = '<img src=x onerror=alert(1)>__CONFIG__'
    player = anim.build_player(str(tmp_path), "/decks/t.html", title=title)
    doc = open(player, encoding="utf-8").read()
    # HTML 上下文必须转义（活标签不可出现）
    assert "<title>&lt;img" in doc and "<h1>&lt;img" in doc
    h1 = doc.split("<h1>")[1].split("</h1>")[0]
    assert '"deck"' not in h1  # 单遍替换生效：title 里的 __CONFIG__ 不得被展开成 CFG JSON
    # 原始标签文本只允许存在于 CFG 的 JS 数据行（textContent 消费，不进 innerHTML）
    raw = [l for l in doc.splitlines() if "<img src=x" in l]
    assert len(raw) == 1 and "const CFG" in raw[0]
