"""阶段二视频配方层：命令构造纯函数 + 真 ffmpeg 合成冒烟（无 ffmpeg 则跳过）。"""

import os
import subprocess

import pytest

import video


def test_build_page_args_ken_burns_16_9():
    args = video.build_page_args("in.png", "out.mp4", seconds=4.0)
    joined = " ".join(args)
    assert "zoompan" in joined and "s=1280x720" in joined     # Ken Burns + 16:9 画幅
    assert "yuv420p" in joined and "libx264" in joined        # 兼容性编码
    assert "-t 4.000" in joined
    assert args[0] == "ffmpeg" and args[-1] == "out.mp4"


def test_build_final_args_xfade_offsets():
    clips = [f"p{i}.mp4" for i in range(3)]  # 3 片段、每页 4s、转场 0.8s
    args = video.build_final_args(clips, "out.mp4", seconds=4.0, fade=0.8)
    joined = " ".join(args)
    assert joined.count("-i") == 3
    # 转场起点：第 k 个 = k×(4−0.8) → 3.2 / 6.4
    assert "offset=3.200" in joined and "offset=6.400" in joined
    assert "-map [v2]" in joined
    assert args[-1] == "out.mp4"


def test_build_final_args_rejects_single_clip():
    with pytest.raises(ValueError):
        video.build_final_args(["only.mp4"], "out.mp4")


def _make_png(path):
    from PIL import Image
    Image.new("RGB", (1280, 720), (30, 60, 120)).save(path)


@pytest.mark.skipif(not video.ffmpeg_path(), reason="未安装 ffmpeg")
def test_synthesize_real_ffmpeg(tmp_path):
    imgs = []
    for i in (1, 2):
        p = tmp_path / f"slide_{i}.png"
        _make_png(p)
        imgs.append(str(p))
    out = str(tmp_path / "out.mp4")
    done = []
    video.synthesize(imgs, out, seconds=1.0, fade=0.4,
                     progress_cb=lambda d, t: done.append((d, t)))
    assert os.path.isfile(out) and os.path.getsize(out) > 0
    assert done == [(1, 2), (2, 2)]
    assert not os.path.exists(out + "_parts")  # 临时片段已清理

    probe = video.ffprobe_path()
    if probe:
        r = subprocess.run(
            [probe, "-v", "error", "-select_streams", "v:0",
             "-show_entries", "stream=width,height,codec_name:format=duration",
             "-of", "json", out],
            capture_output=True, text=True)
        info = __import__("json").loads(r.stdout)
        stream = info["streams"][0]
        assert (stream["width"], stream["height"]) == (1280, 720)
        assert stream["codec_name"] == "h264"
        # 2 页 × 1s − 1 个 0.4s 转场 = 1.6s
        assert abs(float(info["format"]["duration"]) - 1.6) < 0.5


# ---------------- 路由接线（app.py ↔ video.py） ----------------

@pytest.fixture
def client():
    import app as app_mod
    app_mod.app.config["TESTING"] = True
    return app_mod.app.test_client()


@pytest.fixture
def ready_deck(tmp_path, monkeypatch):
    import app as app_mod
    decks = tmp_path / "decks"
    decks.mkdir()
    (decks / "t.html").write_text("<div class='slide'>p</div>", encoding="utf-8")
    monkeypatch.setattr(app_mod, "DECKS_DIR", str(decks))
    monkeypatch.setattr(app_mod, "ANIMATION_DIR", str(tmp_path / "animation"))
    monkeypatch.setattr(app_mod, "VIDEOS_DIR", str(tmp_path / "videos"))
    with app_mod.lock:
        app_mod.state.update({"phase": "ready", "html_path": "/decks/t.html"})
    yield
    with app_mod.lock:
        app_mod.state.update({"phase": "idle", "html_path": None})


def test_export_video_queued(ready_deck, client):
    resp = client.post("/api/export_video")
    assert resp.status_code == 200 and resp.get_json()["queued"] is True
    if video.ffmpeg_path():  # 真 ffmpeg：等后台 worker 写完成日志
        import time
        import app as app_mod
        deadline = time.time() + 90
        while time.time() < deadline:
            with app_mod.lock:
                msgs = [l["msg"] for l in app_mod.state["log"]]
            if any("已导出视频" in m or "视频导出失败" in m for m in msgs):
                break
            time.sleep(0.3)
        assert any("已导出视频" in m for m in msgs)


def test_export_video_requires_ffmpeg(ready_deck, client, monkeypatch):
    import video as v
    monkeypatch.setattr(v, "ffmpeg_path", lambda: None)
    resp = client.post("/api/export_video")
    assert resp.status_code == 503
    assert "winget install ffmpeg" in resp.get_json()["error"]


def test_video_download_whitelist(ready_deck, client):
    import app as app_mod
    vids = os.path.join(app_mod.VIDEOS_DIR)
    os.makedirs(vids, exist_ok=True)
    open(os.path.join(vids, "a.mp4"), "wb").write(b"\x00\x00")
    open(os.path.join(vids, "evil.html"), "w").write("<script>")
    assert client.get("/videos/a.mp4").status_code == 200
    assert client.get("/videos/evil.html").status_code == 403


def test_artifacts_includes_video(ready_deck, client):
    import app as app_mod
    vids = app_mod.VIDEOS_DIR
    os.makedirs(vids, exist_ok=True)
    open(os.path.join(vids, "x_20260906.mp4"), "wb").write(b"\x00")
    body = client.get("/api/artifacts").get_json()
    assert any(a["kind"] == "mp4" for a in body["artifacts"])


# ---------------- 路由接线回归（功能测试缺陷-1/-2 修复） ----------------

@pytest.fixture
def client():
    import app as app_mod
    app_mod.app.config["TESTING"] = True
    return app_mod.app.test_client()


@pytest.fixture
def anim_state(tmp_path, monkeypatch):
    import app as app_mod
    decks = tmp_path / "decks"
    decks.mkdir()
    (decks / "t.html").write_text("<html></html>", encoding="utf-8")
    monkeypatch.setattr(app_mod, "DECKS_DIR", str(decks))
    monkeypatch.setattr(app_mod, "ANIMATION_DIR", str(tmp_path / "animation"))
    monkeypatch.setattr(app_mod, "VIDEOS_DIR", str(tmp_path / "videos"))
    with app_mod.lock:
        app_mod.state.update({"phase": "ready", "topic": "t", "html_path": "/decks/t.html"})
    yield
    with app_mod.lock:
        app_mod.state.update({"phase": "idle", "html_path": None})


def test_export_video_concurrent_guard(anim_state, client, monkeypatch):
    import threading
    import app as app_mod
    release = threading.Event()

    def blocking_shot(html, out_dir):
        os.makedirs(out_dir, exist_ok=True)
        release.wait(timeout=10)  # 拖住第一个 worker，制造并发窗口
        return []

    monkeypatch.setattr(app_mod.shot_mod, "shot_deck", blocking_shot)
    r1 = client.post("/api/export_video")
    assert r1.status_code == 200
    r2 = client.post("/api/export_video")
    assert r2.status_code == 409  # 同稿第二发被守卫拦截
    release.set()


def test_export_animation_quotes_special_topic(tmp_path, monkeypatch, client):
    import app as app_mod
    decks = tmp_path / "decks"
    decks.mkdir()
    # 稿名本身含 #/%（safe_topic 不清洗这两个字符，缺陷-1 的真实触发场景）
    name = "主题#号%测试_20260906_000000"
    (decks / f"{name}.html").write_text("<html></html>", encoding="utf-8")
    monkeypatch.setattr(app_mod, "DECKS_DIR", str(decks))

    def fake_shot(html, out_dir):
        os.makedirs(out_dir, exist_ok=True)
        return [str(tmp_path / "slide_1.png")]

    monkeypatch.setattr(app_mod.shot_mod, "shot_deck", fake_shot)
    with app_mod.lock:
        app_mod.state.update({"phase": "ready", "topic": name,
                              "html_path": f"/decks/{name}.html"})
    try:
        resp = client.post("/api/export_animation")
        assert resp.status_code == 200
        path = resp.get_json()["path"]
        assert "#" not in path and "%" in path  # # 已编码，不再截断 URL
        assert client.get(path).status_code == 200
    finally:
        with app_mod.lock:
            app_mod.state.update({"phase": "idle", "html_path": None})


def test_synthesize_rejects_fade_ge_seconds(tmp_path):
    """审计 L1：fade≥seconds 会让 ffmpeg 静默产出丢页坏视频，必须在入口拒绝。"""
    imgs = []
    for i in (1, 2):
        p = tmp_path / f"s{i}.png"
        _make_png(p)
        imgs.append(str(p))
    with pytest.raises(ValueError, match="转场时长"):
        video.synthesize(imgs, str(tmp_path / "o.mp4"), seconds=1.0, fade=2.0)
