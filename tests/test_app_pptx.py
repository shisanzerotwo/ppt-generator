"""步 8：工作台 pptx 入口（`/api/pptx/import` + `/pptx/<path>`）与前端接线。

分三层：
- **不需要 PowerPoint**：参数校验、路径穿越、并发守卫、前端接线（按钮/输入框/JS 无语法错）；
- **需要 PowerPoint**：上传 → 后台导入 → 播放器可访问的端到端（无 PowerPoint 自动 skip）。

前端接线用真浏览器加载 `/` 并挂 `pageerror` —— 大段内联 JS 光靠肉眼看不出来。
"""

import io
import json
import os
import time
from urllib.parse import quote

import pytest
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt

import app as app_mod
import pptx_io

needs_pptx = pytest.mark.skipif(not pptx_io.powerpoint_available(),
                                reason="本机没有 PowerPoint（COM 导入不可用）")


@pytest.fixture
def client():
    app_mod.app.config["TESTING"] = True
    with app_mod.app.test_client() as c:
        yield c


@pytest.fixture(autouse=True)
def _isolate_output(tmp_path, monkeypatch):
    """把 pptx 产物根隔离到临时目录，别污染开发者的 output/。"""
    monkeypatch.setattr(app_mod, "PPTX_SRC_DIR", str(tmp_path / "pptx_src"))
    monkeypatch.setattr(app_mod, "PPTX_UPLOAD_DIR", str(tmp_path / "pptx_uploads"))
    monkeypatch.setattr(app_mod, "OUTPUT_DIR", str(tmp_path))
    yield


def _pptx_bytes(pages=2):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    for i in range(pages):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(1))
        box.text_frame.text = f"第{i + 1}页标题"
        box.text_frame.paragraphs[0].runs[0].font.size = Pt(32)
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf


def _wait_pptx(timeout=180.0):
    """等后台 worker 收工，返回 state["pptx"]。"""
    deadline = time.time() + timeout
    while time.time() < deadline:
        with app_mod.lock:
            st = dict(app_mod.state["pptx"])
        if st["status"] in ("ready", "error"):
            return st
        time.sleep(0.5)
    raise AssertionError("pptx 导入超时")


# ---------------------------------------------------------------- 参数与安全

def test_rejects_missing_file(client):
    r = client.post("/api/pptx/import", data={})
    assert r.status_code == 400
    assert "未收到文件" in r.get_json()["error"]


def test_rejects_non_pptx(client):
    r = client.post("/api/pptx/import",
                    data={"file": (io.BytesIO(b"hello"), "note.txt")},
                    content_type="multipart/form-data")
    assert r.status_code == 400
    assert ".pptx" in r.get_json()["error"]


def test_rejects_oversize_upload(client, monkeypatch):
    monkeypatch.setattr(app_mod, "MAX_PPTX_BYTES", 1024)
    r = client.post("/api/pptx/import",
                    data={"file": (_pptx_bytes(1), "big.pptx")},
                    content_type="multipart/form-data")
    assert r.status_code == 413


def test_pptx_static_route_rejects_traversal(client):
    for bad in ["../app.py", "..%2fapp.py", "%2e%2e/app.py"]:
        assert client.get(f"/pptx/{bad}").status_code in (400, 404)


def test_pptx_static_route_serves_existing_file(client, tmp_path):
    target = tmp_path / "pptx_src" / "a" / "player"
    target.mkdir(parents=True)
    (target / "index.html").write_text("<h1>hi</h1>", encoding="utf-8")
    r = client.get("/pptx/a/player/index.html")
    assert r.status_code == 200 and b"hi" in r.data


# ---------------------------------------------------------------- 前端接线

def test_index_has_pptx_entry(client):
    html = client.get("/").get_data(as_text=True)
    assert 'id="pptx-input"' in html
    assert 'accept=".pptx"' in html
    assert "doPptxImport" in html
    assert "/api/pptx/import" in html
    assert 'id="pptx-status"' in html


def test_index_js_runs_without_errors(client):
    """真浏览器加载首页：内联 JS 必须零 pageerror（大段 JS 肉眼看不出来）。

    必须起**真实端口**再打开：用 `page.set_content` 的话 base 是 about:blank，
    页面里所有相对 fetch 都会抛 "Failed to parse URL"，测出来的是测试环境的问题
    而不是页面本身的问题。
    """
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        pytest.skip("未安装 playwright")
    import threading

    from werkzeug.serving import make_server

    import shot as shot_mod

    server = make_server("127.0.0.1", 0, app_mod.app)
    port = server.server_port
    threading.Thread(target=server.serve_forever, daemon=True).start()
    try:
        with sync_playwright() as p:
            try:
                browser = shot_mod._launch_browser(p)
            except RuntimeError as e:
                pytest.skip(f"无可用浏览器：{e}")
            try:
                page = browser.new_page()
                errors = []
                page.on("pageerror", lambda e: errors.append(str(e)))
                page.goto(f"http://127.0.0.1:{port}/", wait_until="load")
                page.wait_for_timeout(800)   # 让 refresh()/loadDecks()/pptxRefresh() 跑完
                assert errors == [], f"首页 JS 报错：{errors}"
                assert page.locator("#pptx-input").count() == 1
                page.click('button:has-text("导入 PPTX 高亮讲解")')
            finally:
                browser.close()
    finally:
        server.shutdown()


# ---------------------------------------------------------------- 端到端（需 COM）

@needs_pptx
def test_import_creates_player_and_status(client):
    r = client.post("/api/pptx/import",
                    data={"file": (_pptx_bytes(2), "演示稿.pptx")},
                    content_type="multipart/form-data")
    assert r.status_code == 200, r.get_data(as_text=True)
    assert r.get_json()["ok"] is True

    st = _wait_pptx()
    assert st["status"] == "ready", st
    assert st["pages"] == 2 and st["units"] >= 2
    assert st["player"].startswith("/pptx/")

    out = st["out"]
    assert os.path.isfile(os.path.join(out, "deck.json"))
    assert os.path.isfile(os.path.join(out, "bg", "slide_1.png"))
    assert os.path.isfile(os.path.join(out, "bg", "slide_2.png"))

    page = client.get(st["player"])
    assert page.status_code == 200
    html = page.get_data(as_text=True)
    assert "window.hl" in html and "<iframe" not in html

    # 播放器引用的底图必须真的能取到（L4：缺图会安静出黑帧）
    name = quote(st["name"])
    bg = client.get(f"/pptx/{name}/bg/slide_1.png")
    assert bg.status_code == 200 and len(bg.data) > 0
    with Image.open(io.BytesIO(bg.data)) as im:
        assert im.size[0] > 100  # 真的是导出出来的图，不是空占位


@needs_pptx
def test_concurrent_import_is_rejected(client):
    """同稿并发守卫：state 处于 running 时再传一律 409（确定性，不靠竞态）。"""
    with app_mod.lock:
        app_mod.state["pptx"]["status"] = "running"
    try:
        r = client.post("/api/pptx/import",
                        data={"file": (_pptx_bytes(1), "并发.pptx")},
                        content_type="multipart/form-data")
        assert r.status_code == 409
    finally:
        with app_mod.lock:
            app_mod.state["pptx"]["status"] = "idle"


@needs_pptx
def test_import_of_corrupt_pptx_reports_chinese_error(client):
    """病态输入要落成中文 error（不是让 worker 静默死掉），前端才有话可说。"""
    r = client.post("/api/pptx/import",
                    data={"file": (io.BytesIO(b"not a zip at all"), "坏稿.pptx")},
                    content_type="multipart/form-data")
    assert r.status_code == 200
    st = _wait_pptx()
    assert st["status"] == "error"
    assert st["error"] and "Traceback" not in st["error"]


@needs_pptx
def test_uploaded_source_is_not_served(client):
    """原始上传稿放在被服务目录之外，/pptx/ 不该能取到它。"""
    client.post("/api/pptx/import",
                data={"file": (_pptx_bytes(1), "源稿.pptx")},
                content_type="multipart/form-data")
    _wait_pptx()
    assert client.get(f"/pptx/{quote('源稿.src.pptx')}").status_code == 404
