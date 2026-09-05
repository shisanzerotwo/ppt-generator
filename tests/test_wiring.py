"""本轮新功能的路由接线层测试（测试成果分析报告的缺口闭合）。

模块层（qa/quality/uploads/theme 替换）已有专项测试；这里验证 app.py 的接线：
导出门禁 422 拦截、配图缓存命中/降级、_gen_image_page 分支、export_html 同源、
import_file 的 PDF 分发、/api/quality 端点。全部不触网（mock LLM/生图）。
"""

import io
import os

import pytest

import app as app_mod
import image_gen


@pytest.fixture
def client():
    return app_mod.app.test_client()


@pytest.fixture
def ready_deck(tmp_path, monkeypatch):
    """ready 态两页 + 输出目录指向 tmp，测后恢复 state。"""
    monkeypatch.setattr(app_mod, "OUTPUT_DIR", str(tmp_path))
    monkeypatch.setattr(app_mod, "IMAGES_DIR", str(tmp_path / "images"))
    with app_mod.lock:
        app_mod.state["phase"] = "ready"
        app_mod.state["topic"] = "测试主题"
        app_mod.state["slides"] = [
            {"type": "content", "title": "页一", "points": ["要点：1"],
             "image_prompt": "", "image": None, "imageStatus": "skipped",
             "review": {"ok": True, "reason": "", "tries": 0}},
            {"type": "content", "title": "页二", "points": ["要点：2"],
             "image_prompt": "", "image": None, "imageStatus": "skipped",
             "review": {"ok": True, "reason": "", "tries": 0}},
        ]
    yield tmp_path
    with app_mod.lock:
        app_mod.state["phase"] = "idle"
        app_mod.state["slides"] = []
        app_mod.state["topic"] = ""


# ---------------- 导出门禁接线（缺口清单 ①，风险最高） ----------------

def test_export_blocked_on_qa_errors(ready_deck, client, monkeypatch):
    monkeypatch.setattr(app_mod.qa_mod, "check_pptx", lambda path, expected_pages:
                        {"errors": ["第1页 形状越界"], "warnings": [], "used_estimate": False})
    resp = client.post("/api/export")
    assert resp.status_code == 422
    body = resp.get_json()
    assert body["ok"] is False and "report" in body
    # 拦截后坏文件不留产物区
    assert not list(ready_deck.glob("*.pptx"))


def test_export_passes_with_warnings(ready_deck, client, monkeypatch):
    monkeypatch.setattr(app_mod.qa_mod, "check_pptx", lambda path, expected_pages:
                        {"errors": [], "warnings": ["第1页 文字可能溢出"], "used_estimate": False})
    resp = client.post("/api/export")
    assert resp.status_code == 200
    body = resp.get_json()
    assert body["ok"] is True and body["report"]["warnings"]


def test_export_rejects_when_not_ready(client):
    with app_mod.lock:
        app_mod.state["phase"] = "images"
    try:
        assert client.post("/api/export").status_code == 409
    finally:
        with app_mod.lock:
            app_mod.state["phase"] = "idle"


# ---------------- 配图缓存：命中 / 读失败降级 / key 掺 points（缺口 ②） ----------------

def test_image_cache_hit_skips_generation(tmp_path, monkeypatch):
    monkeypatch.setattr(app_mod, "IMAGES_DIR", str(tmp_path))
    calls = []
    monkeypatch.setattr(image_gen, "generate_image", lambda p, s: calls.append(p))
    with app_mod.lock:
        app_mod.state["slides"] = [{"title": "t", "points": ["p：1"], "image_prompt": "画一棵树",
                                    "image": None, "imageStatus": "pending",
                                    "review": {"ok": True, "reason": "", "tries": 0}}]
    try:
        cache = app_mod._image_cache_path("t", "画一棵树", ["p：1"])
        os.makedirs(os.path.dirname(cache), exist_ok=True)
        with open(cache, "wb") as f:
            f.write(b"\x89PNG fake")  # 预放缓存（内容不作校验，本用例只看接线）
        app_mod._gen_images()
        assert calls == []  # 命中缓存：生图一次都不调
        assert app_mod.state["slides"][0]["imageStatus"] == "done"
    finally:
        with app_mod.lock:
            app_mod.state["slides"] = []


def test_image_cache_read_failure_degrades_to_generate(tmp_path, monkeypatch):
    monkeypatch.setattr(app_mod, "IMAGES_DIR", str(tmp_path))
    monkeypatch.setattr(image_gen, "generate_image", lambda p, s: open(s, "wb").write(b"\x89PNG"))
    monkeypatch.setattr(app_mod.critic, "review_image", lambda t, pts, p: {"ok": True, "reason": "", "advice": ""})
    with app_mod.lock:
        app_mod.state["slides"] = [{"title": "t", "points": ["p：1"], "image_prompt": "画一棵树",
                                    "image": None, "imageStatus": "pending",
                                    "review": {"ok": True, "reason": "", "tries": 0}}]
    try:
        cache = app_mod._image_cache_path("t", "画一棵树", ["p：1"])
        os.makedirs(cache, exist_ok=True)  # 缓存位是个目录 → copyfile 抛 OSError → 降级现场生成（审计 L10）
        app_mod._gen_images()
        assert app_mod.state["slides"][0]["imageStatus"] == "done"
    finally:
        with app_mod.lock:
            app_mod.state["slides"] = []


def test_cache_key_changes_with_points(tmp_path, monkeypatch):
    monkeypatch.setattr(app_mod, "IMAGES_DIR", str(tmp_path))
    k1 = app_mod._image_cache_path("t", "p", ["要点：1"])
    k2 = app_mod._image_cache_path("t", "p", ["要点：2"])
    assert k1 != k2  # 审计 L6：改要点后不得命中旧图


# ---------------- _gen_image_page 分支（缺口 ③） ----------------

def test_gen_image_page_skip_and_failed(tmp_path, monkeypatch):
    monkeypatch.setattr(app_mod, "IMAGES_DIR", str(tmp_path))
    monkeypatch.setattr(image_gen, "generate_image", lambda p, s: (_ for _ in ()).throw(RuntimeError("boom")))
    with app_mod.lock:
        app_mod.state["slides"] = [
            {"title": "无图页", "points": [], "image_prompt": "", "image": None,
             "imageStatus": "pending", "review": {"ok": True, "reason": "", "tries": 0}},
            {"title": "炸图页", "points": [], "image_prompt": "画点啥", "image": None,
             "imageStatus": "pending", "review": {"ok": True, "reason": "", "tries": 0}},
        ]
    try:
        app_mod._gen_images()
        assert app_mod.state["slides"][0]["imageStatus"] == "skipped"   # 无提示词跳过
        assert app_mod.state["slides"][1]["imageStatus"] == "failed"    # 生图异常标 failed 不炸整单
    finally:
        with app_mod.lock:
            app_mod.state["slides"] = []


# ---------------- export_html 同源分支（缺口 ④） ----------------

def test_export_html_same_source(ready_deck, tmp_path, monkeypatch, client):
    decks = tmp_path / "decks"
    decks.mkdir()
    (decks / "t.html").write_text("<html></html>", encoding="utf-8")
    monkeypatch.setattr(app_mod, "DECKS_DIR", str(decks))
    with app_mod.lock:
        app_mod.state["html_path"] = "/decks/t.html"
    try:
        resp = client.post("/api/export_html")
        assert resp.status_code == 200
        assert resp.get_json()["same_source"] is True
    finally:
        with app_mod.lock:
            app_mod.state["html_path"] = None


# ---------------- import_file 的 PDF 分发（缺口 ⑤） ----------------

def _make_pdf(text="Hello Wiring"):
    """最小合法单页 PDF（xref 偏移动态计算），与 test_uploads 同思路。"""
    stream = f"BT /F1 18 Tf 72 720 Td ({text}) Tj ET".encode()
    objs = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] "
        b"/Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>",
        b"<< /Length " + str(len(stream)).encode() + b" >>\nstream\n" + stream + b"\nendstream",
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    ]
    out = b"%PDF-1.4\n"
    offsets = []
    for i, body in enumerate(objs, 1):
        offsets.append(len(out))
        out += f"{i} 0 obj\n".encode() + body + b"\nendobj\n"
    xref = len(out)
    out += f"xref\n0 {len(objs) + 1}\n0000000000 65535 f \n".encode()
    for off in offsets:
        out += f"{off:010d} 00000 n \n".encode()
    out += (f"trailer\n<< /Size {len(objs) + 1} /Root 1 0 R >>\n"
            f"startxref\n{xref}\n%%EOF").encode()
    return out


def test_import_file_pdf_routes_to_parser(ready_deck, client, monkeypatch):
    monkeypatch.setattr(app_mod.outline, "generate_outline_from_text",
                        lambda text: (_ for _ in ()).throw(RuntimeError("stop-here")))
    # 文本须 ≥30 字，否则被「文档内容过短」拦在解析之后
    resp = client.post("/api/import_file", data={
        "file": (io.BytesIO(_make_pdf("Hello Wiring, this is a long enough body text.")), "t.pdf")},
        content_type="multipart/form-data")
    assert resp.status_code == 200 and resp.get_json()["ok"] is True


def test_import_file_bad_pdf_400(ready_deck, client):
    resp = client.post("/api/import_file", data={
        "file": (io.BytesIO(b"not a pdf at all"), "t.pdf")}, content_type="multipart/form-data")
    assert resp.status_code == 400
    assert "文件解析失败" in resp.get_json()["error"]


# ---------------- /api/quality 端点（缺口 ⑥） ----------------

def test_quality_endpoint_reports_duplicates(ready_deck, client):
    with app_mod.lock:
        app_mod.state["slides"][1]["title"] = "页一"  # 与页一同题 → 重复
    resp = client.get("/api/quality")
    assert resp.status_code == 200
    assert resp.get_json()["duplicates"]
