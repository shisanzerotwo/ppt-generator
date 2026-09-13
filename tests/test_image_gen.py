"""image_gen._download / _save_image 单元测试（用 httpx.MockTransport，离线不真联网）。"""

import base64
import io
import types

import httpx
import pytest

import image_gen


def _patch_client(monkeypatch, handler):
    """让 image_gen 里的 httpx.Client 使用 MockTransport。"""
    real_client = image_gen.httpx.Client  # 先捕获原 Client，避免 patch 后自引用

    def factory(*args, **kwargs):
        kwargs.setdefault("transport", httpx.MockTransport(handler))
        return real_client(*args, **kwargs)

    monkeypatch.setattr(image_gen.httpx, "Client", factory)


def test_download_non_image_response_raises(tmp_path, monkeypatch):
    def handler(request):
        return httpx.Response(200, headers={"content-type": "text/html"}, content=b"<html>error</html>")

    _patch_client(monkeypatch, handler)
    with pytest.raises(ValueError, match="非图片响应"):
        image_gen._download("https://example.com/x", str(tmp_path / "x.png"))


def test_download_missing_content_type_raises(tmp_path, monkeypatch):
    def handler(request):
        return httpx.Response(200, content=b"not an image")

    _patch_client(monkeypatch, handler)
    with pytest.raises(ValueError, match="非图片响应"):
        image_gen._download("https://example.com/x", str(tmp_path / "x.png"))


def test_download_image_success(tmp_path, monkeypatch):
    png_bytes = b"\x89PNG\r\n\x1a\n fake-image-bytes"

    def handler(request):
        return httpx.Response(200, headers={"content-type": "image/png"}, content=png_bytes)

    _patch_client(monkeypatch, handler)
    out = tmp_path / "img.png"
    result = image_gen._download("https://example.com/img.png", str(out))

    assert result == str(out)
    assert out.read_bytes() == png_bytes


def test_download_http_error_raises(tmp_path, monkeypatch):
    def handler(request):
        return httpx.Response(404, headers={"content-type": "image/png"})

    _patch_client(monkeypatch, handler)
    with pytest.raises(httpx.HTTPStatusError):
        image_gen._download("https://example.com/missing.png", str(tmp_path / "x.png"))


def test_save_image_url_takes_precedence(tmp_path, monkeypatch):
    """有 url 走下载（Agnes 行为不变）。"""
    png_bytes = b"\x89PNG\r\n\x1a\n from-url"

    def handler(request):
        return httpx.Response(200, headers={"content-type": "image/png"}, content=png_bytes)

    _patch_client(monkeypatch, handler)
    out = tmp_path / "a.png"
    result = image_gen._save_image(types.SimpleNamespace(url="https://example.com/a.png"), str(out))

    assert result == str(out)
    assert out.read_bytes() == png_bytes


def test_save_image_b64_json_fallback(tmp_path):
    """只回内联 b64 的网关（如 OmniRoute）：无 url 时必须仍能落盘成图。"""
    raw = b"\x89PNG\r\n\x1a\n inline-bytes"
    out = tmp_path / "b.png"
    item = types.SimpleNamespace(url=None, b64_json=base64.b64encode(raw).decode())

    assert image_gen._save_image(item, str(out)) == str(out)
    assert out.read_bytes() == raw


def test_save_image_neither_url_nor_b64_raises(tmp_path):
    with pytest.raises(ValueError, match="既无 url 也无 b64_json"):
        image_gen._save_image(types.SimpleNamespace(url=None, b64_json=None), str(tmp_path / "c.png"))


def test_save_image_transcodes_webp_to_png(tmp_path):
    """AI Horde 回 WebP，而 python-pptx 不认 WebP —— 落在 .png 路径上的必须真转成 PNG。"""
    from PIL import Image

    buf = io.BytesIO()
    Image.new("RGB", (4, 4), (255, 0, 0)).save(buf, "WEBP")
    webp = buf.getvalue()
    assert webp[:4] == b"RIFF"

    out = tmp_path / "w.png"
    image_gen._save_image(types.SimpleNamespace(url=None, b64_json=base64.b64encode(webp).decode()),
                          str(out))

    assert out.read_bytes().startswith(b"\x89PNG")
    with Image.open(out) as im:
        assert im.format == "PNG"
        assert im.size == (4, 4)


def test_download_transcodes_webp_to_png(tmp_path, monkeypatch):
    """url 路径同样兜底：网关回 WebP 也不能把 WebP 字节塞进 .png。"""
    from PIL import Image

    buf = io.BytesIO()
    Image.new("RGB", (3, 3), (0, 0, 255)).save(buf, "WEBP")
    webp = buf.getvalue()

    def handler(request):
        return httpx.Response(200, headers={"content-type": "image/webp"}, content=webp)

    _patch_client(monkeypatch, handler)
    out = tmp_path / "d.png"
    image_gen._download("https://example.com/d.webp", str(out))

    assert out.read_bytes().startswith(b"\x89PNG")
