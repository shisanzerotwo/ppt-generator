"""image_gen._download 单元测试（用 httpx.MockTransport，离线不真联网）。"""

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
