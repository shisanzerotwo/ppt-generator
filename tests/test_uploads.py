"""uploads.parse_pdf 的离线单元测试（PDF fixture 随用随建，不依赖外部文件）。"""

import io

import pytest
from PIL import Image

from uploads import parse_pdf


def _make_pdf(page_texts: list[str]) -> bytes:
    """手写最小合法多页 PDF：每页一行 Helvetica 文本，xref 偏移动态计算。

    比 fixture 文件好：文本与页数在用例内一目了然，不往仓库塞二进制文件。
    对象编号：1=Catalog 2=Pages 3..2+n=各页 3+n..2+2n=各页内容流 3+2n=Font。
    """
    n = len(page_texts)
    font_id = 3 + 2 * n
    objs: list[bytes] = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        f"<< /Type /Pages /Kids [{' '.join(f'{i + 3} 0 R' for i in range(n))}] /Count {n} >>".encode(),
    ]
    objs += [
        (
            f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] "
            f"/Contents {n + 3 + i} 0 R /Resources << /Font << /F1 {font_id} 0 R >> >> >>"
        ).encode()
        for i in range(n)
    ]
    for t in page_texts:
        content = f"BT /F1 24 Tf 72 720 Td ({t}) Tj ET".encode("ascii")
        objs.append(b"<< /Length %d >>\nstream\n" % len(content) + content + b"\nendstream")
    objs.append(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>")

    out = io.BytesIO()
    out.write(b"%PDF-1.4\n")
    offsets = []
    for i, body in enumerate(objs, start=1):
        offsets.append(out.tell())
        out.write(f"{i} 0 obj\n".encode() + body + b"\nendobj\n")
    xref_pos = out.tell()
    out.write(f"xref\n0 {len(objs) + 1}\n".encode())
    out.write(b"0000000000 65535 f \n")  # xref 首条：自由对象占位，不指向实际对象
    for off in offsets:
        out.write(f"{off:010d} 00000 n \n".encode())
    out.write(
        f"trailer\n<< /Size {len(objs) + 1} /Root 1 0 R >>\n"
        f"startxref\n{xref_pos}\n%%EOF\n".encode()
    )
    return out.getvalue()


def test_text_pdf_extracts_pages():
    """文本型 PDF 应逐页提取，页与页之间以换行符连接。"""
    data = _make_pdf(["Hello PPT one", "Hello PPT two"])
    assert parse_pdf(data) == "Hello PPT one\nHello PPT two"


def test_scanned_pdf_raises():
    """整份提不出文本（扫描件）应明确报错，而不是静默返回空文本。"""
    buf = io.BytesIO()
    Image.new("RGB", (120, 120), "white").save(buf, format="PDF")
    with pytest.raises(ValueError, match="扫描件"):
        parse_pdf(buf.getvalue())


def test_not_a_pdf_raises_open_error():
    """非 PDF 字节应报「PDF 无法打开」。"""
    with pytest.raises(ValueError, match="PDF 无法打开"):
        parse_pdf(b"this is definitely not a pdf at all")
