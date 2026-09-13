"""M2 对照探针：`read_pages` 对 zip bomb 有没有闸门。

背景（AUDIT_REPORT §2 M2）：把 64 MiB 全零塞进 `[Content_Types].xml`（包只有 109 KB，
压缩比 601:1），`Presentation()` 照单全收 —— 字节在内存里展开完才发现 XML 是坏的。
契约 §8.3 点名要审这一项。

本探针**只读、不启动任何外部程序**（纯内存构造 + tempfile），所以不需要 COM 闸门。

判定：
  修复前 → 没有闸门：样本文件体积 < 200 KB 却让 64 MiB 进内存（打印实际 ratio 佐证）
  修复后 → 抛出 PptxError，且**错误码正确**、message 里带体积

跑法：
    .venv/Scripts/python.exe tools/probes/fix_m2_zipbomb.py
"""

import os
import sys
import tempfile
import tracemalloc
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", ".."))
sys.path.insert(0, ROOT)

BOMB_PAYLOAD = 64 * 1024 * 1024  # 64 MiB 全零
LEGIT = os.path.join(ROOT, "output", "b_multislide.pptx")


def build_bomb(path: str, payload_bytes: int = BOMB_PAYLOAD) -> tuple:
    """造一个 bomb 包：`[Content_Types].xml` 塞满零字节。

    结构与真 pptx 保持一致（`_rels/.rels` + `ppt/presentation.xml`），
    只是把必读的 Content_Types 换成巨物 —— `Presentation()` 打开时必读它。
    """
    parts = {
        "[Content_Types].xml": b"\x00" * payload_bytes,
        "_rels/.rels": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
            'relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>'
        ).encode("utf-8"),
        "ppt/presentation.xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>'
        ).encode("utf-8"),
    }
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)
    return os.path.getsize(path), parts


def describe(path: str) -> dict:
    """读 zip 目录（不解压）报声明体积与最差压缩比 —— 这正是闸门该看的东西。"""
    with zipfile.ZipFile(path) as z:
        total = 0
        worst = ("", 0.0, 0)
        for info in z.infolist():
            total += info.file_size
            if info.compress_size > 0:
                ratio = info.file_size / info.compress_size
                if ratio > worst[1]:
                    worst = (info.filename, ratio, info.file_size)
    return {"total_uncompressed": total, "worst": worst}


def main() -> int:
    tmp = tempfile.mkdtemp(prefix="m2probe_")
    try:
        bomb = os.path.join(tmp, "bomb.pptx")
        size, _ = build_bomb(bomb)
        info = describe(bomb)
        name, ratio, raw = info["worst"]
        print(f"样本：{bomb}")
        print(f"  文件体积           = {size:,} B（{size / 1024:.1f} KB）")
        print(f"  解压后声明总体积   = {info['total_uncompressed']:,} B"
              f"（{info['total_uncompressed'] / 1024 / 1024:.1f} MB）")
        print(f"  最差单条压缩比     = {ratio:.0f}:1（{name}，{raw / 1024 / 1024:.1f} MB）")

        print("\n== 调 read_pages（量内存峰值：字节到底有没有被展开）==")
        import pptx_io
        tracemalloc.start()
        try:
            pptx_io.read_pages(bomb)
            outcome = "**没有闸门** —— read_pages 未拒绝"
        except pptx_io.PptxError as exc:
            outcome = f"被拒 → code={exc.code}\n        message={exc.message}\n        hint={exc.hint}"
        except Exception as exc:  # noqa: BLE001
            outcome = f"抛了别的异常 → {type(exc).__name__}: {exc}"
        _, peak = tracemalloc.get_traced_memory()
        tracemalloc.stop()
        print(f"  结果：{outcome}")
        print(f"  Python 侧内存峰值 = {peak / 1024 / 1024:.1f} MB"
              f"（文件只有 {size / 1024:.1f} KB）")

        gate_small = peak < 8 * 1024 * 1024
        if "code=PPTX_UNREADABLE" in outcome and gate_small:
            print("  判定：**FIXED** —— 闸门在读之前就拦下，字节没进内存")
        elif gate_small:
            print("  判定：闸门拦下了（但错误码不是 PPTX_UNREADABLE，请核对文案）")
        else:
            print(f"  判定：**REPRODUCED** —— 没有闸门，{peak / 1024 / 1024:.0f} MB 已展开进内存")

        print("\n== 正常稿不受影响 ==")
        if os.path.isfile(LEGIT):
            import pptx_io
            deck, _ = pptx_io.read_pages(LEGIT)
            print(f"  b_multislide.pptx → {len(deck.pages)} 页（应仍为 10）")
        else:
            print(f"  （跳过：缺 {LEGIT}）")
        return 0
    finally:
        import shutil
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    raise SystemExit(main())
