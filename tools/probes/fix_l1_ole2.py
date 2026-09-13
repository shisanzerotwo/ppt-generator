"""L1 对照探针：OLE2 复合文档不都是"加密的 pptx"，老 .ppt 也不是。

背景（AUDIT_REPORT §2 L1）：`_classify_bad_package` 只看文件头魔数 `D0CF11E0`，
而**老 `.ppt` / `.doc` / `.xls` 全都是 OLE2** → 用户拿一份 `.ppt` 过来，被告知
"PPTX 已加密，请先用 PowerPoint 去掉打开密码"，而那份文件**根本没有密码** ——
提示不可执行，还把契约 §9 里本该给的正确提示（"确认是 .pptx（非 .ppt/.pdf）"）抢走了。

真正的加密 OOXML 也是 OLE2，但它里面有一个特征流 `EncryptedPackage`（CFB 目录里以
UTF-16LE 存名）。所以判据是**两层**：扩展名 + 该特征流在不在。

样本（纯字节构造，不启动任何外部程序 → 无需 COM 安全闸门）：
  a) legacy.ppt       OLE2 + `PowerPoint Document` 流名
  b) encrypted.pptx   OLE2 + `EncryptedPackage` 流名
  c) renamed.pptx     OLE2 + `PowerPoint Document` 流名，但**扩展名是 .pptx**（改名）

判定：
  修复前 → 三者**一律** PPTX_ENCRYPTED（"请去掉打开密码"，对 a/c 不可执行）
  修复后 → a) 与 c) 得 PPTX_UNREADABLE + "另存为 .pptx"；b) 保持 PPTX_ENCRYPTED

跑法：.venv/Scripts/python.exe tools/probes/fix_l1_ole2.py
"""

import os
import shutil
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", ".."))
sys.path.insert(0, ROOT)

OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def make_ole2(path: str, stream_names: list) -> str:
    """造一个够真的 OLE2 样本：头 + 若干个 CFB 目录流名（UTF-16LE）。

    分类器只依赖文件头与流名，所以这样就足以驱动它 —— 不需要真的实现 CFB。
    """
    blob = bytearray(OLE_MAGIC)
    blob += b"\x00" * 504                          # 头占满 512 字节
    for name in stream_names:
        blob += name.encode("utf-16-le") + b"\x00\x00"
        blob += b"\x00" * 64
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "wb") as f:
        f.write(bytes(blob))
    return path


def classify(path: str) -> tuple:
    import pptx_io
    try:
        pptx_io.read_pages(path)
        return ("（没报错？不应发生）", "")
    except pptx_io.PptxError as exc:
        return (exc.code, exc.hint)
    except Exception as exc:  # noqa: BLE001
        return (f"裸异常 {type(exc).__name__}", str(exc))


def main() -> int:
    tmp = tempfile.mkdtemp(prefix="l1probe_")
    try:
        cases = [
            ("a) 老 .ppt（合法旧格式）", make_ole2(
                os.path.join(tmp, "legacy.ppt"), ["PowerPoint Document", "Current User"])),
            ("b) 真加密 pptx", make_ole2(
                os.path.join(tmp, "encrypted.pptx"), ["EncryptionInfo", "EncryptedPackage"])),
            ("c) 老 .ppt 改名成 .pptx", make_ole2(
                os.path.join(tmp, "renamed.pptx"), ["PowerPoint Document"])),
        ]
        print(f"样本目录：{tmp}\n")
        results = {}
        for label, path in cases:
            code, hint = classify(path)
            results[label] = code
            print(f"{label}")
            print(f"    → code={code}")
            print(f"      hint={hint}")

        print("\n" + "-" * 70)
        legacy_ok = results["a) 老 .ppt（合法旧格式）"] == "PPTX_UNREADABLE"
        enc_ok = results["b) 真加密 pptx"] == "PPTX_ENCRYPTED"
        renamed_ok = results["c) 老 .ppt 改名成 .pptx"] == "PPTX_UNREADABLE"
        if legacy_ok and enc_ok and renamed_ok:
            print("判定：**FIXED** —— 老格式给「另存为 .pptx」的可执行指引，"
                  "真加密仍报 PPTX_ENCRYPTED")
            return 0
        if all(v == "PPTX_ENCRYPTED" for v in results.values()):
            print("判定：**REPRODUCED** —— 三个样本一律被当成「已加密」，"
                  "对老 .ppt 的指引不可执行（L1 成立）")
            return 0
        print(f"判定：**部分修复/异常** → {results}")
        return 1
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    raise SystemExit(main())
