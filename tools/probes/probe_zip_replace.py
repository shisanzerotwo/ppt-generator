"""排查 `_patch_theme_zip` 里 `os.replace` 在 Windows 报"拒绝访问"的真因。

假设 A：python-pptx 的 `save()` 之后原文件仍被某个句柄占着（AV / 索引器 / 未关闭句柄）。
假设 B：`[Content_Types].xml` 之类的条目被 `zipfile` 隐式打开。
假设 C：纯粹的环境噪音（需要重试）。

做法：造一个最小 pptx，分别测
  1) 直接 os.replace（不经过任何 pptx 对象）
  2) 保存后**立即** replace（Presentation 对象仍在作用域）
  3) 保存后 replace（Presentation 已 del + gc）
  4) 失败后短暂重试是否成功
  5) 退路：不用 replace，改成"读全量 → 删原文件 → 原地重写"

跑法：.venv/Scripts/python.exe tools/probes/probe_zip_replace.py
"""

import gc
import os
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

OUT = os.path.abspath("output/spike/zipreplace")


def _make(path):
    from pptx import Presentation
    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    os.makedirs(os.path.dirname(path), exist_ok=True)
    prs.save(path)
    return prs


def _copy_zip(src, dst):
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            zout.writestr(item, zin.read(item.filename))


def _try(label, fn):
    try:
        fn()
        print(f"  [OK]   {label}")
        return True
    except OSError as exc:
        print(f"  [FAIL] {label} → {type(exc).__name__}: {exc}")
        return False


def main() -> int:
    os.makedirs(OUT, exist_ok=True)

    print("== 1) 不经 pptx 对象：纯 zip 复制 + replace ==")
    a = os.path.join(OUT, "a.pptx")
    _make(a)
    tmp = a + ".tmp"
    _copy_zip(a, tmp)
    _try("纯 replace", lambda: os.replace(tmp, a))

    print("\n== 2) 保存后立即 replace（prs 仍在作用域）==")
    b = os.path.join(OUT, "b.pptx")
    prs = _make(b)
    tmp = b + ".tmp"
    _copy_zip(b, tmp)
    _try("prs 存活时 replace", lambda: os.replace(tmp, b))

    print("\n== 3) del prs + gc 之后再 replace ==")
    c = os.path.join(OUT, "c.pptx")
    prs2 = _make(c)
    del prs2
    gc.collect()
    tmp = c + ".tmp"
    _copy_zip(c, tmp)
    _try("gc 之后 replace", lambda: os.replace(tmp, c))

    print("\n== 4) 连做 5 次 replace（看是否偶发）==")
    d = os.path.join(OUT, "d.pptx")
    _make(d)
    for i in range(5):
        tmp = d + f".tmp{i}"
        _copy_zip(d, tmp)
        _try(f"第 {i + 1} 次", lambda t=tmp: os.replace(t, d))

    print("\n== 5) 退路：读全量 → 删原文件 → 原地重写（不用 replace）==")
    e = os.path.join(OUT, "e.pptx")
    _make(e)

    def rewrite():
        with zipfile.ZipFile(e) as zin:
            items = [(it, zin.read(it.filename)) for it in zin.infolist()]
        os.remove(e)
        with zipfile.ZipFile(e, "w", zipfile.ZIP_DEFLATED) as zout:
            for it, data in items:
                zout.writestr(it, data)

    _try("删+重写", rewrite)
    print(f"  重写后大小 = {os.path.getsize(e)}")

    print("\n== 6) 复现真实场景：build_deck_pptx 连做 6 次（含 zip 后处理）==")
    import pptx_out
    src = os.path.abspath("output/fixtures/fixtures_cover.pptx")
    if not os.path.isfile(src):
        print(f"  （跳过：缺 {src}）")
        return 0
    import pptx_io
    deck, _ = pptx_io.read_pages(src, export_width_px=1920)
    fails = 0
    for i in range(6):
        out = os.path.join(OUT, f"real_{i}.pptx")
        try:
            pptx_out.build_deck_pptx(deck, out)
            print(f"  [OK]   第 {i + 1} 次 → {os.path.getsize(out)} 字节")
        except OSError as exc:
            fails += 1
            print(f"  [FAIL] 第 {i + 1} 次 → {type(exc).__name__}: {exc}")
    print(f"  → 6 次里失败 {fails} 次（期望 0）")
    return 0 if fails == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
