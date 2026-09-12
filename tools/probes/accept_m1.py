"""M1 验收：对 output/b_multislide.pptx 读形状 + COM 导出 10 张 1920x1080 PNG。

跑法：.venv/Scripts/python.exe tools/probes/accept_m1.py
"""

import os
import statistics
import sys
import time
from collections import Counter

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

import pptx_io  # noqa: E402

SRC = os.path.abspath("output/b_multislide.pptx")
OUT = os.path.abspath("output/spike/m1")


def main() -> int:
    print(f"素材：{SRC}")
    t0 = time.perf_counter()
    deck, skipped = pptx_io.read_pages(SRC)
    read_s = time.perf_counter() - t0

    print(f"页数 = {len(deck.pages)}（期望 10）")
    top = [s for p in deck.pages for s in p["shapes"]]
    print(f"顶层形状数 = {len(top)}")
    kinds = Counter(s.kind for s in top)
    print(f"kind 分布 = {dict(kinds)}")
    print(f"其中文本形状 = {kinds.get('text', 0)}（契约 §M1 口径：40 个文本框 + 1 图表 = 41 形状）")
    print(f"画布 EMU = {deck.width_emu} x {deck.height_emu}")
    print(f"读形状耗时 = {read_s:.3f}s")
    print(f"skipped = {len(skipped)} 条，reason 分布 = {dict(Counter(x['reason'] for x in skipped))}")

    none_coords = [s.name for p in deck.pages for s in p["shapes"]
                   if None in (s.left_emu, s.top_emu, s.width_emu, s.height_emu)]
    print(f"坐标为 None 的顶层形状 = {none_coords}（期望 []）")

    print("\n== COM 导出 ==")
    t0 = time.perf_counter()
    paths = pptx_io.export_pages(SRC, OUT, width=1920)
    total_s = time.perf_counter() - t0
    from PIL import Image
    sizes = []
    for p in paths:
        with Image.open(p) as im:
            sizes.append(im.size)
    print(f"导出文件数 = {len(paths)}（期望 10）")
    print(f"尺寸集合 = {set(sizes)}（期望 {{(1920, 1080)}}）")
    print(f"总耗时 = {total_s:.2f}s → 平均 {total_s / len(paths):.3f}s/页（验收 ≤1.5s/页）")
    print(f"页面一致性 = {len(deck.pages) == len(paths)}")

    # 记录逐页耗时（第二次导出，含缓存效应，仍逐页计时）
    print("\n== 逐页导出计时（重导一次，热路径） ==")
    times = _timed_rexport()
    print(f"逐页 = {[f'{t:.3f}s' for t in times]}")
    print(f"中位 = {statistics.median(times):.3f}s/页，最大 = {max(times):.3f}s/页")
    return 0


def _timed_rexport():
    import pythoncom
    import win32com.client

    had = pptx_io._powerpoint_running()
    pythoncom.CoInitialize()
    times = []
    app = pres = None
    try:
        app = win32com.client.DispatchEx("PowerPoint.Application")
        pre = app.Presentations.Count
        pres = app.Presentations.Open(SRC, ReadOnly=True, Untitled=False, WithWindow=False)
        for i in range(1, pres.Slides.Count + 1):
            png = os.path.join(OUT, f"t_{i}.png")
            t0 = time.perf_counter()
            pres.Slides(i).Export(png, "PNG", 1920, 1080)
            times.append(time.perf_counter() - t0)
    finally:
        if pres is not None:
            try:
                pres.Close()
            except Exception:
                pass
        if app is not None and not had:
            try:
                if app.Presentations.Count == 0:
                    app.Quit()
            except Exception:
                pass
        app = pres = None
        pythoncom.CoUninitialize()
    return times


if __name__ == "__main__":
    raise SystemExit(main())
