"""M2 诊断：逐个 Unit 比对"预测行矩形"与"底图真实墨迹包围盒"，找覆盖率损失来源。

跑法：.venv/Scripts/python.exe tools/probes/diag_m2_offsets.py
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

import hl_layout  # noqa: E402
import pptx_io  # noqa: E402

SRC = os.path.abspath("output/b_multislide.pptx")
BG = os.path.abspath("output/spike/m1")


def ink_bbox(img, box, thresh=30):
    """框内相对底色的墨迹 bbox（绝对坐标）。底色取四边环中位色。"""
    l, t, w, h = (int(v) for v in box)
    l, t = max(0, l), max(0, t)
    r, b = min(img.width, l + w), min(img.height, t + h)
    if r <= l or b <= t:
        return None
    px = img.load()
    ring = []
    for x in range(l, r):
        ring.append(px[x, t]); ring.append(px[x, b - 1])
    for y in range(t, b):
        ring.append(px[l, y]); ring.append(px[r - 1, y])
    bg = tuple(sorted(v[i] for v in ring)[len(ring) // 2] for i in range(3))
    x0 = y0 = 10 ** 9
    x1 = y1 = -1
    for y in range(t, b):
        for x in range(l, r):
            p = px[x, y]
            if any(abs(p[i] - bg[i]) > thresh for i in range(3)):
                x0, x1 = min(x0, x), max(x1, x)
                y0, y1 = min(y0, y), max(y1, y)
    return None if x1 < 0 else (x0, y0, x1 + 1, y1 + 1)


def main() -> int:
    from PIL import Image

    deck, _ = pptx_io.read_pages(SRC)
    for page, i in zip(deck.pages, range(len(deck.pages))):
        bg = os.path.join(BG, f"slide_{i + 1}.png")
        if not os.path.isfile(bg):
            continue
        img = Image.open(bg).convert("RGB")
        units = hl_layout.build_units(hl_layout.page_shapes(page, deck))
        print(f"\n===== 第 {i + 1} 页 =====")
        for u in units:
            for k, rect in enumerate(u.lines):
                box = (rect.left_px, rect.top_px, rect.width_px, rect.height_px)
                grown = (box[0] - 60, box[1] - 40, box[2] + 120, box[3] + 80)
                bb = ink_bbox(img, grown)
                cov = hl_layout.measure_coverage(bg, box)
                if bb is None:
                    print(f"  [{u.kind}] {u.shape_name} 行{k} 无墨迹! rect={tuple(round(v) for v in box)}")
                    continue
                dx = bb[0] - box[0]
                dy = bb[1] - box[1]
                ink_w, ink_h = bb[2] - bb[0], bb[3] - bb[1]
                print(f"  [{u.kind:6}] {u.shape_name:12} 行{k} size={u.size_pt:5.1f} "
                      f"rect={tuple(round(v) for v in box)} "
                      f"ink=({ink_w}x{ink_h}) 偏移=({dx:+.0f},{dy:+.0f}) cov={cov:.3f} "
                      f"文本={u.text[:14]!r}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
