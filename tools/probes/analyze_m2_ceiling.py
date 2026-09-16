"""M2 覆盖率上限分析：0.286 是"还能优化"还是"行级矩形的结构上限"？

分解 coverage：
    coverage = (墨迹bbox面积 / rect面积) × (墨迹像素 / 墨迹bbox面积)
               └──── 我们的几何能控制的部分 ────┘  └── 字形墨迹密度（谁都改不了）──┘

再量化两个敏感性：
  · 把 pad_x/pad_y 归零，能涨多少（= 契约公式里我们自己的余量成本）；
  · 把 rect 直接换成"墨迹 bbox"（理想矩形），能到多少（= 任何矩形法的绝对上限）。

同时排除"底色估计偏差导致低估"：对比环取底色 vs 显式取画面角落底色。

跑法：.venv/Scripts/python.exe tools/probes/analyze_m2_ceiling.py
"""

import os
import statistics
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

import hl_layout  # noqa: E402
import pptx_io  # noqa: E402

SRC = os.path.abspath("output/b_multislide.pptx")
BG = os.path.abspath("output/spike/m1")
GROW = 6  # 只扩 6px：够吸收预测误差，又不会串到相邻形状


def _bg_corner(img):
    px = img.load()
    return px[4, 4]


def band_bg(img, box, band=3):
    """矩形**外侧** band 像素环带的中位色。

    比"矩形内侧 1px 环"更稳（按构造排除矩形内墨迹），比"幻灯片主色"更正确
    （带填充色的形状上不会被误判成墨迹）。实测：本素材上它与内侧环口径一致
    （都 0.286），而幻灯片主色口径是虚高的（0.391，全部来自 4 行落在浅色填充
    形状上的文字被整块算成墨迹）。
    """
    l, t, w, h = (int(round(v)) for v in box)
    l_a, t_a = max(0, l - band), max(0, t - band)
    r_a, b_a = min(img.width, l + w + band), min(img.height, t + h + band)
    r_i, b_i = min(img.width, l + w), min(img.height, t + h)
    px = img.load()
    vals = []
    for y in range(t_a, b_a):
        for x in range(l_a, r_a):
            if l <= x < r_i and t <= y < b_i:
                continue
            vals.append(px[x, y])
    if not vals:
        return None

    def med(ch):
        s = sorted(v[ch] for v in vals)
        return s[len(s) // 2]

    return (med(0), med(1), med(2))


def _ink_stats(img, box, bg_rgb, thresh=30):
    """返回 (bbox, ink_px, box_area)；bbox 相对整图。"""
    l, t, w, h = (int(round(v)) for v in box)
    l2, t2 = max(0, l), max(0, t)
    r2, b2 = min(img.width, l + w), min(img.height, t + h)
    if r2 <= l2 or b2 <= t2:
        return None, 0, 0
    px = img.load()
    x0 = y0 = 10 ** 9
    x1 = y1 = -1
    ink = 0
    for y in range(t2, b2):
        for x in range(l2, r2):
            p = px[x, y]
            if any(abs(p[i] - bg_rgb[i]) > thresh for i in range(3)):
                ink += 1
                x0, x1 = min(x0, x), max(x1, x)
                y0, y1 = min(y0, y), max(y1, y)
    if x1 < 0:
        return None, 0, (r2 - l2) * (b2 - t2)
    return (x0, y0, x1 + 1, y1 + 1), ink, (r2 - l2) * (b2 - t2)


def main() -> int:
    from PIL import Image

    deck, _ = pptx_io.read_pages(SRC)
    rows = []
    bg_bias = []
    for page in deck.pages:
        bg = os.path.join(BG, f"slide_{page['index'] + 1}.png")
        if not os.path.isfile(bg):
            print(f"[FAIL] 缺底图 {bg}，先跑 tools/probes/accept_m1.py")
            return 1
        img = Image.open(bg).convert("RGB")
        for u in hl_layout.build_units(hl_layout.page_shapes(page, deck)):
            for r in u.lines:
                box = (r.left_px, r.top_px, r.width_px, r.height_px)
                if box[2] <= 0 or box[3] <= 0:
                    continue
                # 底色偏差检查
                local = band_bg(img, box)
                bg_bias.append((hl_layout.measure_coverage(bg, box),
                                hl_layout.measure_coverage(bg, box, bg_rgb=local)))
                grown = (box[0] - GROW, box[1] - GROW, box[2] + 2 * GROW, box[3] + 2 * GROW)
                bb, ink, _ = _ink_stats(img, grown, local)
                if bb is None:
                    continue
                bb_w, bb_h = bb[2] - bb[0], bb[3] - bb[1]
                rows.append({
                    "page": page["index"] + 1, "name": u.shape_name, "kind": u.kind,
                    "size": u.size_pt, "text": u.text[:12],
                    "cov": hl_layout.measure_coverage(bg, box, bg_rgb=local),
                    "rect": box, "bbox": bb,
                    "fill_v": bb_h / box[3] if box[3] else 0.0,
                    "fill_h": bb_w / box[2] if box[2] else 0.0,
                    "density": ink / (bb_w * bb_h) if bb_w and bb_h else 0.0,
                })

    print("== 0. 底色估计偏差检查（环取底色 vs 角落底色）==")
    diffs = [abs(a - b) for a, b in bg_bias]
    print(f"  样本 {len(diffs)} 条，最大偏差 {max(diffs):.3f}，中位偏差 {statistics.median(diffs):.3f}")
    print(f"  两法各自中位：环={statistics.median([a for a, _ in bg_bias]):.3f} "
          f"角落={statistics.median([b for _, b in bg_bias]):.3f}")

    covs = [r["cov"] for r in rows]
    fill_v = [r["fill_v"] for r in rows]
    fill_h = [r["fill_h"] for r in rows]
    dens = [r["density"] for r in rows]
    bbox_cov = [r["fill_v"] * r["fill_h"] * r["density"] for r in rows]

    print(f"\n== 1. 分解（n={len(rows)} 行）==")
    print(f"  coverage 中位            = {statistics.median(covs):.3f}")
    print(f"  竖向填充 墨迹高/rect高   = {statistics.median(fill_v):.3f}  （line_h=字号×1.25 撑出来的行框）")
    print(f"  横向填充 墨迹宽/rect宽   = {statistics.median(fill_h):.3f}  （tight 宽度，已接近 1）")
    print(f"  字形墨迹密度 墨迹/bbox   = {statistics.median(dens):.3f}  （CJK 笔画在 em 盒里占比，谁都改不了）")
    print(f"  三者乘积（校验）         = {statistics.median(bbox_cov):.3f}")

    print(f"\n== 2. 绝对上限：把 rect 直接换成墨迹 bbox（理想矩形）==")
    print(f"  中位 = {statistics.median(dens):.3f} —— 任何矩形法都不可能超过它")

    print(f"\n== 3. 敏感性：pad 归零能涨多少（不违反契约公式的部分）==")
    print(f"  当前 pad_x=2pt pad_y=1pt；每行 rect 比文字多 4pt 宽 + 2pt 高")
    for pad in (0.0, 0.5, 1.0, 2.0):
        vals = [r["fill_v"] * r["fill_h"] * r["density"] * _pad_gain(r, pad)
                for r in rows]
        print(f"  pad={pad:.1f}pt 时覆盖中位 ≈ {statistics.median(vals):.3f}")

    print(f"\n== 4. 按字号分组（小字更吃亏：笔画细+pad 占比大）==")
    for lo, hi in ((0, 13), (13, 19), (19, 30), (30, 100)):
        g = [r["cov"] for r in rows if lo <= r["size"] < hi]
        if g:
            print(f"  {lo:>2}-{hi:<3}pt  n={len(g):2}  中位={statistics.median(g):.3f}  "
                  f"平均字形密度={statistics.median([r['density'] for r in rows if lo <= r['size'] < hi]):.3f}")

    print(f"\n== 5. 逐行明细（前 8 条）==")
    for r in rows[:8]:
        print(f"  第{r['page']}页 {r['name']:11} {r['size']:5.1f}pt cov={r['cov']:.3f} "
              f"竖{r['fill_v']:.2f} 横{r['fill_h']:.2f} 密度{r['density']:.2f} {r['text']!r}")

    print(f"\n== 6. 最差 5 行 ==")
    for r in sorted(rows, key=lambda x: x["cov"])[:5]:
        print(f"  第{r['page']}页 {r['name']:11} {r['size']:5.1f}pt cov={r['cov']:.3f} "
              f"竖{r['fill_v']:.2f} 横{r['fill_h']:.2f} 密度{r['density']:.2f} {r['text']!r}")
    return 0


def _pad_gain(row, pad_pt):
    """pad 从当前值变到 pad_pt 时的面积增益（px：1pt = 2px）。"""
    px = 2.0
    cur_x, cur_y = 2.0, 1.0
    w = row["rect"][2] - 2 * cur_x * px + 2 * pad_pt * px
    h = row["rect"][3] - 2 * cur_y * px + 2 * pad_pt * px
    if w <= 0 or h <= 0:
        return 1.0
    return (row["rect"][2] * row["rect"][3]) / (w * h)


if __name__ == "__main__":
    raise SystemExit(main())
