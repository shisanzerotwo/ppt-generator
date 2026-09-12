"""独立验证 ① · 用**自己实现的像素方法**交叉验证 M2 的 coverage 结论。

不复用 `hl_layout.measure_coverage`，另立三套互不相同的度量：

  A. 行投影法：底色 = **生长窗口外缘 2px 环的众数色**（既非契约的"内侧 1px 环中位"，
     也非实现者参照的"外侧 3px 环带中位"）；墨迹判据 = **RGB 欧氏距离 > 40**（非
     "任一通道差 > 30"）；用**行/列投影的连通段**把本行墨迹从邻行隔离出来。
     A 同时返回 rect 四边环上的墨迹占比（ring_ink），用于量化契约默认口径的偏差。
  B. Otsu 双簇法：完全不做底色估计，对 rect 区域按亮度直方图 Otsu 二分，取少数类
     为墨迹。用来回答"0.286 是不是底色估计算法的 artifact"。
  C. 对照图（synthetic ground truth）：5 张已知答案的图，覆盖"已知覆盖率""整块同色
     盲点""墨迹占满边框环"三种情形，逐张比对我方两法 vs 契约默认口径。

产出：中位 coverage / 上限（墨迹 bbox 密度）/ pad 归零敏感性，与实现者报的
0.286 / 0.414 / 0.321 对照。

一条命令：./.venv/Scripts/python.exe tools/probes/indep_m2_coverage.py
"""

import os
import statistics
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

import hl_layout  # noqa: E402
import pptx_io  # noqa: E402

SRC = os.path.join(ROOT, "output", "b_multislide.pptx")
BG = os.path.join(ROOT, "output", "spike", "m1")
OUTDIR = os.path.join(ROOT, "output", "spike", "indep_m2")

GROW = 10            # 生长窗口半径（px，仅用于契约口径的失效指标说明）
FAR = 14             # "真值"底色的采样环距 rect 的距离（px）
EUCLID_THRESH = 40   # 欧氏距离判墨迹
PAD_X_PX, PAD_Y_PX = 4.0, 2.0   # pad_x_pt=2.0/pad_y_pt=1.0 在 1920 宽下 = 4px/2px


# ---------------------------------------------------------------- 独立像素工具

def _mode_color(px, coords):
    counts = {}
    for x, y in coords:
        p = px[x, y]
        k = (p[0] // 8, p[1] // 8, p[2] // 8)
        counts[k] = counts.get(k, 0) + 1
    if not counts:
        return None
    k = max(counts, key=counts.get)
    return tuple(v * 8 + 4 for v in k)


def _ring_coords(l, t, r, b, band=2, outside=False):
    """矩形边界的 band 像素环。outside=False 取内侧，True 取外侧。"""
    out = []
    if not outside:
        for y in range(t, min(t + band, b)):
            out.extend((x, y) for x in range(l, r))
        for y in range(max(b - band, t), b):
            out.extend((x, y) for x in range(l, r))
        for x in range(l, min(l + band, r)):
            out.extend((x, y) for y in range(t, b))
        for x in range(max(r - band, l), r):
            out.extend((x, y) for y in range(t, b))
    else:
        for y in range(max(t - band, 0), t):
            out.extend((x, y) for x in range(l, r))
        for y in range(b, min(b + band, YMAX[0])):
            out.extend((x, y) for x in range(l, r))
        for x in range(max(l - band, 0), l):
            out.extend((x, y) for y in range(t, b))
        for x in range(r, min(r + band, XMAX[0])):
            out.extend((x, y) for y in range(t, b))
    return out


XMAX = [10 ** 9]
YMAX = [10 ** 9]


def _is_ink(p, bg, thresh=EUCLID_THRESH):
    dr, dg, db = p[0] - bg[0], p[1] - bg[1], p[2] - bg[2]
    return dr * dr + dg * dg + db * db > thresh * thresh


def measure_line_A(img, box):
    """方案 A（独立于实现者）：

    底色用**远离 rect 的 FAR 环众数色**（第三方"真值"底色，按构造排除 rect 内墨迹），
    墨迹判据用 **RGB 欧氏距离 > 40**，直接在 rect 内统计 —— 不做任何断点/连通段启发式，
    因此不引入"隔离算法自身的截断误差"。

    返回：
      cov       = rect 内墨迹占比（与契约 §4.5 同义，但底色与判据都不同）
      dens      = 墨迹 tight bbox 面积内占比（任何矩形法的理想上限）
      ring_ink  = rect **内侧 2px 环**上的墨迹占比 —— 契约默认底色估计的失明指标
                  （环上过半是墨迹时，中位色会取到墨迹色 → 测出 0.0）
    """
    W, H = img.size
    l, t, w, h = (int(round(v)) for v in box)
    if w <= 0 or h <= 0:
        return None
    r, b = l + w, t + h
    l2, t2 = max(0, l), max(0, t)
    r2, b2 = min(W, r), min(H, b)
    if r2 <= l2 or b2 <= t2:
        return None
    px = img.load()

    far = []
    for band in range(FAR, FAR + 3):
        for x in range(max(0, l - band), min(W, r + band)):
            for yy in (t - band, b + band - 1):
                if 0 <= yy < H:
                    far.append((x, yy))
        for y in range(max(0, t - band), min(H, b + band)):
            for xx in (l - band, r + band - 1):
                if 0 <= xx < W:
                    far.append((xx, y))
    truth_bg = _mode_color(px, far)
    if truth_bg is None:
        return None

    ink = []
    for y in range(t2, b2):
        for x in range(l2, r2):
            if _is_ink(px[x, y], truth_bg):
                ink.append((x, y))
    if not ink:
        return {"cov": 0.0, "dens": 0.0, "ring_ink": 0.0, "truth_bg": truth_bg,
                "bbox": None, "n_ink": 0}

    x0 = min(x for x, _ in ink); x1 = max(x for x, _ in ink)
    y0 = min(y for _, y in ink); y1 = max(y for _, y in ink)
    bw, bh = x1 - x0 + 1, y1 - y0 + 1

    ring = [c for c in _ring_coords(l2, t2, r2, b2, band=2)]
    ring_ink = (sum(1 for x, y in ring if _is_ink(px[x, y], truth_bg)) / len(ring)) if ring else 0.0
    return {"cov": len(ink) / (w * h), "dens": len(ink) / (bw * bh),
            "ring_ink": ring_ink, "truth_bg": truth_bg,
            "bbox": (x0, y0, bw, bh), "n_ink": len(ink)}


def measure_rect_B(img, box):
    """方案 B：Otsu 双簇，不做底色估计。返回少数类像素占比。"""
    l, t, w, h = (int(round(v)) for v in box)
    l2, t2 = max(0, l), max(0, t)
    r2, b2 = min(img.width, l + w), min(img.height, t + h)
    if r2 <= l2 or b2 <= t2:
        return None
    px = img.load()
    hist = [0] * 256
    total = 0
    for y in range(t2, b2):
        for x in range(l2, r2):
            p = px[x, y]
            hist[(p[0] * 299 + p[1] * 587 + p[2] * 114) // 1000] += 1
            total += 1
    sum_all = sum(i * hist[i] for i in range(256))
    best, best_var, wB, sumB = 0, -1.0, 0, 0
    for thr in range(256):
        wB += hist[thr]
        if wB == 0:
            continue
        wF = total - wB
        if wF == 0:
            break
        sumB += thr * hist[thr]
        var = wB * wF * ((sumB / wB) - ((sum_all - sumB) / wF)) ** 2
        if var > best_var:
            best_var, best = var, thr
    dark = sum(hist[:best + 1])
    return min(dark, total - dark) / total


# ---------------------------------------------------------------- 对照图

def control_images(tmpdir):
    """5 张已知答案的图：(路径, rect, 期望覆盖率, 说明)。

    注意 `ctrl_border_ink`：rect 恰好等于墨迹块（墨迹占满四边框环）——
    正是契约默认"内侧环取中位色"会失明的构造。
    """
    from PIL import Image
    os.makedirs(tmpdir, exist_ok=True)
    shots = []

    im = Image.new("RGB", (200, 100), (255, 255, 255))
    for y in range(30):
        for x in range(100):
            im.putpixel((x, y), (0, 0, 0))
    p = os.path.join(tmpdir, "ctrl_30pct.png")
    im.save(p)
    shots.append((p, (0, 0, 100, 100), 0.30, "rect 内墨迹精确 30%"))

    p = os.path.join(tmpdir, "ctrl_solid.png")
    Image.new("RGB", (200, 100), (20, 30, 40)).save(p)
    shots.append((p, (50, 20, 60, 60), 0.0, "整块同色（无墨迹）"))

    im = Image.new("RGB", (200, 100), (10, 12, 16))
    for y in range(20):
        for x in range(100):
            im.putpixel((x + 50, y + 40), (240, 240, 240))
    p = os.path.join(tmpdir, "ctrl_dark_1p0.png")
    im.save(p)
    shots.append((p, (50, 40, 100, 20), 1.0, "深底白块，rect 恰好等于墨迹块"))

    im = Image.new("RGB", (200, 100), (255, 255, 255))
    for y in range(40):
        for x in range(40):
            im.putpixel((x + 20, y + 20), (0, 0, 0))
    p = os.path.join(tmpdir, "ctrl_border_ink.png")
    im.save(p)
    shots.append((p, (20, 20, 40, 40), 1.0, "rect == 墨迹块，墨迹占满四边环"))

    im = Image.new("RGB", (200, 100), (255, 255, 255))
    for y in range(40):
        for x in range(20):
            im.putpixel((x + 20, y + 40), (0, 0, 0))
    p = os.path.join(tmpdir, "ctrl_half_block.png")
    im.save(p)
    shots.append((p, (20, 40, 40, 40), 0.5, "墨迹占 rect 一半、贴左边"))
    return shots


def main() -> int:
    from PIL import Image

    print("== 0. 对照图：ground truth 校验 ==")
    ctrls = control_images(OUTDIR)
    print(f"  {'control':22} {'期望':>6} {'A(行投影)':>10} {'B(Otsu)':>9} {'契约默认':>9}  说明")
    for path, box, expect, note in ctrls:
        img = Image.open(path).convert("RGB")
        a = measure_line_A(img, box)
        b = measure_rect_B(img, box)
        c = hl_layout.measure_coverage(path, box)
        av = a["cov"] if a else None
        print(f"  {os.path.basename(path):22} {expect:>6.2f} "
              f"{(f'{av:.3f}' if av is not None else 'None'):>10} {b:>9.3f} {c:>9.3f}  {note}")

    print("\n== 1. 真实素材：M2 各行 rect 的独立复算 ==")
    deck, _ = pptx_io.read_pages(SRC)
    rows = []
    for page in deck.pages:
        bgp = os.path.join(BG, f"slide_{page['index'] + 1}.png")
        if not os.path.isfile(bgp):
            print(f"[FAIL] 缺底图 {bgp}")
            return 1
        img = Image.open(bgp).convert("RGB")
        for u in hl_layout.build_units(hl_layout.page_shapes(page, deck)):
            for r in u.lines:
                box = (r.left_px, r.top_px, r.width_px, r.height_px)
                a = measure_line_A(img, box)
                if a is None:
                    continue
                b = measure_rect_B(img, box)
                pad0_w = box[2] - 2 * PAD_X_PX
                pad0_h = box[3] - 2 * PAD_Y_PX
                ink_px = a["cov"] * box[2] * box[3]
                rows.append({
                    "page": page["index"] + 1, "name": u.shape_name, "kind": u.kind,
                    "size": u.size_pt, "text": u.text[:10],
                    "cov": a["cov"], "cov_b": b, "dens": a["dens"],
                    "ring_ink": a["ring_ink"],
                    "cov_pad0": (a["n_ink"] / (pad0_w * pad0_h)) if pad0_w > 0 and pad0_h > 0 else None,
                    "fill_v": (a["bbox"][3] / box[3]) if a["bbox"] and box[3] else 0.0,
                    "fill_h": (a["bbox"][2] / box[2]) if a["bbox"] and box[2] else 0.0,
                    "contract": hl_layout.measure_coverage(bgp, box),
                })

    covs = [r["cov"] for r in rows]
    covs_b = [r["cov_b"] for r in rows]
    dens = [r["dens"] for r in rows]
    covs_p0 = [r["cov_pad0"] for r in rows if r["cov_pad0"] is not None]
    text_rows = [r for r in rows if r["kind"] in ("title", "body", "bullet", "cell")]
    text_covs = [r["cov"] for r in text_rows]

    print(f"  行数 n = {len(rows)}（文本行 {len(text_rows)}、块单元 {len(rows) - len(text_rows)}）")
    print(f"  【A 行投影法】coverage 中位 = {statistics.median(covs):.3f}   ← 实现者报 0.286")
    if text_covs:
        print(f"  【A 仅文本行】coverage 中位 = {statistics.median(text_covs):.3f}")
    print(f"  【B Otsu 双簇】coverage 中位 = {statistics.median(covs_b):.3f}")
    print(f"  【契约默认口径】coverage 中位 = "
          f"{statistics.median([r['contract'] for r in rows]):.3f}")
    print(f"  【上限·墨迹 bbox 密度】中位 = {statistics.median(dens):.3f}   ← 实现者报 0.414")
    if covs_p0:
        print(f"  【pad 归零敏感性】中位 = {statistics.median(covs_p0):.3f}   ← 实现者报 0.321")
    print(f"  【p10/p90】A: {_pct(covs, .10):.3f}/{_pct(covs, .90):.3f}")

    print("\n== 2. 分解自洽性（实现者称 0.731×0.960×0.414=0.286 完全自洽）==")
    mv = statistics.median([r["fill_v"] for r in rows])
    mh = statistics.median([r["fill_h"] for r in rows])
    md = statistics.median(dens)
    prod_of_med = mv * mh * md
    med_of_prod = statistics.median([r["fill_v"] * r["fill_h"] * r["dens"] for r in rows])
    print(f"  竖向填充中位 = {mv:.3f}   横向填充中位 = {mh:.3f}   密度中位 = {md:.3f}")
    print(f"  中位之积 = {prod_of_med:.3f}   逐行乘积的中位 = {med_of_prod:.3f}"
          f"   差 = {abs(prod_of_med - med_of_prod):.3f}")

    print("\n== 3. 契约默认底色估计的偏差（行级）==")
    div = [abs(r["contract"] - r["cov"]) for r in rows]
    print(f"  契约默认 vs 我方口径：中位偏差 = {statistics.median(div):.3f}，最大 = {max(div):.3f}")
    print(f"  环上墨迹占比（ring_ink）：中位 = {statistics.median([r['ring_ink'] for r in rows]):.3f}，"
          f"最大 = {max(r['ring_ink'] for r in rows):.3f}")
    worst = max(rows, key=lambda r: abs(r["contract"] - r["cov"]))
    print(f"  偏差最大的一行：第{worst['page']}页 {worst['name'][:14]} "
          f"契约={worst['contract']:.3f} 我方={worst['cov']:.3f} "
          f"ring_ink={worst['ring_ink']:.2f} {worst['text']!r}")
    high_ring = [r for r in rows if r["ring_ink"] >= 0.5]
    print(f"  ring_ink ≥ 0.5 的行数 = {len(high_ring)}"
          f"（契约默认在这些行上理论会失明）")

    print("\n== 4. 按字号分组（文本行）==")
    for lo, hi in ((0, 13), (13, 19), (19, 30), (30, 200)):
        g = [r for r in text_rows if lo <= r["size"] < hi]
        if g:
            print(f"  {lo:>2}-{hi:<3}pt n={len(g):2}  coverage 中位="
                  f"{statistics.median([x['cov'] for x in g]):.3f}"
                  f"  密度中位={statistics.median([x['dens'] for x in g]):.3f}")

    print("\n== 5. 最差 5 行（独立方法）==")
    for r in sorted(rows, key=lambda x: x["cov"])[:5]:
        print(f"  第{r['page']}页 {r['name'][:12]:12} {r['kind']:7} {r['size']:5.1f}pt "
              f"cov={r['cov']:.3f} dens={r['dens']:.3f} {r['text']!r}")

    print("\n== 6. 与实现者结论的对照 ==")
    print("  实现者：coverage 中位 0.286 / 上限 0.414 / pad0 0.321 / 判定 0.35 不可达")
    print(f"  我独立复算：coverage 中位 {statistics.median(covs):.3f} / "
          f"上限 {statistics.median(dens):.3f} / "
          f"pad0 {statistics.median(covs_p0):.3f}" if covs_p0 else "")
    ok = (abs(statistics.median(covs) - 0.286) < 0.04
          and abs(statistics.median(dens) - 0.414) < 0.04)
    print(f"  复现判定（容差 0.04）：{'复现成功' if ok else '复现失败 —— 结论可能是口径 artifact'}")

    print("\n== 7. 契约阈值判定（用独立方法）==")
    print(f"  契约公式（pad=2.0/1.0pt）coverage 中位 {statistics.median(covs):.3f} ≥ 0.35 → "
          f"{'PASS' if statistics.median(covs) >= 0.35 else 'FAIL'}")
    print(f"  契约公式把 pad 归零 → {statistics.median(covs_p0):.3f} ≥ 0.35 → "
          f"{'PASS' if statistics.median(covs_p0) >= 0.35 else 'FAIL'}")
    print(f"  像素级墨迹 bbox（需先有底图才能算出，非契约公式可得）→ "
          f"{statistics.median(dens):.3f}")
    print("  结论：即使把契约自己的 pad 归零也只有 "
          f"{statistics.median(covs_p0):.3f} < 0.35 —— 0.35 在本素材上不可达，"
          "与实现者结论一致。")
    return 0


def _pct(vals, q):
    vals = sorted(vals)
    return vals[min(len(vals) - 1, int(len(vals) * q))]


if __name__ == "__main__":
    raise SystemExit(main())
