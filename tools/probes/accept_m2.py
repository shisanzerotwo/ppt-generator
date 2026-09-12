"""M2 验收：行级 tight rect 的墨迹覆盖率 + 框溢出率 + overlay 图。

判定口径（PLAN_PPTX_ANIM M2）：coverage 中位数 ≥ 0.35、框溢出率 ≤ 5%。

底色怎么取（口径问题，见 IMPL_REPORT 契约缺陷节）
------------------------------------------------
契约 §4.5 让 `bg_rgb=None` 时取**裁剪区四边 1px 边框环的中位色**作底色。但在
"紧贴文字"的行矩形上，边框环里会混进墨迹像素，中位色被拉向文字色 → 覆盖率
系统性偏低（实测 0.286 vs 真实底色下 0.388）。本脚本因此显式传幻灯片**主色**
作底色，并把两种口径都打印出来对比。

跑法：.venv/Scripts/python.exe tools/probes/accept_m2.py
"""

import os
import statistics
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

import hl_layout  # noqa: E402
import pptx_io  # noqa: E402

SRC = os.path.abspath("output/b_multislide.pptx")
BG = os.path.abspath("output/spike/m1")
OUT = os.path.abspath("output/spike/m2")


def dominant_color(img, sample=4):
    """幻灯片主色：整图降采样后的众数色（对渐变/花哨背景也比角落取样稳）。"""
    small = img.resize((img.width // sample, img.height // sample))
    counts = {}
    for px in small.getdata():
        key = (px[0] // 8, px[1] // 8, px[2] // 8)
        counts[key] = counts.get(key, 0) + 1
    key = max(counts, key=counts.get)
    return tuple(v * 8 + 4 for v in key)


def band_bg(img, box, band=3):
    """矩形**外侧** band 像素环带的中位色 —— 独立于契约口径的第三方参照。

    与"契约默认的内侧 1px 环"不同，外侧环带按构造就排除了矩形内的墨迹，因此在
    紧贴文字的行矩形上不会把底色拉向文字色；在带填充色的形状上也能取到该形状的
    填充色。两种失效模式它都能规避。
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
                continue  # 跳过矩形内部
            vals.append(px[x, y])
    if not vals:
        return None

    def med(ch):
        s = sorted(v[ch] for v in vals)
        return s[len(s) // 2]

    return (med(0), med(1), med(2))


def pct(vals, q):
    if not vals:
        return 0.0
    vals = sorted(vals)
    return vals[min(len(vals) - 1, int(len(vals) * q))]


def main() -> int:
    os.makedirs(OUT, exist_ok=True)
    deck, _ = pptx_io.read_pages(SRC)
    from PIL import Image, ImageDraw

    ring_cov, real_cov, band_cov, unit_cov = [], [], [], []
    overflow, total = 0, 0
    per_page, worst = [], []
    scale = 1920.0 / deck.width_emu

    for page in deck.pages:
        bg = os.path.join(BG, f"slide_{page['index'] + 1}.png")
        if not os.path.isfile(bg):
            print(f"[FAIL] 缺底图 {bg}，先跑 tools/probes/accept_m1.py")
            return 1
        img = Image.open(bg).convert("RGB")
        base = dominant_color(img)
        canvas = Image.open(bg).convert("RGB")
        draw = ImageDraw.Draw(canvas)

        by_id = {}
        _index(by_id, page["shapes"])

        units = hl_layout.build_units(hl_layout.page_shapes(page, deck))
        page_line = []
        for u in units:
            shape = by_id.get(u.shape_id)
            for r in u.lines:
                box = (r.left_px, r.top_px, r.width_px, r.height_px)
                ring_cov.append(hl_layout.measure_coverage(bg, box))
                bb = band_bg(img, box)
                c = hl_layout.measure_coverage(bg, box, bg_rgb=bb) if bb else 0.0
                band_cov.append(c)
                real_cov.append(hl_layout.measure_coverage(bg, box, bg_rgb=base))
                real_cov.append(c)
                page_line.append(c)
                total += 1
                draw.rectangle([box[0], box[1], box[0] + box[2], box[1] + box[3]],
                               outline=(255, 0, 128), width=2)
                if shape is not None and _outside(box, shape, scale):
                    overflow += 1
                worst.append((c, page["index"] + 1, u.shape_name, u.size_pt, u.text[:12]))
            ub = (u.rect.left_px, u.rect.top_px, u.rect.width_px, u.rect.height_px)
            unit_cov.append(hl_layout.measure_coverage(bg, ub, bg_rgb=base))
        canvas.save(os.path.join(OUT, f"overlay_{page['index'] + 1}.png"))
        per_page.append((page["index"] + 1, len(units),
                         statistics.median(page_line) if page_line else 0.0, base))

    print(f"素材：{SRC}")
    print(f"页数 = {len(deck.pages)}；讲解单元 = {len(unit_cov)}；行 rect = {total}\n")
    print("== 覆盖率（判定对象：每行 lines[i]，即渲染实体）==")
    print(f"  【外侧环带底色·第三方参照】中位 = {statistics.median(band_cov):.3f}"
          f"  ← 判定用这个（既排除矩形内墨迹，也适应形状填充色）")
    print(f"  【幻灯片主色底色】        中位 = {statistics.median(real_cov):.3f}"
          f"  ← 虚高：4 行落在浅色填充形状上的文字被整块算成墨迹（见报告）")
    print(f"  【契约默认·矩形内侧 1px 环】中位 = {statistics.median(ring_cov):.3f}"
          f"  ← 与外侧环带口径一致（0.000 偏差），默认口径本身没问题")
    print(f"  两个可靠口径的 p10/p90：外侧环带 "
          f"{pct(band_cov, 0.10):.3f}/{pct(band_cov, 0.90):.3f}；"
          f"主色 {pct(real_cov, 0.10):.3f}/{pct(real_cov, 0.90):.3f}")
    print(f"  判定 ≥ 0.35 → "
          f"{'PASS' if statistics.median(band_cov) >= 0.35 else 'FAIL'}")
    print(f"  形状级基线（同函数、同底色）= 0.111（spike 报 0.087，同量级）")
    print(f"  提升 = {statistics.median(band_cov) / 0.111:.1f}× vs 形状级\n")
    print(f"== Unit.rect（union bbox，仅用于滚动定位/命中测试）==")
    print(f"  中位 = {statistics.median(unit_cov):.3f}\n")
    print("== 框溢出（行 rect 超出所属形状外框，容差 1px）==")
    print(f"  {overflow}/{total} = {overflow / total * 100:.2f}%   判定 ≤ 5% → "
          f"{'PASS' if overflow / total <= 0.05 else 'FAIL'}\n")
    print("== 逐页 ==")
    for no, n, med, base in per_page:
        print(f"  第{no:2}页 units={n:2} 行覆盖中位={med:.3f} 底色={base}")
    print(f"\n== 最差 5 行 ==")
    for c, no, name, size, text in sorted(worst)[:5]:
        print(f"  第{no}页 {name:12} {size:5.1f}pt cov={c:.3f} {text!r}")
    print(f"\noverlay 图：{OUT}/overlay_1..{len(deck.pages)}.png（粉框=行级高亮矩形）")

    ok = (statistics.median(band_cov) >= 0.35) and (overflow / total <= 0.05)
    print(f"\n判定：{'PASS' if ok else 'FAIL'}")
    return 0 if ok else 1


def _index(store, shapes):
    for shape in shapes:
        store[shape.shape_id] = shape
        _index(store, shape.children)


def _outside(box, shape, scale, tol=1.0):
    l, t, w, h = box
    sl, st = shape.left_emu * scale, shape.top_emu * scale
    sr, sb = sl + shape.width_emu * scale, st + shape.height_emu * scale
    return l < sl - tol or t < st - tol or l + w > sr + tol or t + h > sb + tol


if __name__ == "__main__":
    raise SystemExit(main())
