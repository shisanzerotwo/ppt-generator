"""V1/V2/V4/V5/V6/V7/V9 几何探针：用 COM 底图核对契约里的几何假设。

思路：契约里这些点都是"算出来的矩形 vs PowerPoint 真正画出来的墨迹"是否重合。
所以先 COM 导出 fixture 底图，再用 Pillow 在指定区域量墨迹包围盒，跟预测值比。

跑法：.venv/Scripts/python.exe tools/probes/verify_geometry.py
"""

import os
import statistics
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

FIXTURE = os.path.abspath("output/fixtures/fixtures_cover.pptx")
OUT = os.path.abspath("output/spike/geom")
EMU_PER_PT = 12700


# ---------------------------------------------------------------- COM 导出

def export(pptx_path, out_dir, width=1920):
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    os.makedirs(out_dir, exist_ok=True)
    app = win32com.client.DispatchEx("PowerPoint.Application")
    pre_count = app.Presentations.Count
    paths, timings = [], []
    pres = None
    try:
        pres = app.Presentations.Open(os.path.abspath(pptx_path), ReadOnly=True,
                                      Untitled=False, WithWindow=False)
        w_emu, h_emu = int(pres.PageSetup.SlideWidth), int(pres.PageSetup.SlideHeight)
        height = round(width * h_emu / w_emu)
        for i in range(1, pres.Slides.Count + 1):
            png = os.path.abspath(os.path.join(out_dir, f"slide_{i}.png"))
            t0 = time.perf_counter()
            pres.Slides(i).Export(png, "PNG", width, height)
            timings.append(time.perf_counter() - t0)
            paths.append(png)
    finally:
        if pres is not None:
            try:
                pres.Close()
            except Exception:
                pass
        if pre_count == 0:
            try:
                app.Quit()
            except Exception:
                pass
        pres = app = None
        pythoncom.CoUninitialize()
    return paths, timings


# ---------------------------------------------------------------- 墨迹工具

def load(path):
    from PIL import Image
    return Image.open(path).convert("RGB")


def ink_bbox(img, box, bg=None, thresh=30):
    """box=(l,t,w,h) px；返回框内墨迹的 (l,t,r,b) 相对整图坐标，无墨迹返回 None。"""
    l, t, w, h = box
    l, t, w, h = int(l), int(t), int(max(w, 0)), int(max(h, 0))
    if w <= 0 or h <= 0:
        return None
    px = img.load()
    if bg is None:
        # 四边 1px 边框环中位色当底色
        ring = []
        for x in range(l, min(l + w, img.width)):
            ring.append(px[x, max(0, t)])
            ring.append(px[x, min(img.height - 1, t + h - 1)])
        for y in range(t, min(t + h, img.height)):
            ring.append(px[max(0, l), y])
            ring.append(px[min(img.width - 1, l + w - 1), y])
        bg = tuple(int(statistics.median(c[i] for c in ring)) for i in range(3))

    x0 = y0 = 10 ** 9
    x1 = y1 = -1
    for y in range(t, min(t + h, img.height)):
        for x in range(l, min(l + w, img.width)):
            p = px[x, y]
            if (abs(p[0] - bg[0]) > thresh or abs(p[1] - bg[1]) > thresh
                    or abs(p[2] - bg[2]) > thresh):
                x0 = min(x0, x); x1 = max(x1, x)
                y0 = min(y0, y); y1 = max(y1, y)
    if x1 < 0:
        return None
    return (x0, y0, x1 + 1, y1 + 1)


def coverage(img, box, bg=None, thresh=30):
    l, t, w, h = (int(v) for v in box)
    if w <= 0 or h <= 0:
        return 0.0
    px = img.load()
    if bg is None:
        ring = []
        for x in range(l, min(l + w, img.width)):
            ring.append(px[x, max(0, t)])
            ring.append(px[x, min(img.height - 1, t + h - 1)])
        for y in range(t, min(t + h, img.height)):
            ring.append(px[max(0, l), y])
            ring.append(px[min(img.width - 1, l + w - 1), y])
        bg = tuple(int(statistics.median(c[i] for c in ring)) for i in range(3))
    ink = 0
    total = 0
    for y in range(t, min(t + h, img.height)):
        for x in range(l, min(l + w, img.width)):
            total += 1
            p = px[x, y]
            if (abs(p[0] - bg[0]) > thresh or abs(p[1] - bg[1]) > thresh
                    or abs(p[2] - bg[2]) > thresh):
                ink += 1
    return ink / total if total else 0.0


# ---------------------------------------------------------------- 主流程

def main() -> int:
    if not os.path.isfile(FIXTURE):
        print(f"[FAIL] 先跑 tools/make_fixture_pptx.py 生成 {FIXTURE}")
        return 1

    print("== COM 导出 fixture 底图（顺带测 M1 耗时） ==")
    paths, timings = export(FIXTURE, OUT, width=1920)
    print(f"导出 {len(paths)} 页；每页耗时 {[f'{t:.3f}s' for t in timings]}；"
          f"中位 {statistics.median(timings):.3f}s/页；合计 {sum(timings):.2f}s")
    for p in paths:
        from PIL import Image
        with Image.open(p) as im:
            print(f"  {os.path.basename(p)}: {im.size[0]}x{im.size[1]}")

    scale = 1920 / 12191695  # px per EMU

    def e2p(emu):
        return emu * scale

    # ---------------- V1/V2：组合形状换算 ----------------
    print("\n== V1/V2 组合形状仿射换算（第 1 页） ==")
    from pptx import Presentation
    from pptx.oxml.ns import qn
    prs = Presentation(FIXTURE)
    slide1 = prs.slides[0]
    img1 = load(paths[0])
    for sh in slide1.shapes:
        if sh.shape_type != 6:  # MSO_SHAPE_TYPE.GROUP
            continue
        xfrm = sh._element.find(qn("p:grpSpPr")).find(qn("a:xfrm"))
        cox = int(xfrm.find(qn("a:chOff")).get("x"))
        coy = int(xfrm.find(qn("a:chOff")).get("y"))
        cw = int(xfrm.find(qn("a:chExt")).get("cx"))
        ch = int(xfrm.find(qn("a:chExt")).get("cy"))
        print(f"组 {sh.name}: left={sh.left} top={sh.top} w={sh.width} h={sh.height}")
        print(f"  chOff=({cox},{coy}) chExt=({cw},{ch})  → 缩放 gw/cw={sh.width / cw:.3f} "
              f"gh/ch={sh.height / ch:.3f}")
        for child in sh.shapes:
            abs_l = sh.left + (child.left - cox) * sh.width / cw
            abs_t = sh.top + (child.top - coy) * sh.height / ch
            abs_w = child.width * sh.width / cw
            abs_h = child.height * sh.height / ch
            mapped = (e2p(abs_l), e2p(abs_t), e2p(abs_w), e2p(abs_h))
            identity = (e2p(child.left), e2p(child.top), e2p(child.width), e2p(child.height))
            bb_m = ink_bbox(img1, mapped)
            bb_i = ink_bbox(img1, identity)
            print(f"  子 {child.name} 子坐标=({child.left},{child.top},{child.width},{child.height})")
            print(f"    换算后 px={tuple(round(v) for v in mapped)}  框内墨迹bbox={bb_m} "
                  f"coverage={coverage(img1, mapped):.3f}")
            print(f"    不换算 px={tuple(round(v) for v in identity)} 框内墨迹bbox={bb_i} "
                  f"coverage={coverage(img1, identity):.3f}")

    # 组合 xfrm 缺失时 python-pptx 返回什么（V2）
    print("\n== V2 组 xfrm 缺失时的行为 ==")
    import copy
    from pptx.shapes.group import GroupShape
    prs2 = Presentation(FIXTURE)
    s2 = prs2.slides[0]
    grp = next(s for s in s2.shapes if s.shape_type == 6)
    grp_el = grp._element
    grpSpPr = grp_el.find(qn("p:grpSpPr"))
    saved = copy.deepcopy(grpSpPr)
    grpSpPr.remove(grpSpPr.find(qn("a:xfrm")))
    try:
        print(f"  删掉 a:xfrm 后：left={grp.left} top={grp.top} width={grp.width} height={grp.height}")
    except Exception as exc:
        print(f"  删掉 a:xfrm 后取属性抛异常：{type(exc).__name__}: {exc}")
    grpSpPr.insert(0, saved.find(qn("a:xfrm")))

    # ---------------- V4/V5/V6/V7：排版页（第 5 页） ----------------
    print("\n== V4/V5/V6/V7 垂直锚点 / 首段段前距 / 行距 / fontScale（第 5 页） ==")
    img5 = load(paths[4])
    for name in ("TypographyBox", "AutofitBox"):
        sh = next((s for s in prs.slides[4].shapes if s.name == name), None)
        if sh is None:
            continue
        tf = sh.text_frame
        print(f"\n  ── {name} ──")
        print(f"  框 px=({e2p(sh.left):.0f},{e2p(sh.top):.0f},{e2p(sh.width):.0f},{e2p(sh.height):.0f})"
              f"  margin_t={tf.margin_top} (={tf.margin_top / EMU_PER_PT:.1f}pt)"
              f"  vertical_anchor={tf.vertical_anchor}  auto_size={tf.auto_size}")
        bodyPr = tf._txBody.find(qn("a:bodyPr"))
        norm = bodyPr.find(qn("a:normAutofit"))
        print(f"  normAutofit fontScale = {norm.get('fontScale') if norm is not None else None}")
        full = (e2p(sh.left), e2p(sh.top), e2p(sh.width), e2p(sh.height))
        bb = ink_bbox(img5, full)
        print(f"  整框墨迹 bbox={bb} → 首行墨迹距框顶 {(bb[1] - e2p(sh.top)):.1f}px "
              f"= {(bb[1] - e2p(sh.top)) / 2:.1f}pt（scale 2px/pt）")
        # 行分段：按墨迹行投影切分
        l, t, w, h = (int(v) for v in full)
        px = img5.load()
        rows = []

        def is_ink_row(y):
            for x in range(l, min(l + w, img5.width)):
                p = px[x, y]
                if p[0] < 160 and p[1] < 160:
                    return True
            return False

        y = t
        while y < min(t + h, img5.height):
            if is_ink_row(y):
                y0 = y
                while y < min(t + h, img5.height) and is_ink_row(y):
                    y += 1
                rows.append((y0, y))
            else:
                y += 1
        print(f"  墨迹行带（px）：{rows}")
        if len(rows) >= 2:
            pitches = [rows[i + 1][0] - rows[i][0] for i in range(len(rows) - 1)]
            print(f"  行间距 px={pitches} → pt={[round(p / 2, 1) for p in pitches]}")
        heights = [b - a for a, b in rows]
        print(f"  各墨迹带高度 px={heights} → pt={[round(hh / 2, 1) for hh in heights]}")

    # ---------------- V9：合并单元格（第 2 页） ----------------
    print("\n== V9 合并单元格几何（第 2 页） ==")
    img2 = load(paths[1])
    tbl_shape = next(s for s in prs.slides[1].shapes if getattr(s, "has_table", False) and s.has_table)
    tbl = tbl_shape.table
    print(f"  表格 px 原点=({e2p(tbl_shape.left):.0f},{e2p(tbl_shape.top):.0f})")
    print(f"  列宽 EMU={[c.width for c in tbl.columns]}  行高 EMU={[r.height for r in tbl.rows]}")
    for r in range(len(tbl.rows)):
        for c in range(len(tbl.columns)):
            cell = tbl.cell(r, c)
            print(f"   cell[{r}][{c}] text={cell.text!r:28} origin={cell.is_merge_origin} "
                  f"spanned={cell.is_spanned} span=({cell.span_width}x{cell.span_height})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
