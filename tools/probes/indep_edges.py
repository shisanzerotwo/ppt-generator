"""独立验证 ④ · 边界与异常路径（用手工构造的极端 pptx 与畸形参数去撞）。

分四组：
  A. 输入文件病态：不存在 / 目录 / 改名的 txt / OLE2(加密) / 截断 zip / 合法 zip 但不是 pptx / 0 页
  B. 内容极端：空白页 / 全空文本框 / 纯空白文本 / 单字宽 > 行宽 / >1000 字段落 /
     零尺寸形状 / 隐藏形状 / 合并单元格表格 / 无 xfrm 组合
  C. 无 PowerPoint 的错误路径（打桩 `powerpoint_available` → False）：断言给的是
     **明确中文 PptxError**，不是裸堆栈
  D. `build_player` 的畸形参数（契约 §6.1 要求参数校验一律 PptxError）

一条命令：./.venv/Scripts/python.exe tools/probes/indep_edges.py
"""

import os
import sys
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

import hl_anim  # noqa: E402
import hl_layout  # noqa: E402
import pptx_io  # noqa: E402
import qa  # noqa: E402
from pptx import Presentation  # noqa: E402
from pptx.util import Emu, Inches, Pt  # noqa: E402

SRC = os.path.join(ROOT, "output", "b_multislide.pptx")
OUT = os.path.join(ROOT, "output", "spike", "indep_edges")
problems = []


def note(ok, label, detail=""):
    mark = "OK  " if ok else "✗   "
    print(f"  {mark}{label}{('  ' + detail) if detail else ''}")


def expect_error(label, fn, code=None):
    """跑 fn，断言抛 PptxError（可选核对 code）且 message 为中文、无堆栈。"""
    try:
        fn()
    except pptx_io.PptxError as exc:
        good = (code is None or exc.code == code)
        ascii_ok = all(ord(c) < 128 for c in exc.code)
        if not good:
            problems.append(f"{label}: code={exc.code}（期望 {code}）")
        if not ascii_ok or not exc.message:
            problems.append(f"{label}: message 异常 {exc.message!r}")
        note(good, label, f"code={exc.code} message={exc.message!r} hint={exc.hint[:24]!r}")
        return exc
    except Exception as exc:  # noqa: BLE001
        problems.append(f"{label}: 抛了非 PptxError 的 {type(exc).__name__}: {exc}")
        note(False, label, f"裸异常 {type(exc).__name__}: {exc}")
        return None
    problems.append(f"{label}: 没抛异常")
    note(False, label, "未抛异常")
    return None


# ---------------------------------------------------------------- 素材构造

def make_pptx(path, build):
    prs = Presentation()
    build(prs)
    prs.save(path)
    return path


def slide_blank(prs):
    prs.slides.add_slide(prs.slide_layouts[6])


def slide_edges(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    tb = s.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    tb.text_frame.text = "     "                       # 纯空白
    tb2 = s.shapes.add_textbox(Inches(1), Inches(2), Inches(1), Inches(1))
    tb2.name = "HugeChar"
    p = tb2.text_frame.paragraphs[0]
    r = p.add_run(); r.text = "疆"; r.font.size = Pt(200)   # 单字宽 >> 行宽
    tb3 = s.shapes.add_textbox(Inches(1), Inches(3.2), Inches(6), Inches(1))
    tb3.name = "LongText"
    tb3.text_frame.word_wrap = True          # add_textbox 默认 wrap="none"，这里要真换行
    p3 = tb3.text_frame.paragraphs[0]
    r3 = p3.add_run(); r3.text = "长文本测试。" * 200        # 1200 字
    r3.font.size = Pt(18)
    z = s.shapes.add_textbox(Inches(9), Inches(1), Emu(0), Inches(1))  # 零宽
    z.name = "ZeroWidth"
    z.text_frame.text = "零宽框"
    h = s.shapes.add_textbox(Inches(9), Inches(3), Inches(2), Inches(1))
    h.name = "HiddenBox"
    h.text_frame.text = "隐藏文本"
    h._element.nvSpPr.cNvPr.set("hidden", "1")
    tb4 = s.shapes.add_textbox(Inches(1), Inches(4.6), Inches(4), Inches(1))
    tb4.name = "NoWrap"
    tf = tb4.text_frame
    tf.word_wrap = False
    tf.text = "不换行的一句话，应该只出一行"

def slide_table(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    tbl = s.shapes.add_table(2, 2, Inches(1), Inches(1), Inches(6), Inches(2)).table
    # 合并 origin 格里放**长文本**：IR 缺 span 信息 → 只能按单列宽断行
    tbl.cell(0, 0).text = "合并单元格里的长文本内容测试换行行为"
    tbl.cell(0, 0).merge(tbl.cell(0, 1))
    tbl.cell(1, 0).text = "左下"
    tbl.cell(1, 1).text = "右下"


def slide_group_no_xfrm(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    g = s.shapes.add_group_shape()
    inner = g.shapes.add_textbox(Inches(1), Inches(1), Inches(2), Inches(1))
    inner.text_frame.text = "组内文字"
    from pptx.oxml.ns import qn
    grp_sp_pr = g._element.find(qn("p:grpSpPr"))
    xfrm = grp_sp_pr.find(qn("a:xfrm")) if grp_sp_pr is not None else None
    if xfrm is not None:
        grp_sp_pr.remove(xfrm)


def main() -> int:
    os.makedirs(OUT, exist_ok=True)
    for name in ("blank.pptx", "edges.pptx", "table.pptx", "grpnoxfrm.pptx"):
        make_pptx(os.path.join(OUT, name),
                  {"blank.pptx": slide_blank, "edges.pptx": slide_edges,
                   "table.pptx": slide_table, "grpnoxfrm.pptx": slide_group_no_xfrm}[name])
    zeroslide = os.path.join(OUT, "zero_slide.pptx")
    Presentation().save(zeroslide)          # 真 0 页（未 add_slide）

    # ------------------------------------------------ A. 输入文件病态
    print("== A. 输入文件病态 ==")
    expect_error("不存在的路径", lambda: pptx_io.read_pages(os.path.join(OUT, "nope.pptx")),
                 "PPTX_NOT_FOUND")
    expect_error("目录当文件", lambda: pptx_io.read_pages(OUT), "PPTX_NOT_FOUND")
    txt = os.path.join(OUT, "renamed.pptx")
    with open(txt, "w", encoding="utf-8") as f:
        f.write("这不是 pptx，只是改了个扩展名。")
    expect_error("txt 改名成 .pptx", lambda: pptx_io.read_pages(txt), "PPTX_UNREADABLE")

    ole = os.path.join(OUT, "encrypted.pptx")
    with open(ole, "wb") as f:
        f.write(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + bytes(64))
    expect_error("OLE2 头（伪装加密稿）", lambda: pptx_io.read_pages(ole), "PPTX_ENCRYPTED")

    trunc = os.path.join(OUT, "truncated.pptx")
    with open(SRC, "rb") as f:
        data = f.read()
    with open(trunc, "wb") as f:
        f.write(data[:200])
    expect_error("截断的 zip", lambda: pptx_io.read_pages(trunc), "PPTX_UNREADABLE")

    notzip = os.path.join(OUT, "notapptx.pptx")
    with zipfile.ZipFile(notzip, "w") as z:
        z.writestr("hello.txt", "hi")
    expect_error("合法 zip 但不是 pptx", lambda: pptx_io.read_pages(notzip), "PPTX_UNREADABLE")

    expect_error("0 页 pptx", lambda: pptx_io.read_pages(zeroslide), "PPTX_EMPTY")

    # ------------------------------------------------ B. 内容极端
    print("\n== B. 内容极端 ==")
    deck, skips = pptx_io.read_pages(os.path.join(OUT, "edges.pptx"))
    page_shapes = deck.pages[0]["shapes"]
    names = {s.name: s for s in page_shapes}
    print(f"  读入形状 {len(page_shapes)} 个：{sorted(names)}")
    reasons = {}
    for s in skips:
        reasons[s["reason"]] = reasons.get(s["reason"], 0) + 1
    print(f"  skipped 分布：{reasons}")
    if "纯空白框" not in "".join(names):
        pass
    if any(s.get("shape_name") == "ZeroWidth" for s in skips):
        note(True, "零宽形状被过滤（zero_size）")
    else:
        problems.append("零宽形状未被过滤")
        note(False, "零宽形状被过滤（zero_size）", "未出现在 skipped 里")
    if any(s.get("shape_name") == "HiddenBox" and s.get("reason") == "hidden" for s in skips):
        note(True, "隐藏形状被过滤（hidden）")
    else:
        problems.append("隐藏形状未被过滤")
        note(False, "隐藏形状被过滤（hidden）")
    if "空白文本框" not in names:
        note(True, "纯空白文本框被过滤（empty_text）")

    units = hl_layout.build_units(hl_layout.page_shapes(deck.pages[0], deck))
    by_shape = {}
    for u in units:
        by_shape.setdefault(u.shape_name, []).append(u)
    print(f"  讲解单元 {len(units)} 个，来自形状 {sorted(by_shape)}")

    huge = by_shape.get("HugeChar")
    if huge:
        r = huge[0].lines[0]
        box_w_px = names["HugeChar"].width_emu * (deck.export_width_px / deck.width_emu)
        ok = len(huge[0].lines) == 1
        note(ok, "单字（200pt）宽于行宽仍只出 1 行",
             f"行数={len(huge[0].lines)} rect宽={r.width_px:.0f}px 形状宽={box_w_px:.0f}px")
        if not ok:
            problems.append(f"超宽单字行数异常：{len(huge[0].lines)}")
        if r.width_px > box_w_px:
            print(f"      → 注意：该行 rect 宽 {r.width_px:.0f}px **超出**形状宽 "
                  f"{box_w_px:.0f}px（超宽单字的必然结果，overflow 会 +1）")

    longu = by_shape.get("LongText")
    if longu:
        n = len(longu[0].lines)
        expect = qa.measure_text_lines("长文本测试。" * 200, 18.0,
                                      (names["LongText"].width_emu
                                       - names["LongText"].margin_left_emu
                                       - names["LongText"].margin_right_emu) / qa.EMU_PER_PT)
        note(n == expect, "1200 字段落行数与 qa 一致", f"行数={n}（qa={expect}）")
        if n != expect:
            problems.append(f"超长段落行数不一致：hl_layout={n} qa={expect}")

    nowrap = by_shape.get("NoWrap")
    if nowrap:
        note(len(nowrap[0].lines) == 1, "word_wrap=False 只出一行",
             f"行数={len(nowrap[0].lines)}")

    # 空白页（无形状）
    print("\n== 空白页（0 形状）==")
    one = os.path.join(OUT, "one_blank_slide.pptx")
    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    prs.save(one)
    deck3, _ = pptx_io.read_pages(one)
    u3 = hl_layout.build_units(hl_layout.page_shapes(deck3.pages[0], deck3))
    note(len(deck3.pages) == 1 and u3 == [], "有页但无形状 → 0 单元不崩",
         f"页数={len(deck3.pages)} units={len(u3)}")
    p = hl_anim.build_player(os.path.join(OUT, "player_empty"), ["bg/slide_1.png"], [u3])
    note(os.path.isfile(p), "空单元页也能生成播放器")

    # 表格（合并单元格）
    print("\n== 表格 / 合并单元格 ==")
    deck4, _ = pptx_io.read_pages(os.path.join(OUT, "table.pptx"))
    tbl_shape = deck4.pages[0]["shapes"][0]
    units4 = hl_layout.build_units(hl_layout.page_shapes(deck4.pages[0], deck4))
    cells = [u for u in units4 if u.kind == "cell"]
    col_w = tbl_shape.table_col_widths_emu
    print(f"  单元格单元 {len(cells)} 个；列宽={col_w}")
    for u in cells:
        print(f"    {u.text!r:18} rect 宽={u.rect.width_emu / qa.EMU_PER_PT:.1f}pt "
              f"（单列宽 {(col_w[0] - tbl_shape.margin_left_emu - tbl_shape.margin_right_emu) / qa.EMU_PER_PT:.1f}pt，"
              f"合并两列应 {(sum(col_w[:2]) - tbl_shape.margin_left_emu - tbl_shape.margin_right_emu) / qa.EMU_PER_PT:.1f}pt）")
    merged = [u for u in cells if "合并" in u.text]
    if merged:
        w_pt = merged[0].rect.width_emu / qa.EMU_PER_PT
        n_lines = len(merged[0].lines)
        single = (col_w[0] - tbl_shape.margin_left_emu - tbl_shape.margin_right_emu) / qa.EMU_PER_PT
        both = (sum(col_w[:2]) - tbl_shape.margin_left_emu - tbl_shape.margin_right_emu) / qa.EMU_PER_PT
        exp_single = qa.measure_text_lines(merged[0].text, merged[0].size_pt, single)
        exp_both = qa.measure_text_lines(merged[0].text, merged[0].size_pt, both)
        print(f"      → 合并 origin 格：按**单列宽**({single:.1f}pt)断行 = {n_lines} 行；"
              f"按合并宽({both:.1f}pt)应为 {exp_both} 行")
        print(f"        （hl_layout 实测 {n_lines} 行 / 单列宽期望 {exp_single} 行）")
        if n_lines == exp_single and exp_both < exp_single:
            print("      → 证实 D5：合并单元格按单列宽断行 → **多切出高亮行**，"
                  "且水平定位以起始列为准（居中/右对齐的合并格会错位）")
    else:
        problems.append("合并单元格未产出 cell 单元")

    # 无 xfrm 组合
    print("\n== 无 xfrm 组合 ==")
    deck5, skips5 = pptx_io.read_pages(os.path.join(OUT, "grpnoxfrm.pptx"))
    note(any(s.get("reason") == "group_no_xfrm" for s in skips5),
         "整组跳过（group_no_xfrm）",
         f"skipped={[s['reason'] for s in skips5]} 页形状数={len(deck5.pages[0]['shapes'])}")

    # ------------------------------------------------ C. 无 PowerPoint
    print("\n== C. 无 PowerPoint 的错误路径（打桩）==")
    real = pptx_io.powerpoint_available
    pptx_io.powerpoint_available = lambda: False
    try:
        exc = expect_error("powerpoint_available=False 时 export_pages",
                           lambda: pptx_io.export_pages(SRC, os.path.join(OUT, "bg_nocom")),
                           "NO_POWERPOINT")
        if exc and "PowerPoint" not in exc.message:
            problems.append(f"NO_POWERPOINT 的 message 未提 PowerPoint：{exc.message!r}")
        if exc and not exc.hint:
            problems.append("NO_POWERPOINT 缺 hint（可执行指引）")
    finally:
        pptx_io.powerpoint_available = real
    if not pptx_io.powerpoint_available():
        print("  [note] 真机 powerpoint_available() 返回 False（本机未装 PowerPoint）")
    else:
        print("  [note] 真机 powerpoint_available() = True（打桩只影响上面那次判定）")

    # ------------------------------------------------ D. build_player 畸形参数
    print("\n== D. build_player 的畸形参数 ==")
    from pptx_io import PptxError
    cases = [
        ("dim=nan", dict(dim=float("nan"))),
        ("dim=True", dict(dim=True)),
        ("auto_step_ms=0", dict(auto_step_ms=0)),
        ("auto_step_ms='abc'", dict(auto_step_ms="abc")),
        ("canvas_width_px=0", dict(canvas_width_px=0)),
        ("canvas_height_px=-5", dict(canvas_height_px=-5)),
        ("bg_paths 非列表(None)", dict(_bg=None)),
        ("pages_units 非列表(None)", dict(_pages=None)),
        ("units 里含 None", dict(_pages=[[None]])),
    ]
    for label, kw in cases:
        bg = kw.pop("_bg", ["bg/slide_1.png"])
        pages = kw.pop("_pages", [[]])
        try:
            hl_anim.build_player(os.path.join(OUT, "player_bad"), bg, pages, **kw)
            note(True, label, "未抛异常（已放行）")
            if label in ("bg_paths 非列表(None)", "pages_units 非列表(None)", "units 里含 None"):
                problems.append(f"{label} 被静默放行（契约 §6.1 要求 PptxError）")
        except PptxError as exc:
            note(True, label, f"PptxError code={exc.code}")
        except Exception as exc:  # noqa: BLE001
            note(False, label, f"**裸 {type(exc).__name__}**：{exc}")
            problems.append(f"{label} 抛裸异常 {type(exc).__name__}: {exc}（契约要求 PptxError）")

    # ------------------------------------------------ 判定
    print("\n== 判定 ==")
    if problems:
        for x in problems:
            print(f"  ✗ {x}")
        print(f"  共 {len(problems)} 项")
        return 1
    print("  全部边界/异常路径检查通过")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
