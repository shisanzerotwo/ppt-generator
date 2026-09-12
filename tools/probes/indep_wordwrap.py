"""独立验证 ④ 附录 · 真实稿的 `word_wrap` 分布（关系到"行级高亮"是否真被用上）。

线索：探针里用 `shapes.add_textbox()` 造的 1200 字段落，`hl_layout` 只出 1 行，
而 `qa.measure_text_lines` 算 53 行 —— 因为 python-pptx 的 `add_textbox()` 默认写
`wrap="none"`，于是 `ShapeInfo.word_wrap is False`，`_paragraph_lines` 走"不换行"分支。
若 `builder.py` 产出的稿同样全是 `wrap="none"`，则 M2 的"行级 tight rect"在真稿上
**退化为段级**（每段恒 1 行），契约 §4.3 的断行层等于没被真实素材考过。

一条命令：./.venv/Scripts/python.exe tools/probes/indep_wordwrap.py
"""

import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

import hl_layout  # noqa: E402
import pptx_io  # noqa: E402
import qa  # noqa: E402

SRC = os.path.join(ROOT, "output", "b_multislide.pptx")


def main() -> int:
    deck, _ = pptx_io.read_pages(SRC)
    stats = {}
    multi_para_shapes = 0
    totals = {"units": 0, "lines": 0, "units_multi_line": 0}
    for page in deck.pages:
        units = hl_layout.build_units(hl_layout.page_shapes(page, deck))
        for u in units:
            totals["units"] += 1
            totals["lines"] += len(u.lines)
            if len(u.lines) > 1:
                totals["units_multi_line"] += 1
    for page in deck.pages:
        for shape in page["shapes"]:
            if shape.kind != "text":
                continue
            stats[shape.word_wrap] = stats.get(shape.word_wrap, 0) + 1
            if len(shape.paragraphs) > 1:
                multi_para_shapes += 1

    print(f"素材：{os.path.basename(SRC)}")
    print(f"  文本形状 word_wrap 取值分布：{stats}")
    print(f"  多段文本框数量：{multi_para_shapes}")
    print(f"  讲解单元 {totals['units']} 个 → 行 rect 共 {totals['lines']} 条"
          f"（多行单元 {totals['units_multi_line']} 个）")

    # 反证：把同一个稿的 word_wrap 强置 True，看行数会不会暴涨
    forced_lines = 0
    forced_units = 0
    for page in deck.pages:
        shapes = page["shapes"]
        for s in shapes:
            if s.kind == "text":
                s.word_wrap = True
        for u in hl_layout.build_units(hl_layout.page_shapes(page, deck)):
            forced_units += 1
            forced_lines += len(u.lines)
    print(f"  把 word_wrap 强置 True 后 → 行 rect 共 {forced_lines} 条"
          f"（{forced_units} 个单元）")

    # qa 的口径（永远按换行算）作为对照
    qa_lines = 0
    for page in deck.pages:
        for shape in page["shapes"]:
            if shape.kind != "text":
                continue
            inner_w = (shape.width_emu - shape.margin_left_emu - shape.margin_right_emu) / qa.EMU_PER_PT
            for para in shape.paragraphs:
                if not para.text.strip() or inner_w <= 0:
                    continue
                size = next((r.size_pt for r in para.runs if r.size_pt is not None),
                            qa.DEFAULT_FONT_SIZE_PT)
                qa_lines += qa.measure_text_lines(para.text, float(size), inner_w)
    print(f"  qa.measure_text_lines 口径（无视 word_wrap）→ 共 {qa_lines} 行")

    print("\n== 结论 ==")
    if stats and set(stats) == {False}:
        print("  真稿里所有文本框都是 wrap=\"none\" → hl_layout 对每段只出 1 行 rect，")
        print("  行级 tight 退化为段级 tight：M2 的断行层在这份素材上**未被考查**。")
    elif False in stats:
        print("  部分文本框 wrap=\"none\"（其上退化为段级），部分会换行。")
    else:
        print("  真稿文本框都参与换行，行级 tight 确实生效。")
    print(f"  同一份稿：word_wrap 保持原值 = {totals['lines']} 条行 rect；"
          f"强置 True = {forced_lines} 条；qa 口径 = {qa_lines} 行")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
