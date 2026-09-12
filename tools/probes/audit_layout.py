"""审计复现 · hl_layout 静默失败 / 表格截断 / a:br 与 QA 分叉 / measure_coverage 盲点。

用法：
    PYTHONIOENCODING=utf-8 ./.venv/Scripts/python.exe tools/probes/audit_layout.py

全部为内存内构造，不读不写任何文件（除非显式传入真实底图）。
"""
import io
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

import qa  # noqa: E402
import hl_layout  # noqa: E402
from pptx_io import PageShapes, ParaInfo, RunInfo, ShapeInfo  # noqa: E402

W = 12191695
H = 6858000


def section(t):
    print("\n" + "=" * 78)
    print(t)
    print("=" * 78)


def para(text, size=18.0, **kw):
    runs = [RunInfo(text=text, size_pt=size, bold=None, italic=None, font_name=None)]
    return ParaInfo(text=text, runs=runs, align=kw.get("align"),
                    level=0, space_before_pt=kw.get("sb"), space_after_pt=kw.get("sa"),
                    line_spacing=kw.get("ls"))


def text_shape(text, size=18.0, left=1000000, top=1000000, w=4000000, h=2000000,
               ml_emu=91440, mr_emu=91440, mt_emu=45720, mb_emu=45720, **kw):
    return ShapeInfo(shape_id=2, name=kw.get("name", "TextBox 1"), kind="text",
                     left_emu=left, top_emu=top, width_emu=w, height_emu=h,
                     paragraphs=[para(text, size)], margin_left_emu=ml_emu,
                     margin_right_emu=mr_emu, margin_top_emu=mt_emu,
                     margin_bottom_emu=mb_emu, word_wrap=kw.get("word_wrap", True),
                     vertical_anchor=kw.get("va"))


def page(shapes):
    return PageShapes(index=0, width_emu=W, height_emu=H, shapes=shapes)


def main():
    section("1 · inner_w_pt <= 0：契约 §4.3 要求『记 warning』，实现是否记了？")
    s = text_shape("这是一段会被静默丢弃的文字", w=100000, ml_emu=91440, mr_emu=91440)
    inner_w_pt = (100000 - 91440 * 2) / 12700
    print(f"形状宽 100000 EMU，左右内边距各 91440（合计 182880）→ inner_w_pt = {inner_w_pt:.3f}pt（<=0）")
    units = hl_layout.build_units(page([s]), export_width_px=1920)
    print(f"build_units 返回单元数 = {len(units)}（预期 0 → 形状被跳过）")
    print(f"→ 有没有任何出口告诉调用方『这个形状为什么没有高亮』："
          f"{'有' if len(units) else '无（契约 §4.3 要求的 warning 无处可寻）'}")
    print("   build_units 返回 list[Unit]，没有 skipped 出口；"
          "契约 §4.3 的『记 warning』在本数据模型下不可达（Unit 不存在则 warnings 也传不出去）")

    section("2 · 表格：列宽/行高数组短于单元格 → 静默截断（不报错）")
    tc = [[[para("A1")], [para("B1")], [para("C1")]],
          [[para("A2")], [para("B2")], [para("C2")]]]
    t_short = ShapeInfo(shape_id=3, name="Table 1", kind="table",
                        left_emu=1000000, top_emu=1000000, width_emu=4000000, height_emu=2000000,
                        table_cells=tc,
                        table_col_widths_emu=[1000000],     # 只有 1 列，单元格却有 3 列
                        table_row_heights_emu=[500000, 500000])
    t_ok = ShapeInfo(shape_id=4, name="Table 2", kind="table",
                     left_emu=1000000, top_emu=1000000, width_emu=4000000, height_emu=2000000,
                     table_cells=tc,
                     table_col_widths_emu=[1000000, 1000000, 1000000],
                     table_row_heights_emu=[500000, 500000])
    u_short = hl_layout.build_units(page([t_short]), 1920)
    u_ok = hl_layout.build_units(page([t_ok]), 1920)
    print(f"列宽数组完整（3 列）→ 单元数 = {len(u_ok)}，文本 = {[u.text for u in u_ok]}")
    print(f"列宽数组截短（1 列）→ 单元数 = {len(u_short)}，文本 = {[u.text for u in u_short]}")
    print(f"→ 静默少出 {len(u_ok) - len(u_short)} 个单元，无异常、无 warning")
    # 行高数组短
    t_rows = ShapeInfo(shape_id=5, name="Table 3", kind="table",
                       left_emu=1000000, top_emu=1000000, width_emu=4000000, height_emu=2000000,
                       table_cells=tc,
                       table_col_widths_emu=[1000000, 1000000, 1000000],
                       table_row_heights_emu=[500000])      # 只有 1 行，单元格却有 2 行
    u_rows = hl_layout.build_units(page([t_rows]), 1920)
    print(f"行高数组截短（1 行）→ 单元数 = {len(u_rows)}，文本 = {[u.text for u in u_rows]}")

    section("3 · a:br（段内手动换行）：几何用的 _paragraph_lines 与溢出 QA 分叉")
    for t in ["第一行\n第二行", "A\nB", "前缀" + "字" * 20 + "\n后缀"]:
        for size, bw in ((18.0, 300.0), (18.0, 150.0)):
            n_layout = len(hl_layout._paragraph_lines(t, size, bw, True))
            n_wrap = len(hl_layout.wrap_lines(t, size, bw))
            n_qa = qa.measure_text_lines(t, size, bw)
            print(f"  text={t!r:40s} size={size} 行宽={bw}pt → "
                  f"_paragraph_lines={n_layout} wrap_lines={n_wrap} qa={n_qa}"
                  f"（与 QA 差 {n_layout - n_qa:+d}）")

    section("4 · 契约强制的等价断言只覆盖 wrap_lines，不覆盖 _paragraph_lines")
    print("  契约 §4.2：assert len(wrap_lines(t, s, w)) == qa.measure_text_lines(t, s, w)")
    print("  契约 §4.3：word_wrap 为 True 时应直接 wrap_lines(整段文本)")
    print("  实现 _paragraph_lines 多出一条『先按 \\n 硬拆』的分支 → 与 qa 模型漂移，"
          "而 build_units 走的正是这条分支")

    section("5 · measure_coverage：矩形完全落在底图外 / 退化矩形")
    from PIL import Image

    def png_bytes(img):
        b = io.BytesIO()
        img.save(b, "PNG")
        b.seek(0)
        return b

    white = png_bytes(Image.new("RGB", (64, 64), (255, 255, 255)))
    print(f"  空宽矩形 -> {hl_layout.measure_coverage(white, (10, 10, 0, 10))}")
    print(f"  负宽矩形 -> {hl_layout.measure_coverage(white, (10, 10, -5, 10))}")
    print(f"  完全在图外 (1000,1000,10,10) -> {hl_layout.measure_coverage(white, (1000, 1000, 10, 10))}")
    print("  ↑ 三种都返回 0.0，与『rect 里真的没有墨迹』不可区分（验收脚本据此报低分）")

    img = Image.new("RGB", (64, 64), (255, 255, 255))
    for x in range(16, 48):
        for y in range(16, 48):
            img.putpixel((x, y), (0, 0, 0))
    black_block = png_bytes(img)
    print(f"  纯黑块 (16,16,32,32)（环即墨迹）-> {hl_layout.measure_coverage(black_block, (16, 16, 32, 32))}"
          f"  ← 环取底色取到墨迹色 → 判 0（已知盲点，用例已锁）")
    print(f"  同一矩形显式传底色 -> "
          f"{hl_layout.measure_coverage(black_block, (16, 16, 32, 32), bg_rgb=(255, 255, 255))}")

    section("6 · 被改过的 deck.json（几何为 null / 画布为 0）→ 未捕获崩溃，而非 IR_MISMATCH")
    from pptx_io import _shape_from
    tampered = _shape_from({
        "shape_id": 1, "name": "Picture 1", "kind": "picture",
        "left_emu": None, "top_emu": None, "width_emu": None, "height_emu": None,
    })
    try:
        hl_layout.build_units(page([tampered]), 1920)
        print("  picture(left=None) -> 未抛异常（意外）")
    except Exception as exc:                                   # noqa: BLE001
        print(f"  picture(left=None) -> {type(exc).__name__}: {exc}")
        print(f"     可由 load_deck（契约称 deck.json 是『唯一真源』）读入，"
              f"但不属 PptxError → CLI 只能归 INTERNAL(5)，而非 IR_MISMATCH(3)")
    tampered2 = _shape_from({
        "shape_id": 2, "name": "TextBox 1", "kind": "text",
        "left_emu": 0, "top_emu": 0, "width_emu": 1000000, "height_emu": 1000000,
        "paragraphs": [{"text": "x",
                        "runs": [{"text": "x", "size_pt": 18.0, "bold": None,
                                  "italic": None, "font_name": None}],
                        "align": None, "level": 0}],
    })
    try:
        hl_layout.build_units(PageShapes(index=0, width_emu=0, height_emu=0, shapes=[tampered2]), 1920)
        print("  canvas width_emu=0 -> 未抛异常（意外）")
    except Exception as exc:                                   # noqa: BLE001
        print(f"  canvas width_emu=0 -> {type(exc).__name__}: {exc}")


if __name__ == "__main__":
    main()
