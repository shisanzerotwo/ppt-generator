"""行级高亮矩形：把"文本框外框"细化到"每一行文字真正占的矩形"。

为什么必须行级
--------------
spike 实测（PLAN_PPTX_ANIM §3.1）：文本框矩形内的墨迹覆盖率**中位数仅 0.087**，
p10 低到 0.005——最差的一页里 792×821px 的大框只装了一行小字。若直接拿
`shape.left/top/width/height` 做高亮，等于高亮一大片空白。宽度贴合度
`tight_w` 中位 0.983 但 p10=0.149，说明大量框宽是文字的 6 倍。

所以高亮矩形按**行**算：框内边距 + 每段字号 + FontTools 贪心换行 → 每行的
tight 矩形。度量必须与 `qa.py` 同源——qa 的度量原本是给 PPTX 溢出 QA 用的
（"算出来的行数"要逼近 PowerPoint 实际排出的行数），这里直接复用，两套
算法一旦漂移，溢出告警和高亮位置就会互相打架。

本模块只做纯算术：输入 `PageShapes`（EMU/pt/px 齐全），输出 `Unit`/`Rect`，
无 IO、无 COM、无浏览器。
"""

from dataclasses import dataclass, field

import qa
from pptx_io import PageShapes, ShapeInfo

EMU_PER_PT = qa.EMU_PER_PT


# ---------------------------------------------------------------- 断行层

@dataclass
class Line:
    text: str        # 该行实际承载的字符（不含被丢弃的行尾空格）
    width_pt: float  # 该行已放置宽度 = Σ _char_width_pt(ch, size_pt) + 行内空格
    size_pt: float


def _place_units_record(units: str, size_pt: float, box_width_pt: float,
                        lines_out: list[Line], cur_text: list[str], cur_w: float):
    """`qa._place_units` 的记录版：超长词/超宽单字逐字强拆，同时记下断点。

    返回末行的 (cur_text, cur_w)。中途换行时把已完成的行追加进 lines_out。
    """
    for ch in units:
        w = qa._char_width_pt(ch, size_pt)
        if cur_w + w <= box_width_pt or cur_w == 0:
            cur_text.append(ch)
            cur_w += w
        else:
            lines_out.append(Line("".join(cur_text), cur_w, size_pt))
            cur_text, cur_w = [ch], w
    return cur_text, cur_w


def wrap_lines(text: str, size_pt: float, box_width_pt: float) -> list[Line]:
    """复刻 `qa._measure_lines_ex` 的贪心规则，额外给出每行的文本与宽度。

    规则（与 qa 逐条对齐，注释里的变量名保持一致以免漂移）：
    - 拉丁字母/数字连成一个词，优先在空格处断行；CJK 与中文标点逐字可断；
    - 空格不立即计宽，累进 `pending_space`，随下一个内容 token 一起结算；
    - 行首空格丢弃（`cur == 0` 时 `pending_space = 0`）；
    - 超长单元（整词或单字宽 > 行宽）逐字强拆，新行仍放不下就独占多行。

    空文本返回单行空 Line（与 qa 的 lines=1 对齐）。
    等价性由 tests/test_hl_layout.py 的 8 类语料强制锁定。
    """
    tokens = qa._tokenize(text)
    if not tokens:
        return [Line("", 0.0, size_pt)]

    lines_out: list[Line] = []
    cur_text: list[str] = []
    cur_w = 0.0
    pending_space = 0.0
    pending_text: list[str] = []

    for kind, tok in tokens:
        if kind == "space":
            pending_space += qa._char_width_pt(" ", size_pt)
            pending_text.append(tok)
            continue

        if kind == "word":
            width = sum(qa._char_width_pt(c, size_pt) for c in tok)
        else:  # cjk / 单字符
            width = qa._char_width_pt(tok, size_pt)

        if cur_w == 0:
            pending_space = 0.0
            pending_text = []

        if cur_w + pending_space + width <= box_width_pt or cur_w + pending_space == 0:
            if cur_w == 0:
                cur_text, cur_w = _place_units_record(
                    tok, size_pt, box_width_pt, lines_out, cur_text, cur_w)
            else:
                cur_text.extend(pending_text)
                cur_text.append(tok)
                cur_w += pending_space + width
            pending_space = 0.0
            pending_text = []
            continue

        # 当前行放不下 → 换行，行首空格丢弃
        lines_out.append(Line("".join(cur_text), cur_w, size_pt))
        cur_text, cur_w = [], 0.0
        pending_space = 0.0
        pending_text = []
        if width > box_width_pt:
            # 超长单元：新行仍放不下，逐字强拆
            cur_text, cur_w = _place_units_record(
                tok, size_pt, box_width_pt, lines_out, cur_text, cur_w)
        else:
            cur_text = [tok]
            cur_w = width

    lines_out.append(Line("".join(cur_text), cur_w, size_pt))
    return lines_out


# ---------------------------------------------------------------- 定位层

@dataclass
class Rect:
    left_emu: int
    top_emu: int
    width_emu: int
    height_emu: int
    left_px: float
    top_px: float
    width_px: float
    height_px: float


@dataclass
class Unit:
    order: int                  # 页内讲解序（形状文档序 × 段序，0 基）
    page_index: int
    text: str
    kind: str                   # title|body|bullet|cell|chart|picture|table
    rect: Rect                  # 段落 union bbox（滚动定位/命中测试/QA 汇总用）
    lines: list[Rect]           # 每行 tight rect（渲染实体）
    size_pt: float
    align: str
    shape_id: int
    shape_name: str
    is_estimated: bool
    warnings: list[str] = field(default_factory=list)


def _make_rect(l_pt: float, t_pt: float, w_pt: float, h_pt: float,
               export_width_px: int, canvas_w_emu: int) -> Rect:
    """pt → EMU（权威）+ px（派生）。px 固定在"导出底图像素空间"。"""
    l_emu, t_emu = round(l_pt * EMU_PER_PT), round(t_pt * EMU_PER_PT)
    w_emu, h_emu = round(w_pt * EMU_PER_PT), round(h_pt * EMU_PER_PT)
    scale = export_width_px / canvas_w_emu
    return Rect(l_emu, t_emu, w_emu, h_emu,
                l_emu * scale, t_emu * scale, w_emu * scale, h_emu * scale)


def _para_size_pt(para) -> float:
    """段落字号：逐 run 取第一个显式值，全无回退 qa.DEFAULT_FONT_SIZE_PT。

    与 `qa._para_font_size` 同规则、同常量——另立回退值会让高亮与溢出 QA 打架。
    """
    for run in para.runs:
        if run.size_pt is not None:
            return float(run.size_pt)
    return qa.DEFAULT_FONT_SIZE_PT


def _paragraph_lines(text: str, size_pt: float, inner_w_pt: float,
                     word_wrap: bool | None) -> list[Line]:
    """一段 → 若干行。word_wrap is False 时不换行，只按手动换行拆。"""
    if word_wrap is False:
        return [Line(part, sum(qa._char_width_pt(c, size_pt) for c in part), size_pt)
                for part in text.split("\n")]
    if "\n" in text:
        # 段内手动换行（a:br）：先按硬换行断开，每段再各自贪心换行
        out: list[Line] = []
        for part in text.split("\n"):
            out.extend(wrap_lines(part, size_pt, inner_w_pt))
        return out
    return wrap_lines(text, size_pt, inner_w_pt)


def _page_median_size_pt(page: PageShapes) -> float:
    sizes: list[float] = []
    for shape in _iter_text_shapes(page.shapes):
        for para in shape.paragraphs:
            sizes.append(_para_size_pt(para))
    if not sizes:
        return 0.0
    sizes.sort()
    mid = len(sizes) // 2
    return sizes[mid] if len(sizes) % 2 else (sizes[mid - 1] + sizes[mid]) / 2


def _iter_text_shapes(shapes: list[ShapeInfo]):
    for shape in shapes:
        if shape.kind == "group":
            yield from _iter_text_shapes(shape.children)
        elif shape.kind == "text":
            yield shape


def _classify_text_kind(shape: ShapeInfo, size_pt: float, median_pt: float) -> str:
    """§4.6 标题判定：形状名含 Title/标题 → title；或字号 ≥ 1.3× 本页中位字号。

    契约还允许"is_placeholder 且 ph type ∈ {TITLE, CENTER_TITLE}"这一条，但
    ShapeInfo 没有携带 ph type（见 IMPL_REPORT 的契约缺陷节），故未实现。
    """
    name = (shape.name or "").lower()
    if "title" in name or "标题" in name:
        return "title"
    if median_pt > 0 and size_pt >= median_pt * 1.3:
        return "title"
    return "body"


def _bbox(rects: list[Rect]) -> Rect:
    l = min(r.left_emu for r in rects)
    t = min(r.top_emu for r in rects)
    r_ = max(r.left_emu + r.width_emu for r in rects)
    b = max(r.top_emu + r.height_emu for r in rects)
    l_px = min(r.left_px for r in rects)
    t_px = min(r.top_px for r in rects)
    r_px = max(r.left_px + r.width_px for r in rects)
    b_px = max(r.top_px + r.height_px for r in rects)
    return Rect(l, t, r_ - l, b - t, l_px, t_px, r_px - l_px, b_px - t_px)


class _UnitBuilder:
    """把一页的形状树摊成讲解单元，保持文档序。"""

    def __init__(self, page: PageShapes, export_width_px: int,
                 pad_x_pt: float, pad_y_pt: float, skipped: list | None = None):
        self.page = page
        self.export_width_px = export_width_px
        self.pad_x_pt = pad_x_pt
        self.pad_y_pt = pad_y_pt
        self.skipped = skipped
        self.canvas_w_emu = page.width_emu
        self.median_pt = _page_median_size_pt(page)
        self.is_estimated = not qa.font_available()
        self.units: list[Unit] = []

    def _add(self, *, text, kind, lines, size_pt, align, shape, warnings):
        rect = _bbox(lines) if lines else None
        if rect is None:
            return
        self.units.append(Unit(
            order=len(self.units), page_index=self.page.index, text=text, kind=kind,
            rect=rect, lines=lines, size_pt=size_pt, align=align,
            shape_id=shape.shape_id, shape_name=shape.name,
            is_estimated=self.is_estimated, warnings=list(warnings),
        ))

    def walk(self, shapes):
        for shape in shapes:
            if shape.kind == "group":
                self.walk(shape.children)
            elif shape.kind == "text":
                self._text(shape)
            elif shape.kind in ("picture", "chart"):
                self._block(shape)
            elif shape.kind == "table":
                self._table(shape)

    def _base_warnings(self, shape: ShapeInfo) -> list[str]:
        w: list[str] = []
        if shape.vertical_anchor is None:
            w.append("vertical_anchor_inherited")
        if shape.font_scale is not None:
            w.append("autofit_scaled")
        return w

    def _skip(self, shape: ShapeInfo, reason: str):
        """记一条"这个形状没产出 Unit、为什么"。

        契约 §4.3 对 `inner_w_pt <= 0` 的要求就是"**记 warning**，不产出 Unit" ——
        不传收集器时至少不再是"无声无息地少一块高亮"（审计 M3）。
        """
        if self.skipped is not None:
            self.skipped.append({
                "page_index": self.page.index, "shape_id": shape.shape_id,
                "shape_name": shape.name, "reason": reason,
            })

    def _text(self, shape: ShapeInfo):
        inner_left_pt = (shape.left_emu + shape.margin_left_emu) / EMU_PER_PT
        inner_top_pt = (shape.top_emu + shape.margin_top_emu) / EMU_PER_PT
        inner_w_pt = (shape.width_emu - shape.margin_left_emu - shape.margin_right_emu) / EMU_PER_PT
        inner_h_pt = (shape.height_emu - shape.margin_top_emu - shape.margin_bottom_emu) / EMU_PER_PT
        if inner_w_pt <= 0:
            return self._skip(shape, "no_inner_width")
        if inner_h_pt <= 0:
            return self._skip(shape, "no_inner_height")

        base_warn = self._base_warnings(shape)
        # 契约 §2.3 把 fontScale 存成 ÷1000 后的**百分数**（fontScale="60000" → 60.0），
        # 而 §4.3 写的是 "size_pt *= font_scale" —— 两者不自洽，直接相乘会放大 100 倍。
        # 这里按存储口径除以 100 取真实倍率（见 IMPL_REPORT 契约缺陷节）。
        font_scale = (shape.font_scale / 100.0) if shape.font_scale else 1.0
        own: list[tuple] = []          # (线条, 文本, 字号, 对齐, warnings)
        cursor_pt = 0.0
        for i, para in enumerate(shape.paragraphs):
            if not para.text.strip():
                continue
            size_pt = _para_size_pt(para) * font_scale
            if i > 0 and para.space_before_pt:
                # 首段段前距不计入：实测 PowerPoint 打开文本框时忽略它
                # （V5：首行墨迹距框顶仅 5.6pt ≈ margin_top，没有多出 20pt）
                cursor_pt += para.space_before_pt
            lines = _paragraph_lines(para.text, size_pt, inner_w_pt, shape.word_wrap)
            line_h_pt = size_pt * qa.LINE_HEIGHT_FACTOR * (para.line_spacing or 1.0)

            align = para.align or "LEFT"
            warn = list(base_warn)
            if align in ("JUSTIFY", "DISTRIBUTE"):
                align = "LEFT"
                warn.append("justify_approximated")

            rects: list[Rect] = []
            for line in lines:
                if align == "CENTER":
                    x = inner_left_pt + (inner_w_pt - line.width_pt) / 2
                elif align == "RIGHT":
                    x = inner_left_pt + (inner_w_pt - line.width_pt)
                else:
                    x = inner_left_pt
                rects.append(_make_rect(
                    x - self.pad_x_pt, inner_top_pt + cursor_pt - self.pad_y_pt,
                    line.width_pt + 2 * self.pad_x_pt, line_h_pt + 2 * self.pad_y_pt,
                    self.export_width_px, self.canvas_w_emu))
                cursor_pt += line_h_pt

            kind = _classify_text_kind(shape, size_pt, self.median_pt)
            if len(shape.paragraphs) > 1 and kind == "body":
                kind = "bullet"
            own.append((rects, para.text, size_pt, align, warn, kind,
                        para.space_after_pt or 0.0))
            if para.space_after_pt:
                cursor_pt += para.space_after_pt

        if not own:
            return

        # 垂直锚点：整块算完后统一平移（V4 实测 None 按 TOP 处理）
        consumed = cursor_pt
        if shape.vertical_anchor == "MIDDLE":
            dy = max(0.0, (inner_h_pt - consumed) / 2)
        elif shape.vertical_anchor == "BOTTOM":
            dy = max(0.0, inner_h_pt - consumed)
        else:
            dy = 0.0
        if dy:
            px_per_emu = self.export_width_px / self.canvas_w_emu
            own = [([_shift(r, dy, px_per_emu) for r in rects], *rest)
                   for rects, *rest in own]

        for rects, text, size_pt, align, warn, kind, _sa in own:
            self._add(text=text, kind=kind, lines=rects, size_pt=size_pt,
                      align=align, shape=shape, warnings=warn)

    def _block(self, shape: ShapeInfo):
        """图片/图表：整块一个单元（R4：图表不拆数据点）。图片本身即墨迹。"""
        rect = _make_rect(shape.left_emu / EMU_PER_PT, shape.top_emu / EMU_PER_PT,
                          shape.width_emu / EMU_PER_PT, shape.height_emu / EMU_PER_PT,
                          self.export_width_px, self.canvas_w_emu)
        self._add(text="", kind=shape.kind, lines=[rect], size_pt=0.0,
                  align="LEFT", shape=shape, warnings=self._base_warnings(shape))

    def _table(self, shape: ShapeInfo):
        """单元格级：每个非空单元格的每个段落一个 Unit。

        被并格（span 覆盖格）在 IR 里表现为空文本，天然被跳过；但合并 origin
        格的 rect 只能按**单列宽**算（IR 没带 span_width，见 IMPL_REPORT 契约缺陷）。
        """
        col_x = [0]
        for w in shape.table_col_widths_emu:
            col_x.append(col_x[-1] + w)
        row_y = [0]
        for h in shape.table_row_heights_emu:
            row_y.append(row_y[-1] + h)

        ml, mt = shape.margin_left_emu, shape.margin_top_emu
        mr, mb = shape.margin_right_emu, shape.margin_bottom_emu
        for r, row in enumerate(shape.table_cells):
            if r + 1 >= len(row_y):
                # 同族的静默丢弃：表格行列数组短于单元格矩阵时（被改坏的 deck.json），
                # 剩下的行/列会无声消失（审计 L2）。既然这里已经要堵"静默丢弃"，一并记上。
                self._skip(shape, "row_array_short")
                break
            for c, cell in enumerate(row):
                if c + 1 >= len(col_x):
                    self._skip(shape, "col_array_short")
                    break
                cell_w_pt = (col_x[c + 1] - col_x[c] - ml - mr) / EMU_PER_PT
                cell_h_pt = (row_y[r + 1] - row_y[r] - mt - mb) / EMU_PER_PT
                if cell_w_pt <= 0 or cell_h_pt <= 0:
                    self._skip(shape, "zero_cell")
                    continue
                left_pt = (shape.left_emu + col_x[c] + ml) / EMU_PER_PT
                top_pt = (shape.top_emu + row_y[r] + mt) / EMU_PER_PT

                cursor_pt = 0.0
                for para in cell:
                    if not para.text.strip():
                        continue
                    size_pt = _para_size_pt(para)
                    lines = _paragraph_lines(para.text, size_pt, cell_w_pt, True)
                    line_h_pt = size_pt * qa.LINE_HEIGHT_FACTOR * (para.line_spacing or 1.0)
                    align = para.align or "LEFT"
                    warn = self._base_warnings(shape)
                    if align in ("JUSTIFY", "DISTRIBUTE"):
                        align = "LEFT"
                        warn.append("justify_approximated")
                    rects = []
                    for line in lines:
                        if align == "CENTER":
                            x = left_pt + (cell_w_pt - line.width_pt) / 2
                        elif align == "RIGHT":
                            x = left_pt + (cell_w_pt - line.width_pt)
                        else:
                            x = left_pt
                        rects.append(_make_rect(
                            x - self.pad_x_pt, top_pt + cursor_pt - self.pad_y_pt,
                            line.width_pt + 2 * self.pad_x_pt,
                            line_h_pt + 2 * self.pad_y_pt,
                            self.export_width_px, self.canvas_w_emu))
                        cursor_pt += line_h_pt
                    self._add(text=para.text, kind="cell", lines=rects, size_pt=size_pt,
                              align=align, shape=shape, warnings=warn)
                    cursor_pt += para.space_after_pt or 0.0


def _shift(rect: Rect, dy_pt: float, px_per_emu: float) -> Rect:
    """整块垂直平移（pt 输入，EMU 与 px 同时平移）。"""
    d_emu = round(dy_pt * EMU_PER_PT)
    return Rect(rect.left_emu, rect.top_emu + d_emu, rect.width_emu, rect.height_emu,
                rect.left_px, rect.top_px + d_emu * px_per_emu,
                rect.width_px, rect.height_px)


def page_shapes(page, deck=None) -> PageShapes:
    """deck.json 里的 page dict（或已是 PageShapes）→ PageShapes。

    `DeckIR.pages` 每项只有 {index, bg, shapes}（契约 §2.3），**不含画布尺寸**，
    而 `build_units` 需要画布宽来把 EMU 换算成导出底图像素空间的 px。所以从
    deck 取 width_emu/height_emu 补上——这是契约数据模型里的一处缺口，见
    IMPL_REPORT 契约缺陷节。
    """
    if isinstance(page, PageShapes):
        return page
    if deck is None:
        raise ValueError(
            "page 是 deck.json 的 dict 时缺少画布尺寸，请把 DeckIR 一起传入："
            "page_shapes(page, deck)")
    return PageShapes(index=int(page["index"]), width_emu=int(deck.width_emu),
                      height_emu=int(deck.height_emu), shapes=list(page["shapes"]))


def build_units(page, export_width_px: int = 1920,
                pad_x_pt: float = 2.0, pad_y_pt: float = 1.0,
                skipped: list | None = None) -> list[Unit]:
    """一页的形状树 → 讲解单元列表（文档序，已按行级 tight 定位）。

    page 收 `PageShapes` 或 `dict`（需带 width_emu/height_emu，用 `page_shapes()`
    从 deck 构造）。pad_x_pt / pad_y_pt 是补偿量：spike 实测 12/22 行墨迹触框边，
    纯 tight 会把抗锯齿边缘切掉，留 2pt/1pt 的余量。

    `skipped`：可选的收集器。传 list 时，把"被丢弃、没产出 Unit"的形状追加进去
    （`{page_index, shape_id, shape_name, reason}`）。契约 §4.3 要求
    `inner_w_pt <= 0` 要"**记 warning**" —— 不给出口的话，用户在播放器里只会看到
    "有些文字没有高亮"，排查时没有任何线索（审计 M3）。reason 取值：
    `no_inner_width` / `no_inner_height` / `zero_cell`。
    """
    if not isinstance(page, PageShapes):
        raise ValueError("page 需为 PageShapes；deck.json 的 dict 请先过 page_shapes(page, deck)")
    builder = _UnitBuilder(page, export_width_px, pad_x_pt, pad_y_pt, skipped)
    builder.walk(page.shapes)
    return builder.units


# ---------------------------------------------------------------- 覆盖率度量

def measure_coverage(bg_png: str, rect_px: tuple, bg_rgb: tuple | None = None,
                     diff_thresh: int = 30) -> float:
    """rect_px = (left, top, width, height)，底图像素空间。

    bg_rgb 为 None 时取裁剪区**四边 1px 边框环的中位色**作底色（对深色主题鲁棒：
    底图是整张幻灯片，浅色/深色背景都能自适应）。
    ink = 任一通道 |px - bg| > diff_thresh 的像素。空矩形返回 0.0。
    """
    from PIL import Image

    left, top, width, height = (int(v) for v in rect_px)
    if width <= 0 or height <= 0:
        return 0.0
    with Image.open(bg_png) as im:
        img = im.convert("RGB")
        left = max(0, min(left, img.width))
        top = max(0, min(top, img.height))
        right = max(left, min(left + width, img.width))
        bottom = max(top, min(top + height, img.height))
        if right <= left or bottom <= top:
            return 0.0
        px = img.load()
        if bg_rgb is None:
            ring = []
            for x in range(left, right):
                ring.append(px[x, top])
                ring.append(px[x, bottom - 1])
            for y in range(top, bottom):
                ring.append(px[left, y])
                ring.append(px[right - 1, y])

            def median(ch):
                vals = sorted(v[ch] for v in ring)
                return vals[len(vals) // 2]

            bg_rgb = (median(0), median(1), median(2))

        ink = 0
        total = (right - left) * (bottom - top)
        for y in range(top, bottom):
            for x in range(left, right):
                p = px[x, y]
                if (abs(p[0] - bg_rgb[0]) > diff_thresh
                        or abs(p[1] - bg_rgb[1]) > diff_thresh
                        or abs(p[2] - bg_rgb[2]) > diff_thresh):
                    ink += 1
    return ink / total if total else 0.0
