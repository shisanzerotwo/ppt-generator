"""pptx 读取（python-pptx） + COM 无窗口导出保真底图。

设计理由
--------
1. **为什么留着 PowerPoint 这个硬依赖**：`slide.shapes` 的坐标是文本框外框，与文字
   实际落墨位置无关（spike 实测框内墨迹覆盖率中位数仅 0.087）。要做到视觉 100%
   保真，底图必须由 PowerPoint 自己渲染（D8/D9）。python-pptx 只负责给出"文字在哪
   一行、哪一段"的几何信息，两者叠加成高亮播放器。
2. **EMU 是唯一权威坐标**：`*_pt` / `*_px` 都是派生量（见契约 §2.1）。只读实现不
   反向由 px 推 EMU——导出宽度换了（1280/1920/2560），px 全变，EMU 不变。
3. **组合形状必须做仿射换算**：`BaseShape.left` 在组合内返回的是**子坐标系原值**，
   不做任何变换（python-pptx 实测）。组用 `off/ext`（落在页面上的盒子）与
   `chOff/chExt`（子坐标系）两套矩形描述映射，契约 §2.2 给了公式，本模块逐层累乘。
4. **COM 实例隔离（实测结论，与常见认知相反）**：本机 PowerPoint 16.0 的 COM
   server 是"多用途"的——`DispatchEx` **不会**开新进程，与 `Dispatch` 拿到的是同一
   个实例（实测：往自称"用户实例"的对象里加稿，DispatchEx 实例看到的
   `Presentations.Count` 同步变化；POWERPNT.EXE 进程数始终为 1）。因此
   "DispatchEx 天然隔离"不成立，**安全性完全靠 `Quit` 前的守卫**：
   ①应用启动前就存在 POWERPNT.EXE 进程 → 绝不 Quit；②打开前 `Presentations.Count`
   不为 0 → 绝不 Quit。两道都过才 Quit（宁可留一个空进程，也不关掉用户没保存的稿）。
   同理**绝不触碰 `app.Visible`**——那会把用户正在看的窗口藏起来。
"""

import os
import sys
from dataclasses import asdict, dataclass, field

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn
from pptx.util import Length

import qa

SCHEMA_VERSION = 1

_EMU_PER_PT = qa.EMU_PER_PT
_DEFAULT_MARGIN_LR = 91440   # 0.1in = 7.2pt（python-pptx 默认内边距）
_DEFAULT_MARGIN_TB = 45720   # 0.05in = 3.6pt
_OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"  # 加密 Office 文件是 OLE2 复合文档
_ZIP_MAGIC = b"pk\x03\x04"

_NO_POWERPOINT_HINT = (
    "请安装 Microsoft Office（含 PowerPoint），或改用 --mode redesign / 直接跳过 faithful 模式。"
)


class PptxError(Exception):
    """唯一的对外异常。message 一律中文，供 CLI 直接打印/塞进 JSON。"""

    def __init__(self, message: str, code: str, hint: str = ""):
        self.message = message
        self.code = code
        self.hint = hint
        super().__init__(message)

    def to_dict(self) -> dict:
        return {"code": self.code, "message": self.message, "hint": self.hint}


# ---------------------------------------------------------------- 数据结构

@dataclass
class RunInfo:
    text: str
    size_pt: float | None
    bold: bool | None
    italic: bool | None
    font_name: str | None


@dataclass
class ParaInfo:
    text: str
    runs: list[RunInfo] = field(default_factory=list)
    align: str | None = None
    level: int = 0
    space_before_pt: float | None = None
    space_after_pt: float | None = None
    line_spacing: float | None = None


@dataclass
class ShapeInfo:
    shape_id: int
    name: str
    kind: str
    left_emu: int | None
    top_emu: int | None
    width_emu: int | None
    height_emu: int | None
    rotation_deg: float = 0.0
    is_placeholder: bool = False
    hidden: bool = False
    paragraphs: list[ParaInfo] = field(default_factory=list)
    margin_left_emu: int = _DEFAULT_MARGIN_LR
    margin_top_emu: int = _DEFAULT_MARGIN_TB
    margin_right_emu: int = _DEFAULT_MARGIN_LR
    margin_bottom_emu: int = _DEFAULT_MARGIN_TB
    word_wrap: bool | None = None
    vertical_anchor: str | None = None
    auto_size: str | None = None
    font_scale: float | None = None
    table_cells: list[list[list[ParaInfo]]] = field(default_factory=list)
    table_col_widths_emu: list[int] = field(default_factory=list)
    table_row_heights_emu: list[int] = field(default_factory=list)
    children: list["ShapeInfo"] = field(default_factory=list)


@dataclass
class PageShapes:
    index: int
    width_emu: int
    height_emu: int
    shapes: list[ShapeInfo] = field(default_factory=list)


@dataclass
class DeckIR:
    schema: int
    source_pptx: str
    mode: str
    export_width_px: int
    width_emu: int
    height_emu: int
    pages: list[dict] = field(default_factory=list)
    skipped: list[dict] = field(default_factory=list)


# ---------------------------------------------------------------- 段落 / run

def _enum_name(value) -> str | None:
    """PP_ALIGN 等枚举 → 名字字符串；None 保留 None（= 继承）。"""
    if value is None:
        return None
    return getattr(value, "name", str(value))


def _para_text(paragraph) -> str:
    """段落纯文本，走 XML 文档序：a:r/a:t 取字、a:br 换行、a:fld 取其 a:t。

    不能直接用 `paragraph.runs` / `paragraph.text`：它们不暴露 `a:br`，
    会丢掉段内手动换行（软换行在汇报稿里很常见）。
    """
    parts: list[str] = []
    for child in paragraph._p:
        tag = child.tag
        if tag == qn("a:r"):
            t = child.find(qn("a:t"))
            if t is not None and t.text:
                parts.append(t.text)
        elif tag == qn("a:br"):
            parts.append("\n")
        elif tag == qn("a:fld"):
            t = child.find(qn("a:t"))
            if t is not None and t.text:
                parts.append(t.text)
    return "".join(parts)


def _read_runs(paragraph) -> list[RunInfo]:
    runs: list[RunInfo] = []
    for run in paragraph.runs:
        size = run.font.size
        latin = None
        rPr = run._r.find(qn("a:rPr"))
        if rPr is not None:
            lat = rPr.find(qn("a:latin"))
            if lat is not None:
                latin = lat.get("typeface")
        runs.append(RunInfo(
            text=run.text,
            size_pt=size.pt if size is not None else None,
            bold=run.font.bold,
            italic=run.font.italic,
            font_name=latin,
        ))
    return runs


def _para_size_pt(paragraph) -> float:
    """段落字号：逐 run 取第一个显式值，全无则回退 qa.DEFAULT_FONT_SIZE_PT。

    与 `qa._para_font_size` 同规则（此处只为折算行距形态，不写进 ParaInfo）。
    """
    for run in paragraph.runs:
        if run.font.size is not None:
            return run.font.size.pt
    return qa.DEFAULT_FONT_SIZE_PT


def _len_pt(value) -> float | None:
    if value is None:
        return None
    try:
        return value.pt
    except AttributeError:
        return None


def _read_paragraph(paragraph) -> ParaInfo:
    ls = paragraph.line_spacing
    if ls is None:
        line_spacing = None
    elif isinstance(ls, Length):
        # Length（绝对值，如 Pt(30)）：契约 §2.3 规定折算为"倍数"再存；折不出来则置 None
        size = _para_size_pt(paragraph)
        line_spacing = (ls.pt / size) if size else None
    else:
        # 注意 Length 是 int 的子类，必须先判 Length 再判数值（否则 30pt 会被
        # 当成"381000 倍行距"存下去）
        line_spacing = float(ls)
    return ParaInfo(
        text=_para_text(paragraph),
        runs=_read_runs(paragraph),
        align=_enum_name(paragraph.alignment),
        level=int(paragraph.level or 0),
        space_before_pt=_len_pt(paragraph.space_before),
        space_after_pt=_len_pt(paragraph.space_after),
        line_spacing=line_spacing,
    )


def _read_paragraphs(text_frame) -> list[ParaInfo]:
    return [_read_paragraph(p) for p in text_frame.paragraphs]


# ---------------------------------------------------------------- 形状属性

_NV_TAGS = ("nvSpPr", "nvPicPr", "nvGraphicFramePr", "nvGrpSpPr", "nvCxnSpPr")


def _is_hidden(shape) -> bool:
    """`p:cNvPr/@hidden == "1"`。python-pptx 无公开属性，且各帧类元素的
    nv*Pr 子元素名不同，逐个试。取不到一律当可见。"""
    for tag in _NV_TAGS:
        nv = getattr(shape._element, tag, None)
        if nv is None:
            continue
        cNvPr = getattr(nv, "cNvPr", None)
        if cNvPr is not None:
            return cNvPr.get("hidden") == "1"
    return False


def _read_font_scale(text_frame) -> float | None:
    bodyPr = text_frame._txBody.find(qn("a:bodyPr"))
    if bodyPr is None:
        return None
    norm = bodyPr.find(qn("a:normAutofit"))
    if norm is None:
        return None
    raw = norm.get("fontScale")
    if raw is None:
        return None
    try:
        return int(raw) / 1000.0
    except ValueError:
        return None


def _shape_kind(shape) -> str:
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        return "group"
    if getattr(shape, "has_table", False) and shape.has_table:
        return "table"
    if getattr(shape, "has_chart", False) and shape.has_chart:
        return "chart"
    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
        return "picture"
    if getattr(shape, "has_text_frame", False) and shape.has_text_frame:
        return "text"
    return "other"


def _read_table(shape) -> tuple[list[list[list[ParaInfo]]], list[int], list[int]]:
    table = shape.table
    cells: list[list[list[ParaInfo]]] = []
    for r in range(len(table.rows)):
        row: list[list[ParaInfo]] = []
        for c in range(len(table.columns)):
            cell = table.cell(r, c)
            row.append(_read_paragraphs(cell.text_frame))
        cells.append(row)
    return cells, [int(col.width) for col in table.columns], [int(row.height) for row in table.rows]


def _has_visible_text(kind: str, paragraphs: list[ParaInfo],
                      cells: list[list[list[ParaInfo]]]) -> bool:
    if kind == "text":
        return any(p.text.strip() for p in paragraphs)
    if kind == "table":
        return any(p.text.strip() for row in cells for cell in row for p in cell)
    return True


# ---------------------------------------------------------------- 递归读形状

def _identity_rect(x, y, w, h):
    return x, y, w, h


def _group_mapper(group_left, group_top, group_w, group_h, cox, coy, cw, ch):
    """组内子坐标 → 组自身坐标系（§2.2）。"""
    sx = group_w / cw
    sy = group_h / ch

    def mapper(x, y, w, h):
        return (group_left + (x - cox) * sx,
                group_top + (y - coy) * sy,
                w * sx,
                h * sy)

    return mapper


def _iter_shapes(container):
    return list(container.shapes)


def _read_shape(shape, mapper, parent_rotation: float, skipped: list[dict],
                page_index: int, inherited_left=None, inherited_top=None):
    """读一个形状，返回 ShapeInfo；被过滤则返回 None 并在 skipped 里记一条。

    mapper 是"本层坐标 → 绝对 EMU"的仿射映射（父级已累乘进来）。
    """
    name = getattr(shape, "name", "")
    kind = _shape_kind(shape)

    left = shape.left if shape.left is not None else inherited_left
    top = shape.top if shape.top is not None else inherited_top
    width = shape.width
    height = shape.height

    def skip(reason: str):
        skipped.append({"page_index": page_index, "shape_name": name, "reason": reason})
        return None

    is_group = kind == "group"

    if is_group:
        # 组自身 xfrm 缺失（实测：删掉 a:xfrm 后四个属性全为 None）→ 整组跳过
        if left is None or top is None or width is None or height is None:
            return skip("group_no_xfrm")
        if width <= 0 or height <= 0:
            return skip("zero_size")
        cox, coy, cw, ch = _group_ch_off_ext(shape, left, top, width, height)
        if cw == 0 or ch == 0:
            return skip("degenerate_group")
        # 注意：_group_mapper 必须用**本层坐标原值**（left/top/width/height），
        # 它产出的是"组自身坐标系"里的矩形（= 父层子坐标系），再由外层 mapper
        # 送到幻灯片坐标。若喂绝对坐标，外层缩放会被重复施加（嵌套组合实测踩到）。
        child_mapper = _compose(mapper, _group_mapper(left, top, width, height,
                                                      cox, coy, cw, ch))
        a_left, a_top, a_w, a_h = mapper(left, top, width, height)
        rotation = float(getattr(shape, "rotation", 0.0) or 0.0)
        children = []
        for child in shape.shapes:
            info = _read_shape(child, child_mapper, parent_rotation + rotation, skipped, page_index)
            if info is not None:
                children.append(info)
        return ShapeInfo(
            shape_id=int(shape.shape_id), name=name, kind="group",
            left_emu=int(round(a_left)), top_emu=int(round(a_top)),
            width_emu=int(round(a_w)), height_emu=int(round(a_h)),
            rotation_deg=parent_rotation + rotation, hidden=_is_hidden(shape),
            children=children,
        )

    if left is None or top is None or width is None or height is None:
        return skip("no_xfrm")
    if width <= 0 or height <= 0:
        return skip("zero_size")
    a_left, a_top, a_w, a_h = mapper(left, top, width, height)
    if a_w <= 0 or a_h <= 0:
        return skip("zero_size")
    if _is_hidden(shape):
        return skip("hidden")

    rotation = float(getattr(shape, "rotation", 0.0) or 0.0)
    info = ShapeInfo(
        shape_id=int(shape.shape_id), name=name, kind=kind,
        left_emu=int(round(a_left)), top_emu=int(round(a_top)),
        width_emu=int(round(a_w)), height_emu=int(round(a_h)),
        rotation_deg=parent_rotation + rotation,
        is_placeholder=bool(getattr(shape, "is_placeholder", False)),
        hidden=False,
    )

    if kind == "text":
        tf = shape.text_frame
        info.paragraphs = _read_paragraphs(tf)
        info.margin_left_emu = int(tf.margin_left if tf.margin_left is not None else _DEFAULT_MARGIN_LR)
        info.margin_top_emu = int(tf.margin_top if tf.margin_top is not None else _DEFAULT_MARGIN_TB)
        info.margin_right_emu = int(tf.margin_right if tf.margin_right is not None else _DEFAULT_MARGIN_LR)
        info.margin_bottom_emu = int(tf.margin_bottom if tf.margin_bottom is not None else _DEFAULT_MARGIN_TB)
        info.word_wrap = tf.word_wrap
        info.vertical_anchor = _enum_name(tf.vertical_anchor)
        info.auto_size = _enum_name(tf.auto_size)
        info.font_scale = _read_font_scale(tf)
    elif kind == "table":
        info.table_cells, info.table_col_widths_emu, info.table_row_heights_emu = _read_table(shape)

    if not _has_visible_text(kind, info.paragraphs, info.table_cells):
        return skip("empty_text")
    return info


def _compose(outer, inner):
    """两个仿射映射复合：先 inner 再 outer（嵌套组合逐层累乘）。"""
    def mapper(x, y, w, h):
        return outer(*inner(x, y, w, h))
    return mapper


def _group_ch_off_ext(group, gx, gy, gw, gh):
    """组的 chOff/chExt；缺失时视为恒等映射（chOff=0、chExt=组自身尺寸）。"""
    el = group._element
    try:
        ch_off, ch_ext = el.chOff, el.chExt
    except (AttributeError, ValueError):
        return 0, 0, int(gw), int(gh)
    if ch_off is None or ch_ext is None:
        return 0, 0, int(gw), int(gh)
    return int(ch_off.x), int(ch_off.y), int(ch_ext.cx), int(ch_ext.cy)


# ---------------------------------------------------------------- 读整稿

def _classify_bad_package(path: str) -> PptxError:
    """打不开时区分"加密"与"损坏"：加密的 Office 文件是 OLE2 复合文档，
    不是 zip。靠文件头魔数判定，比猜异常类型可靠。"""
    try:
        with open(path, "rb") as f:
            head = f.read(8)
    except OSError as exc:
        return PptxError(f"PPTX 无法打开：{os.path.basename(path)}（{exc}）", "PPTX_UNREADABLE",
                         "确认文件未被其他程序占用")
    if head.startswith(_OLE_MAGIC):
        return PptxError("PPTX 已加密，无法读取", "PPTX_ENCRYPTED",
                         "请先用 PowerPoint 去掉打开密码再导出")
    if not head.startswith(_ZIP_MAGIC):
        return PptxError(f"PPTX 无法打开：{os.path.basename(path)}（不是有效的 .pptx 包）",
                         "PPTX_UNREADABLE",
                         "确认是 .pptx（非 .ppt/.pdf），且未被其他程序占用")
    return PptxError(f"PPTX 无法打开：{os.path.basename(path)}（包结构损坏）", "PPTX_UNREADABLE",
                     "用 PowerPoint 打开后另存一次，或换一份稿子")


def read_pages(pptx_path: str, mode: str = "faithful",
               export_width_px: int = 1920) -> tuple[DeckIR, list[dict]]:
    """逐页形状清单。返回 (DeckIR, skipped)；skipped 与 DeckIR.skipped 同对象。

    DeckIR.pages[*]["bg"] 为 None —— 底图由 `export_pages` 产出后由调用方回填
    （`read_pages` 不知道 out_dir）。失败抛 PptxError。
    """
    if not os.path.isfile(pptx_path):
        raise PptxError(f"找不到文件：{pptx_path}", "PPTX_NOT_FOUND",
                        "检查路径是否正确（建议用绝对路径）")
    abs_path = os.path.abspath(pptx_path)
    try:
        prs = Presentation(abs_path)
    except Exception as exc:  # noqa: BLE001  包损坏/加密统一在此归类
        raise _classify_bad_package(abs_path) from exc

    slides = list(prs.slides)
    if not slides:
        raise PptxError("PPTX 中没有任何幻灯片", "PPTX_EMPTY", "换一份有内容的稿子")

    width_emu = int(prs.slide_width)
    height_emu = int(prs.slide_height)
    skipped: list[dict] = []
    pages: list[dict] = []
    for index, slide in enumerate(slides):
        page = PageShapes(index=index, width_emu=width_emu, height_emu=height_emu)
        for shape in _iter_shapes(slide):
            info = _read_shape(shape, _identity_rect, 0.0, skipped, index)
            if info is not None:
                if info.kind == "other":
                    # 契约 §2.3 行 6：保号供 redesign 参考，不进 units
                    skipped.append({"page_index": index, "shape_name": info.name,
                                    "reason": "unsupported"})
                page.shapes.append(info)
        pages.append({"index": index, "bg": None, "shapes": page.shapes})

    deck = DeckIR(
        schema=SCHEMA_VERSION, source_pptx=abs_path, mode=mode,
        export_width_px=export_width_px, width_emu=width_emu, height_emu=height_emu,
        pages=pages, skipped=skipped,
    )
    return deck, skipped


# ---------------------------------------------------------------- deck.json

def deck_to_dict(deck: DeckIR) -> dict:
    return asdict(deck)


def _para_from(d: dict) -> ParaInfo:
    return ParaInfo(**{**d, "runs": [RunInfo(**r) for r in d.get("runs", [])]})


def _shape_from(d: dict) -> ShapeInfo:
    data = dict(d)
    data["paragraphs"] = [_para_from(p) for p in d.get("paragraphs", [])]
    data["table_cells"] = [[[_para_from(p) for p in cell]
                            for cell in row] for row in d.get("table_cells", [])]
    data["children"] = [_shape_from(c) for c in d.get("children", [])]
    return ShapeInfo(**data)


def load_deck(path: str) -> DeckIR:
    """读回 deck.json。结构不符抛 PptxError("IR_MISMATCH")。"""
    import json

    try:
        with open(path, encoding="utf-8") as f:
            raw = json.load(f)
    except OSError as exc:
        raise PptxError(f"读不到 deck.json：{path}（{exc}）", "IR_MISMATCH",
                        "重新执行 pptgen import 生成 deck.json") from exc
    except ValueError as exc:
        raise PptxError(f"deck.json 不是合法 JSON：{path}（{exc}）", "IR_MISMATCH",
                        "重新执行 pptgen import 生成 deck.json") from exc
    try:
        pages = [{"index": p["index"], "bg": p.get("bg"),
                  "shapes": [_shape_from(s) for s in p.get("shapes", [])]}
                 for p in raw["pages"]]
        return DeckIR(
            schema=int(raw.get("schema", SCHEMA_VERSION)),
            source_pptx=raw["source_pptx"], mode=raw["mode"],
            export_width_px=int(raw["export_width_px"]),
            width_emu=int(raw["width_emu"]), height_emu=int(raw["height_emu"]),
            pages=pages, skipped=list(raw.get("skipped", [])),
        )
    except (KeyError, TypeError, ValueError) as exc:
        raise PptxError(f"deck.json 结构不完整：{path}（{exc}）", "IR_MISMATCH",
                        "重新执行 pptgen import 生成 deck.json") from exc


# ---------------------------------------------------------------- PowerPoint COM

def powerpoint_available() -> bool:
    """廉价前置探测：pywin32 可导入 且 注册表存在 ProgID "PowerPoint.Application"。

    不启动 PowerPoint 进程（Dispatch 会真启一个实例，探一个布尔值不值得）。
    「探测到装了」不等于「能用」——export_pages 内部的失败仍单独归类。
    """
    if sys.platform != "win32":
        return False
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        return False
    try:
        import winreg
        winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "PowerPoint.Application")
        return True
    except OSError:
        return False


def _powerpoint_running() -> bool:
    """应用启动前是否已有 POWERPNT.EXE —— 决定我们有没有资格 Quit。"""
    if sys.platform != "win32":
        return False
    import subprocess
    try:
        out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/NH"],
                             capture_output=True, text=True, timeout=15).stdout
    except (OSError, subprocess.SubprocessError):
        return True  # 探不到时保守：当作"有"，绝不 Quit
    return "POWERPNT.EXE" in out.upper()


def export_pages(pptx_path: str, out_dir: str, width: int = 1920) -> list[str]:
    """COM 无窗口导出每页 PNG，返回**绝对路径**列表（页序一致），slide_1.png 起。

    高按画布比例取整：height = round(width * height_emu / width_emu)。
    失败抛 PptxError；部分成功时在 message 里带已完成页数。
    """
    if not os.path.isfile(pptx_path):
        raise PptxError(f"找不到文件：{pptx_path}", "PPTX_NOT_FOUND",
                        "检查路径是否正确（建议用绝对路径）")
    if not powerpoint_available():
        raise PptxError("未检测到 PowerPoint，无法导出保真底图", "NO_POWERPOINT",
                        _NO_POWERPOINT_HINT)
    if width <= 0:
        raise PptxError(f"导出宽度非法：{width}", "BAD_ARGS", "宽度需为正整数像素")

    import pythoncom
    import win32com.client

    # 相对路径会被 PowerPoint 按自己的 cwd 解析 → 一律绝对化（spike 实证）
    abs_pptx = os.path.abspath(pptx_path)
    out_dir = os.path.abspath(out_dir)
    os.makedirs(out_dir, exist_ok=True)
    had_powerpoint = _powerpoint_running()

    pythoncom.CoInitialize()
    app = None
    pres = None
    done: list[str] = []
    try:
        try:
            # DispatchEx 用意是"新实例"；实测本机仍会附着到用户实例（见模块头注释），
            # 所以隔离不能靠它，只能靠下面的 Quit 守卫。
            app = win32com.client.DispatchEx("PowerPoint.Application")
        except Exception as exc:  # noqa: BLE001
            raise PptxError("未检测到 PowerPoint，无法导出保真底图", "NO_POWERPOINT",
                            _NO_POWERPOINT_HINT) from exc
        pre_count = app.Presentations.Count
        try:
            # 无窗口靠 WithWindow=False；不要动 app.Visible（多版本 PowerPoint
            # 拒绝隐形实例，且会连带影响用户正在看/没保存的窗口）
            pres = app.Presentations.Open(abs_pptx, ReadOnly=True, Untitled=False,
                                          WithWindow=False)
        except Exception as exc:  # noqa: BLE001
            raise PptxError(
                f"PowerPoint 打开失败：{os.path.basename(abs_pptx)}"
                f"（已完成 {len(done)} 页）", "COM_EXPORT_FAILED",
                "检查 PowerPoint 是否被其他程序占用、是否弹出了对话框") from exc

        w_emu = int(pres.PageSetup.SlideWidth)
        h_emu = int(pres.PageSetup.SlideHeight)
        if w_emu <= 0 or h_emu <= 0:
            raise PptxError("PowerPoint 报告的画布尺寸非法", "COM_EXPORT_FAILED",
                            "在 PowerPoint 里检查幻灯片大小设置")
        height = round(width * h_emu / w_emu)
        total = int(pres.Slides.Count)
        for i in range(1, total + 1):
            png = os.path.abspath(os.path.join(out_dir, f"slide_{i}.png"))
            try:
                pres.Slides(i).Export(png, "PNG", width, height)
            except Exception as exc:  # noqa: BLE001
                raise PptxError(
                    f"第 {i} 页底图导出失败（已完成 {len(done)}/{total} 页）",
                    "COM_EXPORT_FAILED",
                    "检查 PowerPoint 是否被其他程序占用/是否弹出了对话框") from exc
            if not os.path.isfile(png) or os.path.getsize(png) == 0:
                raise PptxError(f"第{i}页底图未生成：{png}", "COM_EXPORT_FAILED",
                                "检查磁盘空间与 PowerPoint 是否被弹窗阻塞")
            done.append(png)
    finally:
        try:
            if pres is not None:
                pres.Close()
        except Exception:  # noqa: BLE001
            pass
        try:
            # 两道守卫都通过才 Quit：宁可留一个空进程，也不关掉用户没保存的稿
            if app is not None and not had_powerpoint and app.Presentations.Count == 0:
                app.Quit()
        except Exception:  # noqa: BLE001
            pass
        # 注：释放 PowerPoint 的 COM 代理时，本机 PowerPoint 会主动断开连接，
        # 触发一次 RPC_E_DISCONNECTED(0x80010108)。实测三种释放顺序（含完全不
        # Quit）都会出现，且从不外泄成 Python 异常——只有 pytest 的 faulthandler
        # 会把它打成 "Windows fatal exception" 到 stderr。属噪音，不用管。
        app = pres = None
        pythoncom.CoUninitialize()
    return done
