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
