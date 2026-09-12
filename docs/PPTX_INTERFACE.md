# pptx 双向链路 · 接口契约（v1）

> 架构 agent 产出（2026-09-12）｜基线 commit `c188b63`｜上游：[PLAN_PPTX_ANIM.md](../PLAN_PPTX_ANIM.md) 的 D1–D16 / M1–M7
> 本文只定**契约**，不含实现。所有"现有符号"均为**实际读到的**（见 §1 出处表），无猜测。
> 环境已核实：`python-pptx 1.0.2`、`pywin32` 已在 `requirements.txt`（`pywin32>=306; platform_system=="Windows"`）且 venv 内 `import win32com.client` 成功、`fonttools>=4.55.0`、`playwright>=1.48` 均在依赖里。

---

## 1. 可复用的现有能力（已读，勿重复造）

| 出处 | 符号 | 签名 / 值 | 本计划用途 |
|---|---|---|---|
| `qa.py` | `EMU_PER_PT` | `12700` | 所有 EMU↔pt 换算 |
| `qa.py` | `LINE_HEIGHT_FACTOR` | `1.25` | 行高 = 字号 × 1.25 |
| `qa.py` | `DEFAULT_FONT_SIZE_PT` | `18.0` | 字号缺失回退值（**与溢出 QA 同源**） |
| `qa.py` | `OVERFLOW_TOLERANCE_PT` / `BOUNDARY_TOLERANCE_PT` | `2.0` / `0.5` | 框溢出、越界容差 |
| `qa.py` | `_load_font()` | `-> (cmap, hmtx, unitsPerEm) \| None`，`lru_cache(maxsize=1)` | 字体可用性判定 |
| `qa.py` | `_char_width_pt(ch, size_pt)` | `-> float`（pt），字体缺失时降级 空格0.33x / ASCII0.55x / 其他1.0x | **单字符宽度，行级定位的核心** |
| `qa.py` | `_tokenize(text)` | `-> list[(kind, token)]`，kind ∈ `word`/`cjk`/`space` | 断行策略（拉丁整词、CJK 逐字、空格可断） |
| `qa.py` | `_measure_lines_ex(text, size_pt, box_width_pt)` | `-> (行数, used_estimate: bool)` | 溢出 QA 行数（**注意：只给行数，不给断点**，见 §3.3） |
| `qa.py` | `measure_text_lines(...)` | `-> int` 公开版 | `pptx_out` 缩字号试算 |
| `qa.py` | `font_available()` | `-> bool` | `used_estimate` 标注 |
| `qa.py` | `_para_font_size(paragraph)` | 逐 run 取第一个显式字号，否则 `18.0` | **必须逐字复刻该规则**，否则高亮与 QA 打架 |
| `qa.py` | `check_pptx(path, expected_pages)` | `-> {"errors","warnings","used_estimate"}` | M6 导出后自检 |
| `video.py` | `ffmpeg_path()` / `ffprobe_path()` | `-> str \| None`（PATH→WinGet Links→WinGet Packages） | `pptgen video` 前置探测，**照抄这套探测思路做 `powerpoint_available()`** |
| `video.py` | `build_page_args(img, out, seconds=4.0, fps=25, size=(1280,720))` | `-> list[str]` 纯函数 | 单页片段 |
| `video.py` | `build_final_args(page_clips, out, seconds=4.0, fade=0.8, fps=25)` | `-> list[str]`；`n<2` 抛 `ValueError` | xfade 拼合 |
| `video.py` | `synthesize(shots, out_path, seconds=4.0, fade=0.8, progress_cb=None)` | `-> str`；`fade` 必须 `0<fade<seconds`，否则 `ValueError`；自动建/清 `out_path+"_parts"` | **`pptgen video` 直接复用，零改动** |
| `video.py` | `_run(args, err, timeout=300)` | 带超时，超时抛 `RuntimeError` | ffmpeg 执行 |
| `shot.py` | `shot_deck(html_path, out_dir, min_settle_ms=300, max_settle_ms=2000)` | `-> list[str]`；要求页面有 `.slide`，否则 `ValueError`；用 `Path(...).as_uri()` | 只服务**设计稿 HTML**，**不适用**高亮播放器（见 §6.4） |
| `shot.py` | `_launch_browser(p)` | 依次试 `channel="chrome"/"msedge"`，全失败抛 `RuntimeError`（含中文安装指引） | 截图必需；`hl_anim` 要复用 |
| `shot.py` | `_screenshot_settled(page, min_ms, max_ms)` | `-> bytes`：`wait_for_timeout` + `screenshot(animations="disabled")` + 连续两帧字节一致轮询 | 三重确定化的三分之二 |
| `anim.py` | `build_player(out_dir, deck_web_path, title="", auto_step_ms=1200)` | `-> str`；`os.makedirs(out_dir, exist_ok=True)` | **转义纪律与单遍替换的范例，照抄** |
| `anim.py` | `TEMPLATE` 的 `__TITLE__`/`__DECK__`/`__CONFIG__` | `re.sub(..., lambda m: mapping[m.group(0)], TEMPLATE)` 单遍替换 | 同上 |
| `builder.py` | `_set_ea(run, font="微软雅黑")` | 补写 `a:ea`，**必须排在 `a:latin` 之后** | `pptx_out` 中文字体（见 §7.4） |
| `builder.py` | `SLIDE_W`/`SLIDE_H`/`FONT`/`THEMES` | `SLIDE_W=Inches(13.333)`、`SLIDE_H=Inches(7.5)`、`FONT="微软雅黑"`、`THEMES` 三套色板 | 版式尺寸基准 |
| `builder.py` | `_CHART_TYPES` | `{"bar","column","pie","line"} → XL_CHART_TYPE.*` | `pptx_out` 图表选型对齐（只对齐取值，不改 builder） |
| `builder.py` | `build_ppt(slides, image_paths, out_path, theme="blue", subtitle="")` | `-> str` | 保留不动（D11 双导出） |
| `llm_util.py` | `llm_client(timeout=60.0)` | `-> ZhipuAI`；无 key 抛 `RuntimeError("未配置 ZHIPUAI_API_KEY")` | redesign / deck 命令 |
| `llm_util.py` | `get_model(kind)` | kind ∈ `chat`/`vision`/`image`/`design` | 同上 |
| `llm_util.py` | `images_enabled()` | `-> bool`（`runtime_config.json` 的 `images` 键，缺省 True） | redesign 需尊重该开关 |
| `app.py` | `_video_jobs: set[str]` + `threading.Lock` | 同稿并发守卫范例 | CLI 不共享该 state（D2：B 方向不复用 Flask 全局） |
| `app.py` | `OUTPUT_DIR/IMAGES_DIR/DECKS_DIR/PROJECTS_DIR/ANIMATION_DIR/VIDEOS_DIR` | 均为绝对路径（`os.path.join(BASE_DIR, "output", ...)`） | CLI 默认输出根 |
| `template.py` | `_custom_path(key)` | `custom:<8位hex>` → `CUSTOM_DIR/<id>.json`；`re.fullmatch(r"[0-9a-f]{8}", cid)` 杜绝路径逃逸 | `--template` 的路径校验范式 |

**已实测的 python-pptx 行为**（`output/b_multislide.pptx`，python-pptx 1.0.2）：

| 事实 | 实测值 |
|---|---|
| `TextFrame` 默认内边距 | `margin_left/right = 91440 EMU`（0.1in = 7.2pt）、`margin_top/bottom = 45720 EMU`（0.05in = 3.6pt）；未显式设置时 `python-pptx` 返回上述**默认值**而非 `None` |
| `BaseShape.left/width` | `return self._element.x / .cx` —— **组合内子形状返回的是子坐标系原值，不做任何变换** |
| `GroupShape.left` 可为 `None` | `CT_GroupShape._get_xfrm_attr`：`xfrm = self.xfrm; if xfrm is None: return None` |
| 组合子坐标空间 | `CT_GroupShape` 存在 `chOff` / `chExt`，子形状坐标相对该空间（§2.2 换算公式） |
| 隐藏标记 | `<p:cNvPr hidden="1">`；python-pptx **无公开 `hidden` 属性**，须走 `shape._element.nvSpPr.cNvPr.get("hidden")`（缺失返回 `None`） |
| `TableCell` | 有 `margin_left/top/right/bottom`（同为 91440/45720）、`is_merge_origin`、`is_spanned`、`span_width`、`span_height`；`table.columns[i].width`、`table.rows[j].height` 可取 |
| 默认模板 11 个 layout | `0 Title Slide`(idx0 CENTER_TITLE, idx1 SUBTITLE)、`1 Title and Content`(idx0 TITLE, idx1 OBJECT)、`2 Section Header`、`3 Two Content`、`4 Comparison`、`5 Title Only`(仅 idx0 TITLE)、`6 Blank`、`7 Content with Caption`、`8 Picture with Caption`(idx1 PICTURE)、`9/10 Vertical *` |
| **默认模板主题无 CJK 字体** | `theme1.xml`：`majorFont/minorFont` 的 `a:latin="Calibri"`、**`a:ea=""`（空）**，且 `<a:font script="Hans" typeface="宋体"/>` → 占位符继承文本的中文渲染为**宋体**，不是微软雅黑（§7.4） |

---

## 2. 共用数据契约（单位约定先定死）

### 2.1 单位

| 单位 | 定义 |
|---|---|
| `*_emu` | `int`，OOXML 原生 EMU，**唯一权威坐标** |
| `*_pt` | `float`，`pt = emu / qa.EMU_PER_PT` |
| `*_px` | `float`，**固定在"导出底图像素空间"**：`px = emu × export_width_px / canvas_width_emu`。播放器不重算坐标，靠 CSS `transform: scale()` 缩放整个舞台 |
| 换算常量 | `px_per_pt = export_width_px / (canvas_width_emu / EMU_PER_PT)`；1920 宽 @13.333in 画布时 `= 2.0`（1pt 恰好 2px） |

**规则：EMU 是源，pt/px 是派生。** 只读实现不得反向由 px 推 EMU。

### 2.2 组合形状坐标换算（R4）

子形状的 `left/top` 位于**组的子坐标系**，须做仿射映射：

```
gx, gy, gw, gh  = group.left_emu, group.top_emu, group.width_emu, group.height_emu
cox, coy, cw, ch = group.ch_off_x_emu, group.ch_off_y_emu, group.ch_ext_cx_emu, group.ch_ext_cy_emu

abs_x = gx + (child_left - cox) * gw / cw
abs_y = gy + (child_top  - coy) * gh / ch
abs_w = child_w * gw / cw
abs_h = child_h * gh / ch
```

- 取 `chOff/chExt` 走 `group._element.chOff` / `.chExt`（`CT_GroupShape` 已提供，实测存在）；缺失时视为 `chOff=(0,0)`、`chExt=(gw,gh)`（即恒等映射）。
- **嵌套组合逐层累乘**：递归时把父级的映射函数作为参数传下去，不要只对顶层做一次。
- `gw/cw` 或 `gh/ch` 为 0 → 该组整体跳过并记 `skipped(reason="degenerate_group")`，**不要除零**。
- 组自身 `left is None` → 整组跳过（记 `skipped(reason="group_no_xfrm")`），不阻断其余页。
- 子形状的 `rotation` 与组 `rotation` 需叠加（`child_rot + group_rot`），本期只**透传角度**、由播放器用 `transform: rotate()` 近似，不做坐标重算（列 §10 待验证）。

### 2.3 结构（全部 `@dataclass`，字段名即 JSON 键名）

```python
# ---- 文本原子 ----
@dataclass
class RunInfo:
    text: str
    size_pt: float | None          # None = 继承（由 hl_layout 按 qa 规则回退）
    bold: bool | None
    italic: bool | None
    font_name: str | None          # a:latin；不含 a:ea（本期不区分中西文字体）

@dataclass
class ParaInfo:
    text: str                      # 段落纯文本；文档序拼接 a:r/a:t 与 a:br(→ "\n")
    runs: list[RunInfo]
    align: str | None              # "LEFT"/"CENTER"/"RIGHT"/"JUSTIFY"/"DISTRIBUTE"，None = 继承
    level: int                     # 0 = 一级
    space_before_pt: float | None
    space_after_pt: float | None
    line_spacing: float | None     # 仅记录"倍数"形态；Length 形态折算为 pt/字号 后再存

# ---- 形状 ----
@dataclass
class ShapeInfo:
    shape_id: int
    name: str
    kind: str                      # "text"|"picture"|"table"|"chart"|"group"|"other"
    left_emu: int | None
    top_emu: int | None
    width_emu: int | None
    height_emu: int | None
    rotation_deg: float            # shape.rotation，已叠加父组
    is_placeholder: bool
    hidden: bool                   # cNvPr/@hidden == "1"
    # kind == "text" 时有效
    paragraphs: list[ParaInfo] = []
    margin_left_emu: int = 91440
    margin_top_emu: int = 45720
    margin_right_emu: int = 91440
    margin_bottom_emu: int = 45720
    word_wrap: bool | None = None  # None = 继承（文本框有效默认 True）
    vertical_anchor: str | None = None   # "TOP"/"MIDDLE"/"BOTTOM"/None
    auto_size: str | None = None         # "NONE"/"SHAPE_TO_FIT_TEXT"/.../None
    font_scale: float | None = None      # a:bodyPr/a:normAutofit/@fontScale ÷ 1000；无 = None
    # kind == "table"
    table_cells: list[list[list[ParaInfo]]] = []   # [row][col] → 该单元格的段落列表
    table_col_widths_emu: list[int] = []
    table_row_heights_emu: list[int] = []
    # kind == "group"（children 的坐标已按 §2.2 换算为**绝对 EMU**）
    children: list["ShapeInfo"] = []

# ---- 页与稿 ----
@dataclass
class PageShapes:
    index: int                     # 0 基
    width_emu: int
    height_emu: int
    shapes: list[ShapeInfo]        # 已展开组合、已过滤

@dataclass
class DeckIR:
    schema: int                    # 常量 1
    source_pptx: str               # 绝对路径
    mode: str                      # "faithful"|"redesign"
    export_width_px: int
    width_emu: int
    height_emu: int
    pages: list[dict]              # {index, bg: "bg/slide_1.png"|None, shapes: [ShapeInfo…]}
    skipped: list[dict]            # [{page_index, shape_name, reason}]，供 stderr 汇报
```

**过滤规则（导入期一次性执行，`read_pages` 内）**，按序判定：

| # | 条件 | reason 标记 |
|---|---|---|
| 1 | `left/top/width/height` 任一为 `None` | `no_xfrm`（组合见 §2.2 例外） |
| 2 | `width_emu <= 0 or height_emu <= 0` | `zero_size` |
| 3 | `hidden is True` | `hidden` |
| 4 | `kind == "text"` 且所有段落 `text.strip() == ""` | `empty_text`（spike：40 框里 18 个命中） |
| 5 | `kind == "table"` 且所有单元格文本为空 | `empty_text` |
| 6 | `kind == "other"` | `unsupported`（保号，供 redesign 参考，不进 units） |

> **不要**过滤 `kind == "picture"`（要作为讲解单元）与 `kind == "chart"`（R4：整块一个单元）。

---

## 3. `pptx_io.py` —— pptx 读取 + COM 底图导出

### 3.1 错误类型

```python
class PptxError(Exception):
    """唯一的对外异常。message 一律中文，供 CLI 直接打印/塞进 JSON。"""
    def __init__(self, message: str, code: str, hint: str = ""):
        self.message = message; self.code = code; self.hint = hint
        super().__init__(message)

    def to_dict(self) -> dict:
        return {"code": self.code, "message": self.message, "hint": self.hint}
```

`code` 取值（与 §9 的 CLI 错误码同表，实现期不得自行扩写）：

| code | 触发 |
|---|---|
| `PPTX_NOT_FOUND` | 路径不存在 |
| `PPTX_UNREADABLE` | `Presentation()` 抛异常 / 非 zip / 损坏 |
| `PPTX_ENCRYPTED` | 打开需要密码（python-pptx 报 `PackageNotFoundError` 等，需在异常路径里识别） |
| `PPTX_EMPTY` | `len(prs.slides) == 0` |
| `NO_POWERPOINT` | COM 不可用（未装 Office / pywin32 缺失） |
| `COM_EXPORT_FAILED` | 打开或导出过程中 COM 报错 |
| `NO_FFMPEG` / `NO_BROWSER` | 由 `cli` 侧探测后构造 |
| `TEMPLATE_INVALID` | 自定义模板不可用 |
| `IR_MISMATCH` | deck.json 内部不一致（页数 ≠ 底图数等） |
| `BAD_ARGS` | 参数组合非法 |

### 3.2 `read_pages`

```python
def read_pages(pptx_path: str) -> tuple[DeckIR, list[dict]]:
    """逐页形状清单。
    返回 (DeckIR, skipped) —— DeckIR.pages[*]["shapes"] 已展开组合并过滤；
    skipped 与 DeckIR.skipped 同对象，便于调用方单独汇报。
    失败抛 PptxError。"""
```

实现要点（每条都对应一个真实 API 行为，见 §1）：

1. `prs = Presentation(os.path.abspath(pptx_path))`；`width_emu = int(prs.slide_width)`、`height_emu = int(prs.slide_height)`。
2. 逐 `slide.shapes` 递归；判定顺序：`MSO_SHAPE_TYPE.GROUP` → 递归 `shape.shapes`（`GroupShape.shapes` 是 `lazyproperty`，可迭代）；`shape.has_table` → `table`；`shape.has_chart` → `chart`；`shape.shape_type == MSO_SHAPE_TYPE.PICTURE` → `picture`；`shape.has_text_frame` → `text`；否则 `other`。
3. 顺序即 `order` 的天然来源：**文档序**（`slide.shapes` 的迭代顺序），不要按坐标排序。
4. 段落文本必须走 XML：`paragraph._p` 的子元素按文档序 `a:r/a:t` → 文本、`a:br` → `"\n"`、`a:fld` → 取其 `a:t`。**`paragraph.runs` 不暴露 `a:br`**，直接用它会丢掉段内手动换行。
5. `size_pt` 逐 run 取 `run.font.size.pt`（`None` 保留 `None`，**不要在此回退**，回退集中在 `hl_layout`）。
6. `line_spacing`：`paragraph.line_spacing` 可能是 `float`（倍数）或 `Length`（绝对值）——绝对值需除该段字号折算成倍数再存，存不出则置 `None`（列 §10）。
7. `auto_size`/`word_wrap`/`vertical_anchor` 直接取 `text_frame` 对应属性（实测可能为 `None`，**保留 `None` 语义 = 继承**）。
8. `font_scale`：解析 `text_frame._txBody.find(qn("a:bodyPr"))` 下的 `a:normAutofit`，读 `@fontScale` 后 `÷1000`。
9. `hidden`：`shape._element.nvSpPr.cNvPr.get("hidden") == "1"`；`nvSpPr` 不存在的帧类形状（如 `GraphicFrame`）按 `False`。
10. **不做**任何布局推断（那是 `hl_layout` 的事）。

### 3.3 `powerpoint_available`

```python
def powerpoint_available() -> bool:
    """廉价前置探测：pywin32 可导入 且 注册表存在 ProgID "PowerPoint.Application"。
    不启动 PowerPoint 进程（Dispatch 会真启一个实例，探一个布尔值不值得）。
    """
```

- 注册表：`winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "PowerPoint.Application")`，`try/except OSError → False`。
- 非 Windows（`sys.platform != "win32"`）直接 `False`。
- **探测"装了"不等于"能用"**：`export_pages` 内部的 `Dispatch` 失败仍须独立包成 `PptxError(code="NO_POWERPOINT")`。
- 不一致的兜底文案（中文，必须含可执行指引）：
  `"未检测到 PowerPoint，无法导出保真底图。请安装 Microsoft Office（含 PowerPoint），或改用 --mode redesign / 直接跳过 faithful 模式。"`

### 3.4 `export_pages`

```python
def export_pages(pptx_path: str, out_dir: str, width: int = 1920) -> list[str]:
    """COM 无窗口导出每页 PNG，返回**绝对路径**列表（页序一致），文件名为 slide_1.png 起。
    width 为像素宽，高按画布比例取整：height = round(width * height_emu / width_emu)。
    失败抛 PptxError；部分成功时在 message 里带已完成页数。"""
```

硬性要求（spike 实证 + 资源安全）：

1. **绝对路径**：`os.path.abspath(pptx_path)`、每页输出路径也 `abspath`。相对路径会被 PowerPoint 按自己的 cwd 解析 → 报"找不到文件"。
2. 用 `win32com.client.DispatchEx("PowerPoint.Application")` 而非 `Dispatch` —— `Dispatch` 会**附着到用户正在用的 PowerPoint 实例**，随后的 `Quit()` 会关掉用户没保存的稿子。`DispatchEx` 强制新实例。
3. 打开：`app.Presentations.Open(abs_path, ReadOnly=True, Untitled=False, WithWindow=False)`。"无窗口"靠 `WithWindow=False`，**不要**依赖 `app.Visible = False`（多版本 PowerPoint 拒绝隐形实例，见 §10）。
4. 导出：`pres.Slides(i).Export(abs_png, "PNG", width, height)`。
5. **`Close()`/`Quit()` 必须在 `finally` 里**，且遵守"不关别人的实例"规则：
   ```
   pre_count = app.Presentations.Count      # Open 之前记
   try:  … Open / Export …
   finally:
       try: pres.Close()
       except Exception: pass
       if pre_count == 0:                   # 打开前一个都没有 → 这个实例是我们的
           try: app.Quit()
           except Exception: pass
       pres = app = None
   ```
6. **线程安全**：COM 在 worker 线程里必须 `pythoncom.CoInitialize()` / `finally: pythoncom.CoUninitialize()`（计划 R5 的 >50 页后台路径会踩到）。
7. 每页导出后校验文件存在（`os.path.isfile`）且非 0 字节，否则 `PptxError("COM_EXPORT_FAILED", f"第{i}页底图未生成：…")`。
8. 返回值为**绝对路径列表**，`len(...) == len(prs.slides)` —— 调用方据此做页数一致性断言，不依赖调用方自己数。
9. `out_dir` 用 `os.makedirs(out_dir, exist_ok=True)`；**不清理已有同名文件**（同稿重导会覆盖，属预期）。

---

## 4. `hl_layout.py` —— 行级高亮矩形（本期最高风险点）

### 4.1 为什么必须行级

spike 实测：文本框矩形内的墨迹覆盖率**中位数 0.087**（p10=0.005），最差案例框 792×821px 里只有一行小字（coverage 0.005）；宽度贴合度 `tight_w` 中位 0.983 但 p10=0.149 —— 即**大量框宽是文字的 6 倍**。直接拿 `shape.left/top/width/height` 做高亮 = 高亮一大片空白。

### 4.2 断行：`qa` 只给行数，必须自己补断点

`qa._measure_lines_ex` 返回 `(行数, used_estimate)`，**没有断行位置**。要让高亮矩形贴合每一行，须复刻同一套贪心规则并记录断点：

```python
@dataclass
class Line:
    text: str          # 该行实际承载的字符（不含被丢弃的行尾空格）
    width_pt: float    # 该行已放置宽度 = Σ _char_width_pt(ch, size_pt) + 行内空格
    size_pt: float

def wrap_lines(text: str, size_pt: float, box_width_pt: float) -> list[Line]:
    """复刻 qa._measure_lines_ex 的贪心规则（qa._tokenize + _place_units 同策略），
    额外返回每行的文本与宽度。空文本返回单行空 Line（与 qa 的 lines=1 对齐）。"""
```

**强制性等价测试**（防两套算法漂移）：

```python
assert len(wrap_lines(t, s, w)) == qa_mod.measure_text_lines(t, s, w)
# 对以下语料各跑一遍：纯中文、纯英文、中英混排、含超长无空格英文词、
# 含全角标点、含行首/行尾空格、单字宽 > 行宽、空串
```

实现等价性的关键点（`_measure_lines_ex` 的三个易错处，必须一致）：
- **行首空格丢弃**：`cur == 0` 时 `pending_space = 0.0`；
- 空格不立即计宽，累进 `pending_space`，随下一个内容 token 一起结算；
- 超长单元（整词或单字宽 > 行宽）走"逐字强拆"（`_place_units` 语义），且**新行仍放不下时**强行独占多行。

### 4.3 `build_units`

```python
@dataclass
class Rect:
    left_emu: int; top_emu: int; width_emu: int; height_emu: int
    left_px: float; top_px: float; width_px: float; height_px: float

@dataclass
class Unit:
    order: int                  # 页内讲解序（= 形状文档序 × 段序，0 基）
    page_index: int
    text: str
    kind: str                   # "title"|"body"|"bullet"|"cell"|"chart"|"picture"|"table"
    rect: Rect                  # 段落 union bbox（见 4.4 的用途说明）
    lines: list[Rect]           # 每行 tight rect（渲染默认用这个）
    size_pt: float
    align: str                  # "LEFT"|"CENTER"|"RIGHT"（None 已归一为 "LEFT"）
    shape_id: int
    shape_name: str
    is_estimated: bool          # 字体度量走了降级估算（qa._load_font() is None）
    warnings: list[str]         # 例如 "vertical_anchor_inherited" / "autofit_scaled"

def build_units(page: PageShapes, export_width_px: int = 1920,
                pad_x_pt: float = 2.0, pad_y_pt: float = 1.0) -> list[Unit]:
```

**坐标推导（逐形状，单位全程 pt，最后一次性转 EMU）**：

```
inner_left_pt = (shape.left_emu + margin_left_emu) / EMU_PER_PT
inner_top_pt  = (shape.top_emu  + margin_top_emu)  / EMU_PER_PT
inner_w_pt    = (shape.width_emu  - ml - mr) / EMU_PER_PT
inner_h_pt    = (shape.height_emu - mt - mb) / EMU_PER_PT
```

`inner_w_pt <= 0` → 跳过该形状（记 warning，不产出 Unit）。

逐段落累加（`cursor_pt` 从 0 起）：

| 步骤 | 规则 |
|---|---|
| 段前距 | `cursor_pt += space_before_pt or 0`（**仅非首段生效**，首段是否计入见 §10） |
| 字号 | **复用 `qa._para_font_size` 的规则**：取该段第一个有显式 `font.size` 的 run；全无 → `qa.DEFAULT_FONT_SIZE_PT`（18.0）。**不要另立回退值**，否则高亮与溢出 QA 会互相矛盾 |
| 断行 | `word_wrap is False` → 不换行，`text.split("\n")` 每段一行；否则 `wrap_lines(text, size_pt, inner_w_pt)` |
| 行高 | `line_h_pt = size_pt * qa.LINE_HEIGHT_FACTOR * (line_spacing or 1.0)` |
| 行宽 | `line.width_pt`（`wrap_lines` 给出，行尾丢弃的空格不计） |
| 水平定位 | `LEFT`/None → `x = inner_left_pt`；`CENTER` → `x = inner_left_pt + (inner_w_pt - width_pt)/2`；`RIGHT` → `x = inner_left_pt + (inner_w_pt - width_pt)`；`JUSTIFY`/`DISTRIBUTE` → 按 `LEFT` 处理并在 `warnings` 里记 `justify_approximated` |
| 行矩形 | `left_emu = round((x - pad_x_pt) * EMU_PER_PT)`、`top_emu = round((inner_top_pt + cursor_pt - pad_y_pt) * EMU_PER_PT)`、`width_emu = round((width_pt + 2*pad_x_pt) * EMU_PER_PT)`、`height_emu = round((line_h_pt + 2*pad_y_pt) * EMU_PER_PT)` |
| 推进 | `cursor_pt += line_h_pt` |
| 段后距 | `cursor_pt += space_after_pt or 0` |

**垂直锚点**（覆盖 `builder.py` 大量 `vertical_anchor=MIDDLE` 的实际稿，实测首形状即 `MIDDLE`）：

```
if vertical_anchor == "MIDDLE": dy = max(0.0, (inner_h_pt - consumed_pt) / 2)
elif vertical_anchor == "BOTTOM": dy = max(0.0, inner_h_pt - consumed_pt)
else: dy = 0.0                      # TOP / None（None 的继承语义见 §10）
所有 rect.top_emu += round(dy * EMU_PER_PT)     # 需在算完本形状再统一平移
```

`font_scale is not None` → 所有 `size_pt *= font_scale`，并给每个 Unit 记 `autofit_scaled` warning（见 §10）。

### 4.4 宽度取值：推荐**行级 tight**，不是整行宽度

| 方案 | 覆盖率 | 观感 |
|---|---|---|
| 整行宽度（`inner_w_pt`） | 与形状级同病——短行后面拖一片空白 | 高亮条整齐，但"高亮空白"仍在 |
| **行级 tight（推荐）** | 单行矩形只包住该行墨迹，coverage 天然高 | 高亮像荧光笔，跟着文字走 |

**推荐行级 tight，理由**：M2 的验收指标就是"高亮块内墨迹 coverage 中位数 ≥ 0.35"——这是覆盖率指标，tight 是唯一能稳定达标的取法；且 ragged-right（段末短行）是中文汇报稿的常态，整行宽度在段末行必然塌到 0.1 量级。留 `pad_x_pt=2.0` 补偿墨迹触框边（spike：**12/22 行墨迹触框边**，边界判定本就需容差，纯 tight 会把抗锯齿边缘切掉）。

**渲染建议**：播放器按 `Unit.lines` 画**多个 rect**（每行一个），`Unit.rect` 只作 union bbox 用于滚动定位、命中测试与 QA 汇总。单行段落时两者等价。

### 4.5 覆盖率度量（M2 验收的可复现定义）

```python
def measure_coverage(bg_png: str, rect_px: tuple[int, int, int, int],
                     bg_rgb: tuple[int, int, int] | None = None,
                     diff_thresh: int = 30) -> float:
    """rect_px = (left, top, width, height) 在底图像素空间。
    bg_rgb 为 None 时取裁剪区**四边 1px 边框环的中位色**作为底色（对深色主题鲁棒）。
    ink = 任一通道 |px - bg| > diff_thresh 的像素。
    return ink_count / (width*height)；空矩形返回 0.0。"""
```

- 需要 `Pillow`（已在依赖）与 `statistics.median`，无新依赖。
- M2 验收脚本：对 spike 同素材（`output/b_multislide.pptx`）逐 Unit 调 `measure_coverage`，报**中位数 / p10 / p90**；中位数 ≥ 0.35 即通过。要求同时输出叠加图（把 rect 画到副本上，`output/spike/overlay_N.png` 已有先例），人工确认无"高亮空白"。
- 度量对象是**单个 `lines[i]`**（渲染实体），不是 `Unit.rect`；两者都报，差异过大说明该 Unit 的行宽异常。

### 4.6 非文本形状的 Unit

| kind | 处理 |
|---|---|
| `title` | `ParaInfo` 里的标题样式段落；`kind` 由"形状名含 Title 或 `is_placeholder` 且 `ph type ∈ {TITLE, CENTER_TITLE}` 或字号 ≥ 1.3× 本页正文中位字号"三条任一命中判定，全部不中则记 `body` |
| `picture` | 一个 Unit，`rect` = 形状外框（图片本身就是墨迹，coverage 天然接近 1），`lines = [rect]` |
| `chart` | 一个 Unit，`rect` = 形状外框（R4：**不拆数据点**），`text` = 空串或图表标题 |
| `table` | 单元格级：每个非空单元格的每个段落各一个 Unit，`kind="cell"`，`text` 为单元格文本；x/y 由 `table_col_widths_emu` / `table_row_heights_emu` 前缀和 + 单元格 `margin_*_emu` 推出。合并单元格（`is_merge_origin`/`is_spanned`）本期**只对 origin 单元格出 Unit、被并格跳过**（列 §10） |
| `other` | 不出 Unit（已在下游过滤），但保留在 `shapes` 里供 redesign 参考 |

---

## 5. 产物目录与 `deck.json`

```
<out>/                       # import 的 --out
├── deck.json                # DeckIR（源形状 + 页序 + 画布），**唯一真源**
├── bg/slide_1.png …         # COM 底图（faithful）或 None（--no-com / redesign）
├── units.json               # build_units 的快照（QA / 调试用，非真源）
└── player/index.html        # animate 产出
```

**决策：`deck.json` 只存 `shapes`，不存 `units`。** 理由：`units` 是 `shapes` + `export_width_px` 的纯函数结果，存两份必然漂移；`animate`/`video` 现场用 `build_units` 重算（纯算术，无 IO、无 COM）。`units.json` 是给 M2 验收与人工排查用的**快照**，程序不读它。

`bg` 一律写**相对 `<out>` 的 POSIX 风格相对路径**（`"bg/slide_1.png"`），播放器与 Flask 两条路径都能直接用；绝对路径会让 Flask 静态路由与 `file://` 打架。

**页数一致性**：`len(pages) == len(bg 文件)`；faithful 模式下不符即 `PptxError("IR_MISMATCH", …)`。

---

## 6. `hl_anim.py` —— 高亮播放器

### 6.1 接口

```python
def build_player(out_dir: str,
                 bg_paths: list[str],              # 相对 out_dir，如 "bg/slide_1.png"
                 pages_units: list[list[Unit]],
                 title: str = "",
                 dim: float = 0.25,                # 降暗比例，0<dim<0.6
                 auto_step_ms: int = 2000,
                 canvas_width_px: int = 1920,
                 canvas_height_px: int = 1080) -> str:
    """写 <out_dir>/index.html，返回其路径。零截图、零 iframe、零外部依赖。"""
```

参数校验：`len(bg_paths) != len(pages_units)` → `PptxError("IR_MISMATCH", "底图数与讲解单元页数不一致：…")`；`dim` 越界 → `PptxError("BAD_ARGS", "降暗比例需在 (0, 0.6) 之间")`。

### 6.2 结构：**不用 iframe**

与 `anim.py` 的关键差异：`anim.py` 必须 iframe——它要把**设计稿 HTML** 连脚本一起加载并注入动画引擎。高亮播放器的"幻灯片"是**一张 PNG** + 绝对定位的 div，**同源限制在这里毫无收益**，所以：

- 推荐**零 iframe**：`<div class="stage"><img class="bg" src="bg/slide_1.png"><div class="hl" …></div>…</div>`。攻击面比 anim.py 更小。
- 若实现者出于别的原因引入 iframe（如 redesign 模式预览），**必须** `sandbox="allow-scripts"` 且**禁止**加 `allow-same-origin`（KNOWLEDGE.md 已知风险表 + 计划硬约束 2）。这是不可协商的红线，审计 agent 会查。

### 6.3 转义纪律（照抄 `anim.py` 的三上下文 + 单遍替换）

沿用 `anim.build_player` 的做法，逐项对齐：

| 注入点 | 上下文 | 处理 |
|---|---|---|
| `title` | HTML 文本 **且** `<title>` | `html.escape(title)` |
| 每个 `bg_path` | `src` **属性** | `html.escape(p, quote=True)`（另：路径本身须经 §8 的"必须位于 out_dir 内"校验） |
| `CFG` 对象 | `<script>` **字符串** | `json.dumps(…, ensure_ascii=False)` 后 `.replace("</", "<\\/").replace("<!--", "<\\!--")` |
| 模板占位符 | — | **单遍** `re.sub(r"__TITLE__\|__BGJSON__\|__UNITS__\|__CONFIG__", lambda m: mapping[m.group(0)], TEMPLATE)`。**禁止链式 `str.replace`**——后一次 replace 会扫到前一次替换插入的文本（`anim.py` 审计 M2 的原话） |

数值参数（`dim`/`autoStepMs`/`canvas_width_px`）先做 `float()`/`int()` 强转 + 范围校验，再 `json.dumps`——不进字符串拼接。

### 6.4 交互与供程序驱动的 JS API（`video` 步骤模式依赖它）

- 步进/回退：`Space`/`→` 下一步、`←` 上一步；点击画面下一步；`revealAll()` 显示本页全部；页点导航。
- 自动讲解：`setInterval(stepFwd, CFG.autoStepMs)`，同 `anim.py` 的开关语义。
- 视觉：当前 Unit 的 rect 高亮（`border`/`background` 可辨识即可），其余 Unit 所在区域降暗 `dim`（实现：全屏一层 rgba 遮罩 + 当前 rect 挖洞，或对每个 rect 设 `opacity`；**推荐遮罩+挖洞**，因为底图是整张 PNG、无法按形状降暗）。
- 16:9 等比缩放：舞台固定 `canvas_width_px × canvas_height_px`，`transform-origin: 0 0; transform: scale(stage.clientWidth / canvas_width_px)`，同 `anim.py` 的 `fit()`。

```js
window.hl = {
  goto(page, step),   // 切到第 page 页并揭示前 step 个单元；**同步改 DOM，不用 rAF/过渡**
  next(), back(),
  state(),            // -> {page, step, totalPages, totalUnits}
  ready               // Promise<void>，DOM 就绪后 resolve
};
```

`goto` 必须**同步生效**（无 CSS 过渡/无 `requestAnimationFrame` 依赖）——截图侧只靠 `screenshot(animations="disabled")` 兜底，不要再引入一个等待源。

### 6.5 截图序列（供 `pptgen video --mode step`）

`synthesize` 复用 `video.py` 后，截图这一环是唯一缺口：`shot.shot_deck` 硬绑 `.slide` 选择器与设计稿结构，**不能用于播放器**。契约：

```python
def shot_player(player_html: str, out_dir: str,
                steps: list[tuple[int, int]],         # [(page, step), …] 按序
                min_settle_ms: int = 300,
                max_settle_ms: int = 2000) -> list[str]:
    """逐 (page, step) 驱动 window.hl.goto → 截图 → 写 step_0001.png…，返回路径列表。"""
```

必须沿用 `shot.py` 的三重确定化（KNOWLEDGE.md：**改 `shot.py` 时别退回固定 sleep**）：
① 打开后 `page.evaluate("window.hl.ready")`；② `screenshot(animations="disabled")`；③ 连续两帧字节一致才收工、上限 `max_settle_ms`。

> ⚠️ **需要碰现有文件的一处（请批准）**：`shot.py` 的 `_launch_browser` / `_screenshot_settled` 是私有名。直接跨模块 import 私有函数违背"共享底层"的意图，但把 15 行启动+轮询逻辑抄进 `hl_anim.py` 会造出**第二份**确定化实现——正是 D2 要避免的"真 bug 修两遍"。
> **建议**：在 `shot.py` 里加两个**纯新增**的公开薄函数（不改动任何现有行为，248 条测试零影响）：
> ```python
> def launch_browser(p):            return _launch_browser(p)
> def shot_page(page, min_ms=300, max_ms=2000): return _screenshot_settled(page, min_ms, max_ms)
> ```
> 若评审坚持 `shot.py` 一字不动，退路是 `hl_anim` 内 `from shot import _launch_browser, _screenshot_settled` 并在注释里写明"有意复用私有符号，见 PPTX_INTERFACE §6.5"。**二选一，不要复制粘贴第三份。**

---

## 7. `pptx_out.py` —— 母版 + 占位符导出

### 7.1 接口

```python
@dataclass
class TemplateInfo:
    path: str
    layout_names: list[str]
    # layout_idx -> [(ph_idx: int, ph_type: str, name: str), …]
    placeholders: dict[int, list[tuple[int, str, str]]]
    has_theme_cjk: bool        # 主题 a:ea 或 <a:font script="Hans"> 非空

def read_template(path: str) -> TemplateInfo:
    """枚举模板的 slide_layouts 与占位符；不可用抛 PptxError("TEMPLATE_INVALID")。"""

def build_builtin_master(out_path: str, title: str = "演示文稿") -> str:
    """生成内置 5 版式的空白母版（真占位符），返回 out_path。供用户在 PowerPoint 里手改。"""

def build_deck_pptx(deck: DeckIR, out_path: str,
                    template: str | None = None,
                    no_com: bool = False) -> str:
    """deck.json → 可编辑 pptx。template=None 用内置母版；否则读自定义模板。
    no_com=True：不支持降级的形状（SmartArt/复杂组合）跳过而非报错。"""
```

### 7.2 内置 5 版式（映射到 python-pptx 默认模板的真实 layout）

| 版式 | layout idx | 填充的占位符 idx |
|---|---|---|
| 封面 | `0` Title Slide | `0` CENTER_TITLE ← 标题；`1` SUBTITLE ← 副标题 |
| 目录 | `1` Title and Content | `0` TITLE；`1` OBJECT ← 目录条目（多段） |
| 内容 | `1` Title and Content | `0` TITLE；`1` OBJECT ← 要点（多段、可用 `level` 做缩进） |
| 数据 | `5` Title Only | `0` TITLE；图表/表格用 `shapes.add_chart` / `add_table` 落到正文区（OBJECT 占位符也可承载，见 7.3） |
| 尾页 | `0` Title Slide | `0` CENTER_TITLE ← 结束语 |

> `SlideLayout.placeholders` 在 `layout 6 (Blank)` 上只有 `DATE/FOOTER/SLIDE_NUMBER`，**不要用 6 做内容页**。

### 7.3 占位符类型映射

| 输入形状 | 占位符类型 | 填充方式 |
|---|---|---|
| title | `TITLE` / `CENTER_TITLE` | `slide.placeholders[0].text_frame` |
| body / bullet | `BODY` / `OBJECT` / `SUBTITLE` | 同上，逐段 `tf.add_paragraph()`，`level` 透传 |
| picture | `PICTURE` | `ph.insert_picture(abs_path)`；无 PICTURE 占位符时退化为 `slide.shapes.add_picture(...)` |
| table | 无对应占位符枚举 | `shapes.add_table(rows, cols, …)` 落在 OBJECT 占位符的矩形内；表格本身在 PowerPoint 里可编辑 |
| chart | 无对应占位符枚举 | `shapes.add_chart(XL_CHART_TYPE.X, …)`；类型映射复用 `builder._CHART_TYPES` 的语义（`bar/column/pie/line`，**不改 builder，只对齐取值**） |

**清空占位符文本**：用 `tf.clear()`（留下一个空段落），再逐段 `tf.paragraphs[0]` / `tf.add_paragraph()` 写入。**不要**用 `tf.text = "…"`——它只能造一个无字体控制的 run。

**占位符可能不存在**：自定义模板的 layout 未必有 `TITLE`/`OBJECT`（用户上传的稿子千奇百怪）。规则：按 `ph_type` 查找，找不到则退化为在**该 layout 的对应占位符矩形**（无则用安全区 `Inches(0.6), Inches(0.7), Inches(12.1), Inches(6.1)`）上 `add_textbox`，并在 stderr 记 `degraded_placeholder`。

### 7.4 中文字体：必须显式写 `a:ea`

实测（§1）：python-pptx 默认模板的 `theme1.xml` 里 `a:ea=""`、`<a:font script="Hans" typeface="宋体"/>`。`run.font.name` **只写 `a:latin`**，中文属东亚文字范围、按 `a:ea` 匹配 → 不做处理时演示稿里的中文是**宋体**，而不是 `builder.py` 用的微软雅黑。

| 母版来源 | 策略 |
|---|---|
| **内置母版** | 双管齐下：① 每个写入的 run 用 `builder._set_ea` 同款逻辑补 `a:ea`（**必须 append 在 `a:latin` 之后**，OOXML 元素序要求）；② `prs.save()` 后**后处理 zip**：把 `ppt/theme/theme1.xml` 的 `<a:ea typeface=""/>` → `typeface="微软雅黑"`、`<a:font script="Hans" typeface="宋体"/>` → `typeface="微软雅黑"`。**② 是必要的**——用户在 PowerPoint 里新敲的字走主题字体，只改 run 救不了后续编辑 |
| **自定义模板** | **不覆盖字体**（只填文本、不设 `font.name`），尊重品牌模板的字体意图；仅在模板 `has_theme_cjk is False`（既无 `a:ea` 也无 `script="Hans"`）时补写 run 级 `a:ea`，并 stderr 提示 |

zip 后处理的实现约束：`zipfile.ZipFile(src)` 读 → 新建临时 zip 全量复制（除 `theme1.xml`）→ `os.replace` 到目标。**必须保留 `[Content_Types].xml` 与其余全部条目**，不得重排压缩包结构（PowerPoint 对条目顺序不敏感，但对缺失条目零容忍）。

### 7.5 超长文本：**先缩字号，再截断**

```python
def fit_text(text: str, size_pt: float, box_width_pt: float, box_height_pt: float,
             min_size_pt: float = 12.0) -> tuple[float, str, bool]:
    """返回 (最终字号, 最终文本, 是否被截断)。
    用 qa.measure_text_lines 试算：逐档下调字号（步长 1pt，下限 min_size_pt）
    直到 needed = Σ 行数×字号×LINE_HEIGHT_FACTOR + 段距 ≤ box_height_pt；
    仍放不下则按行宽二分截断并追加 "…"，返回 truncated=True。"""
```

**推荐缩字号优先**的理由：截断会**静默丢内容**——CLI 的调用方是 agent，`{ok:true}` 里少了一句话它无从察觉；缩字号保住信息，只在到 12pt 仍溢出（极端输入）时才截断，且 `truncated` 必须冒泡到 CLI 的 stderr 与 JSON 的 `warnings`。**不支持** `MSO_AUTO_SIZE`/`normAutofit` 自动缩（Python 侧算不出 PowerPoint 的渲染结果，且会让"导出后文字突然变小"变成不可预期的动作）。

---

## 8. `cli.py` —— `pptgen` 入口（方向 B）

### 8.1 进程契约

- **stdout 恒为一行 JSON**（机器可读，`ensure_ascii=False`）；**人类可读进度/警告走 stderr**。理由：调用方是 bash 里的 agent，`json.loads(stdout)` 必须永远成立，日志混入即解析失败。
- **永不上抛裸堆栈**：所有 `PptxError` / `ValueError` / `RuntimeError` / 未捕获异常统一收敛成 `{"ok": false, "error": {code, message, hint}}` + 非 0 退出码。
- 中文错误必须带**可执行的修复指引**（`hint`），例如"未检测到 PowerPoint…请安装 Microsoft Office，或改用 --mode redesign"。
- 退出码：`0` 成功｜`2` 参数错误（argparse 原生）｜`3` 输入问题（`PPTX_*`/`IR_MISMATCH`/`TEMPLATE_INVALID`）｜`4` 缺外部依赖（`NO_POWERPOINT`/`NO_FFMPEG`/`NO_BROWSER`）｜`5` 内部错误。

```json
{"ok": true, "cmd": "import", "data": {"out": "...", "pages": 10, "units": 41, "skipped": 18, "warnings": []}}
{"ok": false, "cmd": "import", "error": {"code": "NO_POWERPOINT", "message": "…", "hint": "…"}}
```

### 8.2 子命令与参数

| 命令 | 必填 | 可选 | 关键行为 |
|---|---|---|---|
| `import <pptx>` | `--out DIR` | `--mode faithful\|redesign`（默认 `faithful`）、`--width {1280,1920,2560}`（默认 1920）、`--pages N`（只导前 N 页）、`--no-com` | 写 `<out>/{deck.json,bg/,units.json}`；`--mode redesign` 额外跑 `outline`/`style`/`html_gen` 流水线 |
| `animate <dir>` | — | `--dim 0.25`、`--auto-ms 2000`、`--highlight`（默认开，仅作显式化） | 读 `<dir>/deck.json` → `hl_anim.build_player` → `<dir>/player/index.html` |
| `video <dir>` | — | `--sec 4`、`--fps 25`、`--fade 0.8`、`--mode page\|step`（默认 `page`）、`-o out.mp4` | `page`＝底图直出（`video_mod.synthesize`）；`step`＝`hl_anim.shot_player` 截步进序列再合成 |
| `export <deck.json>` | `-o out.pptx` | `--template brand.pptx`、`--no-com` | `pptx_out.build_deck_pptx`；导出后自动跑 `qa.check_pptx`，把 `errors/warnings` 放进 JSON |
| `deck "<主题>"`（可选） | `--out DIR` | `--template`、`--mode` | 复用 `outline`→`style`→`html_gen` 现有流水线（`llm_util.images_enabled()` 生效） |

**参数语义的硬边界（不许含糊）**：

- `--no-com` + `--mode faithful` = **非法组合**，报 `BAD_ARGS`：`"faithful 模式必须有 PowerPoint 底图；请去掉 --no-com，或改用 --mode redesign"`。faithful 的视觉保真**完全依赖** COM 底图（D8/D9），没有底图的高亮播放器是无本之木。
- `--no-com` 的合法用途：`import --mode redesign`（不要视觉参考图）、`export`（不支持的形状跳过而非报错）。
- `--pages N` 只截断页数，不改画布；`deck.json` 的 `pages` 与 `bg/` 同步截断。
- `--width` 只在 `import` 生效并**固化进 `deck.json`**；`animate`/`video` 不再接受 `--width`（坐标已按该宽度算好 px，改宽度必须重导）。
- `-o` 省略时默认 `<dir>/out.mp4`（`video`）/ `<dir>/<stem>.pptx`（`export`）。
- 默认 `--out`：`<项目根>/output/pptx_src/<pptx 主名>/`（沿用 `app.py` 的绝对路径基址风格），失败时 stderr 打印实际使用的绝对路径。

### 8.3 `--template` 与 `--out` 的路径安全

- `--template` 必须：`os.path.isfile` 且后缀 `.pptx` 且**解析后（`os.path.realpath`）位于允许根**（仓库内 `output/templates/` 或用户显式给出的绝对路径——后者需 `--allow-external-template` 之外的**隐式**放行，因为 CLI 是本地工具，读用户自己的文件不是越权）。**真正的红线是"不把上传/传参路径拼进输出路径"**：所有输出路径一律 `os.path.basename()` 后拼到 `<out>` 下，`..`/绝对路径一律拒绝。
- `deck.json` 里的 `bg` 路径必须校验为 `<out>` 内的相对路径（`os.path.normpath` 后不以 `..` 开头、非绝对），否则 `IR_MISMATCH` —— 防"被改过的 deck.json 指向任意文件"。
- 审核项（交给审计 agent）：pptx 上传的 zip bomb / 恶意 XML、`--template` 的路径穿越、CLI 参数注入。

### 8.4 skill 侧（D14/D15，本次只定形状）

`skill/ppt-anim/scripts/pptgen.py` 是**薄封装**：定位项目 venv 的 python，`os.exec` 转发到 `cli.py`，**不复制任何逻辑**。SKILL.md 的 frontmatter 按 agentskills.io 标准写，正文写"何时用/怎么调/产物在哪/常见错误码含义"——错误码直接引用 §9 表。

---

## 9. 错误码总表（CLI 与 PptxError 共用）

| code | 退出码 | message 示例 | hint |
|---|---|---|---|
| `BAD_ARGS` | 2 | `--no-com 与 --mode faithful 不能同时使用` | `faithful 模式需要 PowerPoint 导出底图；请去掉 --no-com，或用 --mode redesign` |
| `PPTX_NOT_FOUND` | 3 | `找不到文件：xxx.pptx` | 检查路径是否正确（建议用绝对路径） |
| `PPTX_UNREADABLE` | 3 | `PPTX 无法打开：xxx.pptx（<原因>）` | 确认是 .pptx（非 .ppt/.pdf），且未被其他程序占用 |
| `PPTX_ENCRYPTED` | 3 | `PPTX 已加密，无法读取` | 请先用 PowerPoint 去掉打开密码再导出 |
| `PPTX_EMPTY` | 3 | `PPTX 中没有任何幻灯片` | — |
| `IR_MISMATCH` | 3 | `底图数与讲解单元页数不一致：10 vs 9` | 重新执行 pptgen import 生成 deck.json |
| `TEMPLATE_INVALID` | 3 | `模板不可用：xxx.pptx（<原因>）` | 模板需为标准 .pptx，且至少含一个版式 |
| `NO_POWERPOINT` | 4 | `未检测到 PowerPoint，无法导出保真底图` | 安装 Microsoft Office（含 PowerPoint）；或改用 --mode redesign |
| `COM_EXPORT_FAILED` | 5 | `第 7 页底图导出失败（已完成 6/10 页）` | 检查 PowerPoint 是否被其他程序占用/是否弹出了对话框 |
| `NO_FFMPEG` | 4 | `未找到 ffmpeg` | `winget install ffmpeg`（沿用 `video.py` 的原文案） |
| `NO_BROWSER` | 4 | `未找到可用的 Chrome/Edge` | 安装 Google Chrome，或 `playwright install chromium`（沿用 `shot._launch_browser` 的原文案） |
| `INTERNAL` | 5 | `内部错误：<异常摘要>` | 附 stderr 完整堆栈；这属于 bug，请附命令与 deck.json |

**文案纪律**：`message` 说"发生了什么"，`hint` 说"下一步做什么"。二者都不出现 Python 异常类型名与英文堆栈（堆栈只写 stderr，供人 debug）。

---

## 10. 待实现阶段验证清单（不确定项，**必须在对应里程碑开工前实测**）

| # | 不确定点 | 影响 | 验证方法 |
|---|---|---|---|
| V1 | **组合形状坐标换算**（§2.2 公式） | 组合内的高亮会整体错位 | 造一份手工含组合的 pptx（R3 素材），把换算后的 rect 叠加到底图上（`measure_coverage` 同一套 overlay），人眼看是否贴合；对比"不换算"的错位程度 |
| V2 | `GroupShape.left is None` 的真实触发场景 | 整组丢失 | 同上素材，统计 `xfrm` 缺失的组占比 |
| V3 | **图表能否拆数据点**（R4 已定不拆） | 高亮粒度 | 读 `graphic_frame.chart.plots[0].series[0].values` 是否可用；若可用，评估"数据点级高亮"的收益，**结论写回本文而非直接实现** |
| V4 | `vertical_anchor is None` 的**继承**语义 | 内容页文字整体偏移 | 用真实稿 + COM 底图逐页对比：None 时按 TOP 处理，看墨迹是否压在 rect 顶部；body 占位符可能从 layout 继承 MIDDLE |
| V5 | **首段 `space_before` 是否计入** | 首个 rect 偏移 20pt（`builder.py` 高频 `space_before=Pt(20)`） | 同上对比；PowerPoint 对文本框首段是否渲染段前距有兼容性差异 |
| V6 | `line_spacing` 为 `Length`（绝对值）时的折算 | 行高错 | 造一页显式行距的稿，量底图行间距 |
| V7 | `a:normAutofit/@fontScale` 的真实取值 | rect 偏大会重现"高亮空白" | 造一页触发自动缩排的稿，读 XML 的 `fontScale` 与底图实际字高对比 |
| V8 | `app.Visible` 取值 / `DispatchEx` 是否真隔离用户实例（§3.4-2/3） | **可能关掉用户没保存的 PowerPoint** | 先手动开一份未保存的稿 → 跑 `export_pages` → 确认该稿仍在；同时确认 `WithWindow=False` 下 `Export` 可用 |
| V9 | 合并单元格（`is_merge_origin`/`is_spanned`）的几何 | 表格内高亮错位 | 造含合并单元格的表，比对 rect 与底图 |
| V10 | 真实稿的字体替换（R6，微软雅黑→宋体） | 行宽失真 | 手工稿若用非雅黑字体，`_char_width_pt` 仍按 msyh 度量；量化误差后决定是否要"按 `a:latin`/`a:ea` 选字体文件"（本期先不扩） |
| V11 | 旋转形状（`rotation != 0`） | rect 不跟随 | 量"不旋转 + CSS rotate 近似"与底图的偏差；决定是否列为已知限制 |
| V12 | spike 素材片面（R3，`pictures=0 tables=0` 无 SmartArt） | M2 结论只在规整稿成立 | **用户须提供一份手工真实 PPT**（计划 §10 开放项 1）；缺失则 M2 验收范围必须显式声明为"仅 builder 规整稿" |
| V13 | `output/templates/` 命名冲突 | 自定义 **pptx 模板**与 `template.py` 已有的 **JSON 风格模板**（`CUSTOM_DIR`，实测已有 `d24ceec9.json`）同目录 | 建议自定义 pptx 模板改放 `output/templates/pptx/`，**不下发到 `output/templates/` 根**，避免 `template.custom_templates()` 把 `.pptx` 当 JSON 解析 |

---

## 11. 给实现者的清单

### 11.1 实现顺序（严格串行，前一步的验收不过不进下一步）

| 步 | 模块 / 里程碑 | 要实现的函数 | 验收命令 |
|---|---|---|---|
| 1 | **`pptx_io.py`**（M1） | `PptxError`、`powerpoint_available()`、`read_pages()`、`export_pages()` | 对 `output/b_multislide.pptx`：读出 **10 页 / 40 形状**、坐标非 `None`；COM 导出 10 张 1920×1080 PNG，`≤1.5s/页` 量级 |
| 2 | **`hl_layout.py` 的断行层**（M2 前半） | `Line`、`wrap_lines()` + **与 `qa.measure_text_lines` 的等价测试** | 等价测试语料全绿（纯中文/纯英文/中英混排/超长词/全角标点/首尾空格/单字超宽/空串） |
| 3 | **`hl_layout.py` 的定位层**（M2 后半） | `Rect`、`Unit`、`build_units()`、`measure_coverage()` | 同素材 coverage **中位数 ≥ 0.35**、框溢出率 ≤ 5%，**并产出 overlay 图人工确认** |
| 4 | **`hl_anim.py` 的播放器**（M3） | `build_player()` | 浏览器打开：逐条高亮定位准、其余降暗、步进/回退/自动可用；**页数 == 底图数**；转义用例（`title` 含 `</title><script>`、`__CONFIG__`、bg 路径含 `#/%`） |
| 5 | **`hl_anim.py` 的截图**（M3 尾） | `shot_player()` +（按 §6.5 决定）`shot.py` 的两个公开薄函数 | 步进序列截图数 == Σ units，画面非空白 |
| 6 | **`cli.py`**（M4） | 5 个子命令 + JSON 信封 + 退出码 + 错误码表 | `import/animate/video/export` 四命令跑通；**无 PowerPoint 的机器上返回明确中文错误而非堆栈**（可用 `powerpoint_available` 打桩模拟） |
| 7 | **`pptx_out.py`**（M6） | `read_template()`、`build_builtin_master()`、`build_deck_pptx()`、`fit_text()` + theme zip 后处理 | 导出的 .pptx 在 PowerPoint 打开**可编辑**：改文字不破版、占位符可选中；中文渲染为**微软雅黑而非宋体**；`qa.check_pptx` 无 error |
| 8 | 工作台入口（M5）/ skill（M7） | `app.py` +2 路由、`index.html` +1 按钮；`skill/ppt-anim/` | 上传 pptx → 浏览器可见播放器；pi 里按 SKILL.md 跑通一次全流程 |

**回归红线**：每步结束跑 `.venv/Scripts/python.exe -m pytest tests/ -q`，**248 条保持全绿**；新增用例约 20–25 条。`anim.py` / `builder.py` **一字不改**。

### 11.2 必须先验证再动笔的点

按"会写错就白写"排序，**开工前**必须实测：

1. **V12 真实稿素材**（R3）—— 没有它，第 3 步的 coverage 结论只在 `builder.py` 的规整稿上成立。**这是唯一需要用户提供输入的阻塞项**，请最先索取。
2. **V1/V2 组合形状换算 + `left is None`** —— 决定 `hl_layout` 与 `read_pages` 的核心递归结构，写错了整个 M2 要返工。
3. **V8 COM 实例隔离（`DispatchEx`）** —— 写错会**关掉用户没保存的 PowerPoint**，是本次唯一有"破坏用户数据"后果的点。
4. **V4/V5 垂直锚点继承 + 首段段前距** —— 直接决定高亮框垂直位置，看底图一眼可判。
5. **V7 `fontScale`** —— 决定 rect 是否再次"高亮一大片空白"（正是 M2 要消灭的现象）。
6. **V13 `output/templates/` 命名冲突** —— 一个目录决策，写进代码前定好，改起来要动 CLI 与文档两处。

### 11.3 交付物自查（Loop 工程：模型不能批准自己的完成）

- [ ] 只读了代码，**未修改任何 `.py`**（本任务唯一产物是 `docs/PPTX_INTERFACE.md`）
- [ ] 每个引用的现有符号都能在 §1 表里找到出处（文件 + 符号名）
- [ ] 每个数据结构都写了字段名与单位（EMU / pt / px 三者标清）
- [ ] §10 的 13 条不确定项都在文里被显式标注，未混入"确定"叙述
- [ ] §11.1 的顺序可执行：每步都有**可复现的验收命令**，不是"看起来对"
