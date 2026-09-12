# pptx 双向链路 · 接口契约（v2）

> 版本 **v2**（2026-09-12）｜v1 基线 commit `c188b63`｜v2 依据 `docs/IMPL_REPORT.md`（D1–D9）、`docs/TEST_REPORT.md`（F1–F6）、`docs/AUDIT_REPORT.md`（S1 / M1–M4 / L1–L7 / N1–N4）
> 上游：[PLAN_PPTX_ANIM.md](../PLAN_PPTX_ANIM.md) 的 D1–D16 / M1–M7
> 本文只定**契约**，不含实现。所有"现有符号"均为**实际读到的**（见 §1），无猜测。
> **v2 的立场**：v1 里被实测/审计推翻的**前提**一律改写，不保留"原样但加注释"；每条修订的可追溯记录（三段式）在 **§12**。

---

## 0. v2 变更摘要 —— 被推翻的 16 条前提

| # | v1 的前提 | 实测/审计结论 | 处置（详见 §12） |
|---|---|---|---|
| 1 | `DispatchEx` 强制新实例，可与用户的 PowerPoint 隔离 | **不成立**：`Dispatch`/`DispatchEx`/`GetActiveObject` 拿到**同一个实例**（同进程、`Presentations.Count` 同步） | §3.4 改为"实例是共享的，安全只靠守卫"，并补 `had_powerpoint` + PID 归属双守卫 | R-D1 |
| 2 | 行级高亮 coverage 中位数 ≥ **0.35** 可验收 | **物理不可达**：实测 0.286；矩形法上限 0.414（CJK 字形墨迹只占 em 盒 41%）；pad 归零也仅 0.321 | §4.5 改为可达表述：相对 ≥2.5× **或** pad 归零 ≥0.30，并写明上限 | R-D2 |
| 3 | 图片/图表单元"coverage 天然接近 1" | **不成立**：图表单元实测 **0.166**（全场最低） | §4.6 删除该断言，验收对 chart/picture/table 单列口径 | R-D8 |
| 4 | `font_scale` 存百分数、用时直接相乘 | **口径不自洽**：`60000`→存 60.0，直接相乘得 1200pt 而非 12pt | §2.3 改为存**真实倍率**（`÷100000`→0.6），§4.3 才可直接相乘 | R-D3 |
| 5 | 标题判定第 ② 条（`ph type ∈ {TITLE, CENTER_TITLE}`）可从 IR 实现 | **不可实现**：`ShapeInfo` 只有 `is_placeholder: bool`，无 ph type | §2.3 增 `ph_type` 字段；§4.6 第 ② 条改为可落地 | R-D4 |
| 6 | 合并单元格 origin 的 rect 按单列宽算 | **偏窄**：实测合并 origin 按单列宽断出 2 行，实际渲染 1 行 | §2.3 增 `table_spans`；§4.6 rect 按跨列/跨行求和 | R-D5 |
| 7 | `DeckIR.pages` 可直接喂 `build_units` | **缺桥**：`pages` 不带画布尺寸，`build_units` 却要画布宽才能做 EMU→px | §2.3/§5.1 正式定义 `page_shapes(page, deck)` 桥接函数 | R-D6 |
| 8 | `Length` 形态行距折成倍数后仍乘 `LINE_HEIGHT_FACTOR` | **偏差 +19%**：16pt 字号 + 固定 30pt 行距，契约算 37.5pt，实测渲染 31.5pt | §2.3 拆成 `line_spacing_pt` / `line_spacing_mult`，绝对值**不经** 1.25 | R-D7 |
| 9 | 首段 `space_before` 计入 | **PowerPoint 忽略**：实测首行墨迹距框顶 5.9pt ≈ margin_top | §4.3 改为"仅非首段计入" | R-V5 |
| 10 | 验收写"读出 10 页 / **40 形状**" | 裸 XML 实为 **41** 个顶层形状（40 `p:sp` + 1 chart），过滤后 **23** | §11.1 改写验收数字并说明口径 | R-D9 |
| 11 | COM 导出 "≤1.5s/页" | 仅**稳态**成立（逐页中位 0.125–0.146s）；**冷调用端到端 2.64s/页**，且不摊销 | §11.1 写明"稳态"口径，冷调用成本单列 | R-F3 |
| 12 | 自定义 pptx 模板放 `output/templates/` 会"被当 JSON 解析崩掉" | 实际是**静默跳过**（`custom_templates()` 只认 `.json` 且逐文件吞异常） | §7.1 保留"另放 `pptx/`"的结论，但改掉理由 | R-V13 |
| 13 | `PageShapes.shapes` 是"**已展开**组合" | 实测 `read_pages` **保留 group 节点**（children 嵌套） | §2.3 措辞改为"组合以 `children` 嵌套表示" | R-N1 |
| 14 | §4.2 的等价断言足以防"两套算法漂移" | **盖不住**：真正产出几何的是 `_paragraph_lines`，它在 `a:br` 上与 QA 差 +1 行 | §4.2 断言改测 `_paragraph_lines`，并把该分歧写成**显式契约** | R-M1 |
| 15 | `prez.Close()` 与 `Quit()` 同样安全（都吞异常） | **Close 一道守卫都没有**：我们附着在用户实例上，用户开着同一份稿时会被关掉 | §3.4 新增"只关自己新开的"规则 | R-S1 |
| 16 | `word_wrap` 语义无关紧要 | `python-pptx` 的 `add_textbox` **默认写 `wrap="none"`**（该素材 13/22 个框）→ M2 数值上等价于段级，折行层**未被验收素材覆盖** | §4.5 增加**作用域声明**；M2 结论仅对单行段落成立 | R-F2 |

**另有三条严重度提醒（同样在 §12）**：`a:br` 与 QA 的值分歧（R-M1）、zip bomb 无闸门（R-M2）、`export_pages` 拆掉调用方 COM apartment（R-F1）。

---

## 1. 可复用的现有能力（已读，勿重复造）

| 出处 | 符号 | 签名 / 值 | 本计划用途 |
|---|---|---|---|
| `qa.py` | `EMU_PER_PT` | `12700` | 所有 EMU↔pt 换算 |
| `qa.py` | `LINE_HEIGHT_FACTOR` | `1.25` | 行高 = 字号 × 1.25（**仅倍数行距时用**，见 §4.3） |
| `qa.py` | `DEFAULT_FONT_SIZE_PT` | `18.0` | 字号缺失回退值（**与溢出 QA 同源**） |
| `qa.py` | `OVERFLOW_TOLERANCE_PT` / `BOUNDARY_TOLERANCE_PT` | `2.0` / `0.5` | 框溢出、越界容差 |
| `qa.py` | `_load_font()` | `-> (cmap, hmtx, unitsPerEm) \| None`，`lru_cache(maxsize=1)` | 字体可用性判定 |
| `qa.py` | `_char_width_pt(ch, size_pt)` | `-> float`（pt），字体缺失时降级 空格0.33x / ASCII0.55x / 其他1.0x | **单字符宽度，行级定位的核心** |
| `qa.py` | `_tokenize(text)` | `-> list[(kind, token)]`，kind ∈ `word`/`cjk`/`space` | 断行策略（拉丁整词、CJK 逐字、空格可断）。⚠️ **`"\n"` 落进 `cjk` 分支，不是硬换行**（见 R-M1） |
| `qa.py` | `_measure_lines_ex(text, size_pt, box_width_pt)` | `-> (行数, used_estimate: bool)` | 溢出 QA 行数（**只给行数，不给断点**） |
| `qa.py` | `measure_text_lines(...)` | `-> int` 公开版 | `pptx_out` 缩字号试算、等价断言 |
| `qa.py` | `font_available()` | `-> bool` | `used_estimate` 标注 |
| `qa.py` | `_para_font_size(paragraph)` | 逐 run 取第一个显式字号，否则 `18.0` | **必须逐字复刻该规则**，否则高亮与 QA 打架 |
| `qa.py` | `check_pptx(path, expected_pages)` | `-> {"errors","warnings","used_estimate"}` | M6 导出后自检 |
| `video.py` | `ffmpeg_path()` / `ffprobe_path()` | `-> str \| None`（PATH→WinGet Links→WinGet Packages） | `pptgen video` 前置探测 |
| `video.py` | `build_page_args(img, out, seconds=4.0, fps=25, size=(1280,720))` | `-> list[str]` 纯函数 | 单页片段 |
| `video.py` | `build_final_args(page_clips, out, seconds=4.0, fade=0.8, fps=25)` | `-> list[str]`；`n<2` 抛 `ValueError` | xfade 拼合 |
| `video.py` | `synthesize(shots, out_path, seconds=4.0, fade=0.8, progress_cb=None)` | `-> str`；`fade` 必须 `0<fade<seconds`；自动建/清 `out_path+"_parts"` | **`pptgen video` 直接复用，零改动** |
| `shot.py` | `shot_deck(html_path, out_dir, min_settle_ms=300, max_settle_ms=2000)` | `-> list[str]`；要求 `.slide`，否则 `ValueError` | 只服务**设计稿 HTML**，**不适用**高亮播放器（§6.5） |
| `shot.py` | `_launch_browser(p)` / `_screenshot_settled(page, min_ms, max_ms)` | 私有；三重确定化 | 截图必需，复用方式见 §6.5 |
| `anim.py` | `build_player(out_dir, deck_web_path, title="", auto_step_ms=1200)` | `-> str`；三上下文转义 + **单遍替换** | 转义纪律的范例，照抄 |
| `builder.py` | `_set_ea(run, font="微软雅黑")` | 补写 `a:ea`，**必须排在 `a:latin` 之后** | `pptx_out` 中文字体（§7.4） |
| `builder.py` | `SLIDE_W`/`SLIDE_H`/`FONT`/`THEMES` / `_CHART_TYPES` | `Inches(13.333)`/`Inches(7.5)`/`"微软雅黑"`/三套色板/图表选型 | 版式尺寸与图表选型基准（只对齐取值，不改 builder） |
| `builder.py` | `build_ppt(slides, image_paths, out_path, theme="blue", subtitle="")` | `-> str` | 保留不动（D11 双导出） |
| `llm_util.py` | `llm_client(timeout=60.0)` / `get_model(kind)` / `images_enabled()` | 无 key 抛 `RuntimeError("未配置 ZHIPUAI_API_KEY")` | redesign / deck 命令 |
| `app.py` | `_video_jobs: set[str]` + `threading.Lock` | 同稿并发守卫范例 | CLI 不共享该 state（D2） |
| `app.py` | `OUTPUT_DIR` … `VIDEOS_DIR` | 均为绝对路径 | CLI 默认输出根 |
| `template.py` | `_custom_path(key)` / `custom_templates()` | `custom:<8位hex>` → `CUSTOM_DIR/<id>.json`；**只认 `.json`**，逐文件吞异常 | `--template` 路径校验范式；目录冲突见 R-V13 |

### 1.2 实测事实（v2 新增/更正；v1 的对应条目已按此改写）

| 事实 | 实测值 | 来源 |
|---|---|---|
| `TextFrame` 默认内边距 | `margin_left/right = 91440 EMU`（7.2pt）、`margin_top/bottom = 45720 EMU`（3.6pt） | 现场内省 |
| `BaseShape.left/width` | `return self._element.x / .cx` —— **组合内子形状返回子坐标系原值，不做变换** | 现场内省 |
| `GroupShape.left` 可为 `None` | `CT_GroupShape._get_xfrm_attr`：`xfrm is None → return None`；**实测删 xfrm 后四属性全 `None`** | 内省 + `docs/IMPL_REPORT.md` 开工前验证 V1/V2 ✅ |
| 组合子坐标空间 | `CT_GroupShape` 有 `chOff`/`chExt`；仿射换算方向**已实测正确** | V1 ✅ |
| ⚠️ **嵌套组合的坑** | 子映射必须吃**本层坐标原值**，由 compose 链送到幻灯片坐标；直接把"绝对坐标"喂进子级会**重复施加外层缩放**（实测 leaf 落 3in vs 期望 1in） | `docs/IMPL_REPORT.md` 步 1 真 bug 2 + `test_nested_group_mapping_is_cumulative` |
| 隐藏标记 | `<p:cNvPr hidden="1">`；python-pptx **无公开 `hidden` 属性**，走 `shape._element.nvSpPr.cNvPr.get("hidden")` | 现场内省 |
| `TableCell` | 有 `margin_left/top/right/bottom`（91440/45720）、`is_merge_origin`、`is_spanned`、`span_width`、`span_height`；`is_spanned=True` 实测可用 | 内省 + V9 ✅ |
| `paragraph.line_spacing` | 可能是 `float`（倍数）或 `Length`（**`int` 子类**，先判数值会把 `Pt(30)` 当成 381000 倍行距） | `docs/IMPL_REPORT.md` 步 1 真 bug 1 + `test_line_spacing_multiple_and_length` |
| ⚠️ **`add_textbox` 默认 `wrap="none"`** | 即 `word_wrap=False`；该素材 **13/22** 个文本框如此 → 每段恒 1 行 | `docs/TEST_REPORT.md` F2 / `tools/probes/indep_wordwrap.py` |
| ⚠️ **`a:br` 在仓库内素材全为 0** | 44 / 25 / 51 段全部 `a:br=0 a:fld=0`；但手工汇报稿里很常见 | `docs/AUDIT_REPORT.md` M1 / `tools/probes/audit_pptx_io.py` §7 |
| `prs.slide_width` 可为 `None` | 缺 `p:sldSz` 的包 → `int(None)` 抛裸 `TypeError` | `docs/AUDIT_REPORT.md` M4 |
| 默认模板 11 个 layout | `0 Title Slide`(idx0 CENTER_TITLE, idx1 SUBTITLE)、`1 Title and Content`(idx0 TITLE, idx1 OBJECT)、`5 Title Only`、`6 Blank`、`8 Picture with Caption`(idx1 PICTURE)… | 现场内省 |
| **默认模板主题无 CJK 字体** | `theme1.xml`：`a:ea=""`（空）且 `<a:font script="Hans" typeface="宋体"/>` → 中文渲染为**宋体** | 现场内省 |
| **COM 实例是共享的** | `Dispatch`/`DispatchEx`/`GetActiveObject` → 同进程、`Presentations.Count` 同步 | `docs/IMPL_REPORT.md` D1 / `tools/probes/verify_v8c_isolation.py`；`docs/TEST_REPORT.md` §3 / `indep_d1_com.py` |
| COM 导出耗时 | **稳态**逐页中位 0.125–0.146s（最大 0.343s）；**冷调用端到端 2.64s/页**（10 页 26.4s），第二次冷调用不摊销 | `docs/TEST_REPORT.md` F3 / `indep_export_cost.py` |
| ⚠️ **COM apartment 会被拆** | `export_pages` 返回后，同线程的旧代理 `RPC_E_DISCONNECTED`、后续 COM 调用 `CO_E_NOTINITIALIZED`；重新 `CoInitialize()` 可恢复 | `docs/TEST_REPORT.md` F1 / `indep_com_apartment2.py` |
| 素材形状计数 | `output/b_multislide.pptx` 顶层 **41** = 40 `p:sp` + 1 chart；过滤 18 空文本框后 **23**（22 text + 1 chart） | 裸 XML 独立核实，`docs/TEST_REPORT.md` §6(a) |
| M2 覆盖率 | 行级中位 **0.286**；矩形法上限 **0.414**；pad 归零 **0.321**；形状级基线 **0.111**；比值 **2.6×** | `docs/IMPL_REPORT.md` D2 / `analyze_m2_ceiling.py`；独立法 0.287 / 0.417 / 0.323 / — / — |
| 图表单元覆盖率 | **0.166**（全场最低） | `docs/IMPL_REPORT.md` D8；`docs/TEST_REPORT.md` §6 |
| 回归基线 | commit `c188b63` = **248**；实现步 1–4 后 545；再加独立用例 **960 collected**。`KNOWLEDGE.md:53` 的"244"**已过时** | `docs/AUDIT_REPORT.md` N3 |

---

## 2. 共用数据契约（单位约定先定死）

### 2.1 单位

| 单位 | 定义 |
|---|---|
| `*_emu` | `int`，OOXML 原生 EMU，**唯一权威坐标** |
| `*_pt` | `float`，`pt = emu / qa.EMU_PER_PT` |
| `*_px` | `float`，**固定在"导出底图像素空间"**：`px = emu × export_width_px / canvas_width_emu`。播放器不重算坐标，靠 CSS `transform: scale()` 缩放整个舞台 |
| 换算常量 | `px_per_pt = export_width_px / (canvas_width_emu / EMU_PER_PT)`；1920 宽 @13.333in 画布时 `= 2.0` |

**规则：EMU 是源，pt/px 是派生。** 只读实现不得反向由 px 推 EMU。
**v2 补充**：`export_width_px` 由 `import` 固化进 `deck.json`，而 `hl_anim.build_player` 有独立的 `canvas_width_px` —— **两者必须一致**，目前**没有任何断言保护**，是最容易在步 5/6 埋雷的点（`docs/IMPL_REPORT.md` §「我不确定的地方」6）。

### 2.2 组合形状坐标换算

子形状的 `left/top` 位于**组的子坐标系**，须做仿射映射：

```
gx, gy, gw, gh  = group.left_emu, group.top_emu, group.width_emu, group.height_emu
cox, coy, cw, ch = group.ch_off_x_emu, group.ch_off_y_emu, group.ch_ext_cx_emu, group.ch_ext_cy_emu

abs_x = gx + (child_left - cox) * gw / cw
abs_y = gy + (child_top  - coy) * gh / ch
abs_w = child_w * gw / cw
abs_h = child_h * gh / ch
```

- **方向已实测正确**（V1 ✅）；取 `chOff/chExt` 走 `group._element.chOff` / `.chExt`；缺失时视为恒等映射。
- **嵌套逐层累乘**：递归时把父级的映射函数作为参数传下去。
- ⚠️ **实现期踩过的坑（写进契约防重蹈）**：子级映射必须吃**本层坐标原值**，绝对化由 compose 链负责；直接把已绝对化的坐标喂进子级会**重复施加外层缩放**（实测 leaf 落在 3in，期望 1in）。
- `gw/cw` 或 `gh/ch` 为 0 → 整组跳过并记 `skipped(reason="degenerate_group")`，**不要除零**。
- 组自身 `left is None` → 整组跳过（记 `group_no_xfrm`），不阻断其余页。
- 子形状 `rotation_deg = child_rot + group_rot`，本期只透传角度、由播放器 `transform: rotate()` 近似。

### 2.3 结构（全部 `@dataclass`，字段名即 JSON 键名）

```python
@dataclass
class RunInfo:
    text: str
    size_pt: float | None          # None = 继承（回退集中在 hl_layout，按 qa._para_font_size 规则）
    bold: bool | None
    italic: bool | None
    font_name: str | None          # a:latin；不含 a:ea（本期不区分中西文字体）

@dataclass
class ParaInfo:
    text: str                      # 文档序拼接 a:r/a:t 与 a:br(→ "\n")、a:fld(取其 a:t)
    runs: list[RunInfo]
    align: str | None              # "LEFT"/"CENTER"/"RIGHT"/"JUSTIFY"/"DISTRIBUTE"，None = 继承
    level: int
    space_before_pt: float | None
    space_after_pt: float | None
    line_spacing_mult: float | None   # v2 新增：倍数形态（float）
    line_spacing_pt: float | None     # v2 新增：Length 绝对值形态（pt）；与上者互斥
    # ⚠️ v1 的单一 line_spacing: float 已废弃（Length 折倍数会重复计入 1.25，见 R-D7）

@dataclass
class ShapeInfo:
    shape_id: int
    name: str
    kind: str                      # "text"|"picture"|"table"|"chart"|"group"|"other"
    left_emu: int | None
    top_emu: int | None
    width_emu: int | None
    height_emu: int | None
    rotation_deg: float
    is_placeholder: bool
    ph_type: str | None            # v2 新增：占位符类型名 "TITLE"/"CENTER_TITLE"/"BODY"/"OBJECT"/
                                   #   "PICTURE"/"SUBTITLE"/...；非占位符为 None（见 R-D4）
    hidden: bool
    # kind == "text"
    paragraphs: list[ParaInfo] = []
    margin_left_emu: int = 91440
    margin_top_emu: int = 45720
    margin_right_emu: int = 91440
    margin_bottom_emu: int = 45720
    word_wrap: bool | None = None       # None = 继承（文本框有效默认 True，但 add_textbox 写的是 False）
    vertical_anchor: str | None = None  # "TOP"/"MIDDLE"/"BOTTOM"/None（None 按 TOP 处理，见 R-V5/V4）
    auto_size: str | None = None
    font_scale: float | None = None     # v2 口径变更：真实倍率，a:normAutofit/@fontScale ÷ 100000
                                        #   "60000" → 0.6；无 = None（见 R-D3）
    # kind == "table"
    table_cells: list[list[list[ParaInfo]]] = []
    table_col_widths_emu: list[int] = []
    table_row_heights_emu: list[int] = []
    table_spans: list[list[tuple[int, int]]] = []   # v2 新增：[row][col] -> (span_width, span_height)
                                                    #   默认 (1, 1)（见 R-D5）
    table_cell_margins_emu: list[list[tuple[int,int,int,int]]] = []  # v2 新增：[row][col] -> (l,t,r,b)
                                                    #   TableCell 内边距未必等于形状级默认值
    # kind == "group"（children 的坐标已按 §2.2 换算为**绝对 EMU**）
    children: list["ShapeInfo"] = []

@dataclass
class PageShapes:
    index: int
    width_emu: int
    height_emu: int
    shapes: list[ShapeInfo]        # 组合以 children 嵌套表示（**不是**已展开为平铺列表，见 R-N1）；已按过滤规则过滤

@dataclass
class DeckIR:
    schema: int                    # 常量 1
    source_pptx: str               # 绝对路径
    mode: str                      # "faithful"|"redesign"
    export_width_px: int
    width_emu: int
    height_emu: int
    pages: list[dict]              # {index, bg: "bg/slide_1.png"|None, shapes: [ShapeInfo…]}
    skipped: list[dict]            # [{page_index, shape_name, reason}]
```

**过滤规则（导入期一次性执行，`read_pages` 内）**，按序判定：

| # | 条件 | reason 标记 |
|---|---|---|
| 1 | `left/top/width/height` 任一为 `None` | `no_xfrm`（组合见 §2.2 例外） |
| 2 | `width_emu <= 0 or height_emu <= 0` | `zero_size` |
| 3 | `hidden is True` | `hidden` |
| 4 | `kind == "text"` 且所有段落 `text.strip() == ""` | `empty_text`（实测命中 18 个） |
| 5 | `kind == "table"` 且所有单元格文本为空 | `empty_text` |
| 6 | `kind == "other"` | `unsupported`（**保号留在 `shapes` 里**供 redesign 参考，不进 units） |

> 过滤顺序已被实测确认与实现一致（`docs/AUDIT_REPORT.md` §4）。**不要**过滤 `picture` 与 `chart`。

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

| code | 触发 |
|---|---|
| `PPTX_NOT_FOUND` | 路径不存在 / 目录当文件 |
| `PPTX_UNREADABLE` | 非 zip / 损坏 / 非 pptx 包 / **包体超限（v2 新增，见 R-M2）** / **缺 `p:sldSz`（v2 新增，见 R-M4）** |
| `PPTX_ENCRYPTED` | 确为加密的 OOXML（见 R-L1：**不得**用 OLE2 魔数一概而论） |
| `PPTX_EMPTY` | `len(prs.slides) == 0` |
| `NO_POWERPOINT` | COM 不可用（未装 Office / pywin32 缺失） |
| `COM_EXPORT_FAILED` | 打开或导出过程中 COM 报错 |
| `NO_FFMPEG` / `NO_BROWSER` | 由 `cli` 侧探测后构造 |
| `TEMPLATE_INVALID` | 自定义模板不可用 |
| `IR_MISMATCH` | deck.json 内部不一致（页数 ≠ 底图数、几何为 `None`、画布为 0 等，见 R-L3） |
| `BAD_ARGS` | 参数组合非法 |

**v2 纪律**：`hl_layout.page_shapes` / `build_units` 等对外函数也一律抛 `PptxError`，**不得**裸 `ValueError`（v1 未规定，实测偏差见 R-D6）。

### 3.2 `read_pages`

```python
def read_pages(pptx_path: str, *, mode: str = "faithful",
               export_width_px: int = 1920) -> tuple[DeckIR, list[dict]]:
    """逐页形状清单。返回 (DeckIR, skipped)，skipped 与 DeckIR.skipped 同对象。
    失败抛 PptxError。"""
```

实现要点（v2 已并入实测教训）：

1. **zip 门禁（v2 新增，必做）**：`Presentation()` **之前**先开 `zipfile.ZipFile`，遍历 `infolist()` 累计 `file_size`（zip 目录里的**声明值，不需解压**）；总量超阈值（建议 200 MB）或**单条** `compress_size/file_size` 比率异常（建议 > 200:1）→ `PptxError("PPTX_UNREADABLE", "包体异常：解压后体积 … MB")`。实测：109 KB 的包靠 `[Content_Types].xml` 展开 **64 MiB** 进内存（比率 ≈601:1），pptx 侧零校验（R-M2）。
2. `prs = Presentation(os.path.abspath(pptx_path))`；**`width_emu = int(prs.slide_width)` 必须并入同一个 `try`**（或读后判空），否则缺 `p:sldSz` 的包会以裸 `TypeError` 冒成 `INTERNAL(5)`（R-M4）。
3. 逐 `slide.shapes` 递归：`MSO_SHAPE_TYPE.GROUP` → 递归 `shape.shapes`；`shape.has_table` → `table`；`shape.has_chart` → `chart`；`shape_type == PICTURE` → `picture`；`shape.has_text_frame` → `text`；否则 `other`。
4. 顺序即 `order` 的来源：**文档序**，不要按坐标排序。
5. 段落文本走 XML：`paragraph._p` 的子元素按文档序，`a:r/a:t` → 文本、`a:br` → `"\n"`、`a:fld` → 取其 `a:t`。**`paragraph.runs` 不暴露 `a:br`/`a:fld`**。⚠️ 这里产生的 `"\n"` 是**硬换行**，语义见 §4.3（R-M1）。
6. `size_pt` 逐 run 取 `run.font.size.pt`（`None` 保留，**不在此回退**）。
7. `line_spacing`：**先判 `isinstance(ls, Length)`**（`Length` 是 `int` 子类，先判数值会把 `Pt(30)` 存成 381000 倍行距）→ 写 `line_spacing_pt`；否则写 `line_spacing_mult`。两者互斥，另一个为 `None`。
8. `auto_size`/`word_wrap`/`vertical_anchor` 直接取属性并**保留 `None` = 继承**语义。
9. `font_scale`：解析 `text_frame._txBody.find(qn("a:bodyPr"))` 下的 `a:normAutofit`，读 `@fontScale` 后 **`÷ 100000`** 存真实倍率（`"60000"` → `0.6`）。
10. `ph_type`：`shape.placeholder_format.type` 的名字（实测取值如 `TITLE (1)`/`CENTER_TITLE (3)`/`BODY (2)`/`OBJECT (7)`/`PICTURE (18)`/`SUBTITLE (4)`）；非占位符为 `None`。
11. `table_spans` / `table_cell_margins_emu`：逐格取 `span_width`/`span_height` 与 `cell.margin_*`（实测可用，且默认值与形状级相同，但不保证模板不改）。
12. `hidden`：`shape._element.nvSpPr.cNvPr.get("hidden") == "1"`；无 `nvSpPr` 的帧类形状（如 `GraphicFrame`）按 `False`。
13. **不做**任何布局推断（那是 `hl_layout` 的事）。

### 3.3 `powerpoint_available`

```python
def powerpoint_available() -> bool:
    """廉价前置探测：pywin32 可导入 且 注册表存在 ProgID "PowerPoint.Application"。
    不启动 PowerPoint 进程（Dispatch 会真启实例，探一个布尔值不值得）。"""
```

- 注册表：`winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "PowerPoint.Application")`，`try/except OSError → False`；非 Windows 直接 `False`。
- **探测"装了"不等于"能用"**：`export_pages` 内部 `DispatchEx` 失败仍须独立包成 `PptxError(code="NO_POWERPOINT")`。
- 中文兜底文案：`"未检测到 PowerPoint，无法导出保真底图。请安装 Microsoft Office（含 PowerPoint），或改用 --mode redesign / 直接跳过 faithful 模式。"`

### 3.4 `export_pages` —— COM 安全节（**v2 大幅重写**）

```python
def export_pages(pptx_path: str, out_dir: str, width: int = 1920) -> list[str]:
    """COM 无窗口导出每页 PNG，返回**绝对路径**列表（页序一致），文件名 slide_1.png 起。
    height = round(width * height_emu / width_emu)。失败抛 PptxError。"""
```

**前提（v2 更正）**：本机 PowerPoint 的 COM server 是**共享的**——`Dispatch` / `DispatchEx` / `GetActiveObject` 拿到**同一个实例**（R-D1）。因此：

> **一切"靠新实例隔离"的假设作废。用户的数据安全完全落在下面这些守卫上，一条都不能省。**

| # | 规则 | 为什么 |
|---|---|---|
| 1 | **绝对路径**：`os.path.abspath(pptx_path)` 与每页输出路径都绝对化 | spike 实证：相对路径被 PowerPoint 按自己的 cwd 解析 → "找不到文件" |
| 2 | **概不写 `app.Visible`** | 理由比 v1 更强：既然我们附着在**用户的**实例上，设 `Visible=False` 会**把用户正在看的窗口藏起来** |
| 3 | 打开：`Presentations.Open(abs_path, ReadOnly=True, Untitled=False, WithWindow=False)` | "无窗口"靠 `WithWindow=False`；**不要**依赖 `app.Visible=False` |
| 4 | **归属判定（Quit 的唯一依据）** | 见下方伪码：`had_powerpoint`（启动前有无 `POWERPNT.EXE`）+ **PID 集合**双守卫 |
| 5 | **只关自己新开的（v2 新增，R-S1）** | Open 前记 `before_names = {p.Name for p in app.Presentations}`；`finally` 里**仅当 `pres.Name not in before_names`** 才 `pres.Close()`；否则只释放引用。用户自己开着同一份稿时，被关掉的可能是**他的**窗口（有未保存修改 → 保存对话框 / 静默丢弃） |
| 6 | **CoInitialize 配对必须"谁初始化谁反初始化"（v2 新增，R-F1）** | 见下 |
| 7 | **模块级 `threading.Lock` 串行化（v2 新增，R-L6）** | 共享实例上并发导出会交错 |
| 8 | 每页导出后校验文件存在且非 0 字节，否则 `PptxError("COM_EXPORT_FAILED", f"第{i}页底图未生成：…")` | 沿用 v1 |
| 9 | 返回**绝对路径列表**，`len(...) == len(prs.slides)` | 沿用 v1 |
| 10 | `out_dir` 用 `makedirs(exist_ok=True)`，**不清理**同名文件（同稿重导会覆盖，属预期） | 沿用 v1 |

**Quit / Close 伪码（v2）

```
had_powerpoint = _powerpoint_running()        # tasklist 找 POWERPNT.EXE；探不到 → True（保守）
                                              #   ⚠️ 采样与 Dispatch 之间有 TOCTOU 窗口（R-L5）
app = win32com.client.DispatchEx("PowerPoint.Application")
pids_ours = _powerpoint_pids()                # Dispatch 之后立刻再探一次 PID 集合
before_names = {p.Name for p in app.Presentations}
we_initialized = _co_initialize()             # 只有返回 S_OK 才算"我们初始化的"
pres = None
try:
    pres = app.Presentations.Open(abs_path, ReadOnly=True, Untitled=False, WithWindow=False)
    … Export 逐页 …
finally:
    try:
        # v2：只关自己新开的那一份（R-S1）
        if pres is not None and pres.Name not in before_names:
            pres.Close()
    except Exception:
        pass
    try:
        # 双守卫：启动前没有 PowerPoint，且 PID 集合仍是我们记下的那一个，且已无稿
        if (app is not None and not had_powerpoint
                and pids_ours and _powerpoint_pids() == pids_ours
                and app.Presentations.Count == 0):
            app.Quit()
    except Exception:
        pass
    pres = app = None
    if we_initialized:
        pythoncom.CoUninitialize()
```

- **`pre_count` 已废弃**：v1 伪码用 `pre_count == 0`，实测 `pre_count` 赋值后从未被读取；实现换成了更严的"退出时 `Presentations.Count == 0`"（属加固方向）。v2 采用实现的判据并**明确写进契约**，不再留死变量。
- **残余 TOCTOU（R-L5，明确接受）**：用户在"采样之后、Dispatch 之前"启动 PowerPoint 时，`had_powerpoint=False` 且结束时 `Count==0` → 会被 Quit。影响有限（起手界面无未保存数据），但与本模块"宁可留一个空进程，也不关用户的东西"的立场不一致。PID 双探把窗口收窄到"我们 dispatch 之后用户才启动"这一小段。若要把窗口降到 0，唯一彻底做法是**永不 `Quit`**（代价：常驻一个隐形 PowerPoint 进程）——作为可选策略记录，默认不启用。
- **R-F1（apartment）**：`export_pages` 返回后，同线程旧代理会 `RPC_E_DISCONNECTED`、后续 COM 调用 `CO_E_NOTINITIALIZED`（**偶发**，实测 1/3）。根因是那对 `CoInitialize`/`CoUninitialize` 与 PowerPoint 打开/导出路径的交互（机制未定）。**规则**：只在**本次真正初始化了**时才 `CoUninitialize`（据 `CoInitializeEx` 的 `S_OK`/`S_FALSE`，或自记 `we_initialized`）。调用方若还要在同一线程用 COM，**必须自行重新 `CoInitialize()`**，代码注释要这么写（v1 的实现注释写成"噪音，不用管"是错的）。
- **V8 未覆盖的一半**：`Presentations.Open` 对**已打开文件**的语义（返回既有 Presentation / 报错 / 另开只读一份）尚未实测，**它直接决定 R-S1 的风险等级**。落地前必须补：手工打开一份含未保存修改的稿 → 跑 `export_pages(同一文件)` → 确认该稿存活。
- **同威胁的另一处（R-N4，不在本次三模块内）**：`app.py::_export_pdf_via_com`（`app.py:753-757`）仍是 `Dispatch` + **无条件** `Close`/`Quit` 的旧写法，失败一律 `return False`。既然是同一个共享实例，它**可能关掉用户的 PowerPoint**。落地步 8 时统一到本节守卫版本。

---

## 4. `hl_layout.py` —— 行级高亮矩形（本期最高风险点）

### 4.1 为什么必须行级

spike 实测文本框矩形内墨迹覆盖率**中位 0.087**（p10=0.005），最差案例框 792×821px 里只有一行小字（0.005）；宽度贴合度 `tight_w` p10=0.149 —— **大量框宽是文字的 6 倍**。直接拿形状矩形做高亮 = 高亮一大片空白。

### 4.2 断行：`qa` 只给行数，必须自己补断点；且**硬换行**语义必须与 QA 分叉（v2 重写）

```python
@dataclass
class Line:
    text: str          # 该行承载的字符（不含被丢弃的行尾空格）
    width_pt: float    # 已放置宽度 = Σ _char_width_pt(ch, size_pt) + 行内空格
    size_pt: float

def wrap_lines(text: str, size_pt: float, box_width_pt: float) -> list[Line]:
    """复刻 qa._measure_lines_ex 的贪心规则，额外返回每行文本与宽度。"""

def paragraph_lines(para: ParaInfo, size_pt: float, box_width_pt: float) -> list[Line]:
    """v2 新增（几何真正的入口）：先按 "\n" 硬拆成硬行段，每段再各自走 wrap_lines，
    最后按序拼接。word_wrap is False 时不做折行，每个硬行段恰 1 行。
    实现中现有私有名 `_paragraph_lines`（`hl_layout.py:170`），契约采用公开名；
    落地时提升为公开或加一层公开薄封装即可，行为以本节为准。"""
```

**v2 的等价性要求（断言对象改了）**：

```python
# ① 折行层与 qa 同源（无硬换行时逐字等价）—— v1 已有，保留
assert len(wrap_lines(t, s, w)) == qa_mod.measure_text_lines(t, s, w)

# ② v2 新增：真正产出几何的函数必须有断言，且要覆盖"无 \n"这一主要情形
assert len(paragraph_lines(para, s, w)) == qa_mod.measure_text_lines(para.text, s, w)   # 仅当 "\n" not in text

# ③ v2 新增：把与 QA 的分歧写成**显式契约**，而不是让它在暗处漂移
n_br = para.text.count("\n")
assert len(paragraph_lines(para, s, w)) == qa_mod.measure_text_lines(para.text, s, w) + n_br
```

**为什么必须分叉（R-M1）**：`qa._tokenize` 把 `"\n"` 当成**一个 1.0em 宽的 CJK 字形**（既非空格、也非 alnum → 落 `else` 分支），**它不构成硬换行**。于是 v1 的"§2.3 把 `a:br` 存成 `"\n"`" + "§4.3 把整段交给 `wrap_lines`"**不可兼得**：

- 照 v1 字面做 → **软换行被吃掉**（渲染错）；
- 实现里 `paragraph_lines` 先硬拆 → 渲染对，但**与 qa 对同一段差 +1 行**（实测 `'第一行\n第二行'`：`paragraph_lines=2`、`wrap_lines=1`、`qa=1`）。

**v2 的取舍（不碰 `qa.py`）**：
1. 折行层保持与 `qa` 严格同源（宽度计算不另写第二份）；
2. 硬换行在 **`paragraph_lines`** 层处理，作为**契约明文规定的有意分歧**，并用断言 ③ 锁住差的**恰好是 `\n` 的个数**；
3. 不修改 `qa._tokenize` 认 `\n` 的写法——那会改动既有 248 条基线里溢出告警的行为，代价大于收益；
4. **已知后果，必须写进风险**：对含软换行的段落，`qa` 的溢出告警会**少算行数**。即"高亮位置"与"溢出告警"在 `a:br` 上**刻意不同**，这是两侧各自的正确性优先，不是 bug。

**现实性边界（不许夸大）**：仓库内三份稿 `a:br` **全为 0**（44/25/51 段，`tools/probes/audit_pptx_io.py` §7），故**当前素材上不发作**；手工汇报稿常见，接真实稿即会踩到。

### 4.3 `build_units`

```python
@dataclass
class Rect:
    left_emu: int; top_emu: int; width_emu: int; height_emu: int
    left_px: float; top_px: float; width_px: float; height_px: float

@dataclass
class Unit:
    order: int                  # 页内讲解序（形状文档序 × 段序，0 基）
    page_index: int
    text: str
    kind: str                   # "title"|"body"|"bullet"|"cell"|"chart"|"picture"|"table"
    rect: Rect                  # 段落 union bbox
    lines: list[Rect]           # 每行 tight rect（渲染默认用这个）
    size_pt: float
    align: str                  # "LEFT"|"CENTER"|"RIGHT"（None 归一为 "LEFT"）
    shape_id: int
    shape_name: str
    is_estimated: bool
    warnings: list[str]

def build_units(page: PageShapes,
                export_width_px: int = 1920,
                pad_x_pt: float = 2.0,
                pad_y_pt: float = 1.0,
                skipped: list[dict] | None = None) -> list[Unit]:
    """skipped 非 None 时，追加 {"page_index","shape_name","shape_id","reason"}。"""
```

**坐标推导（全程 pt，最后一次性转 EMU）**：

```
inner_left_pt = (shape.left_emu + margin_left_emu) / EMU_PER_PT
inner_top_pt  = (shape.top_emu  + margin_top_emu)  / EMU_PER_PT
inner_w_pt    = (shape.width_emu  - ml - mr) / EMU_PER_PT
inner_h_pt    = (shape.height_emu - mt - mb) / EMU_PER_PT
```

| 步骤 | 规则 |
|---|---|
| 退化框 | `inner_w_pt <= 0 or inner_h_pt <= 0` → **跳过该形状，并写一条 `skipped`（v2 强制，R-M3）**。v1 只写"记 warning"却没给承载通道，实测实现里连诊断都没有，`build_units` 也没有出口 → "高亮为什么少了这一块"无从得知。与 `read_pages.skipped` 同构，并由 CLI 冒泡进 JSON 的 `skipped` |
| 段前距 | `cursor_pt += space_before_pt or 0`，**仅非首段生效**（R-V5：实测 PowerPoint **忽略首段 `space_before`**——首行墨迹距框顶仅 5.9pt ≈ margin_top） |
| 字号 | 复用 `qa._para_font_size` 规则：该段第一个有显式 `font.size` 的 run；全无 → `qa.DEFAULT_FONT_SIZE_PT`（18.0）。**不要另立回退值** |
| 断行 | `word_wrap is False` → 每个硬行段恰 1 行；否则 `paragraph_lines(...)`（含硬换行语义，见 §4.2） |
| 行高 | `line_h_pt = line_spacing_pt`（Length 绝对值，**直接使用，不再乘任何系数**）**或** `size_pt * qa.LINE_HEIGHT_FACTOR * (line_spacing_mult or 1.0)`（R-D7） |
| 行宽 | `line.width_pt`（行尾丢弃的空格不计） |
| 水平定位 | `LEFT`/None → `x = inner_left_pt`；`CENTER` → `+ (inner_w_pt - width_pt)/2`；`RIGHT` → `+ (inner_w_pt - width_pt)`；`JUSTIFY`/`DISTRIBUTE` → 按 `LEFT` 并记 `justify_approximated` |
| 行矩形 | `left_emu = round((x - pad_x_pt) * EMU_PER_PT)`、`top_emu = round((inner_top_pt + cursor_pt - pad_y_pt) * EMU_PER_PT)`、`width_emu = round((width_pt + 2*pad_x_pt) * EMU_PER_PT)`、`height_emu = round((line_h_pt + 2*pad_y_pt) * EMU_PER_PT)` |
| 推进 | `cursor_pt += line_h_pt`；段尾 `cursor_pt += space_after_pt or 0` |

**垂直锚点**（`builder.py` 的大量 `vertical_anchor=MIDDLE` 实测高频）：

```
if vertical_anchor == "MIDDLE": dy = max(0.0, (inner_h_pt - consumed_pt) / 2)
elif vertical_anchor == "BOTTOM": dy = max(0.0, inner_h_pt - consumed_pt)
else: dy = 0.0                      # TOP / None（None 实测按 TOP 处理正确，V4 ✅）
所有 rect.top_emu += round(dy * EMU_PER_PT)
```

**`font_scale` 用法（v2）**：字段存的是**真实倍率**，故 `size_pt *= font_scale`（直接相乘）。给每个 Unit 记 `autofit_scaled` warning。
> 实测补充（V7 ✅）：注入的 `fontScale` 被 PowerPoint 打开时**重算覆盖**（`100000` → 写成裸 `<a:normAutofit/>`），导出仍按原字号渲染 → 该风险的现实概率低于 v1 预期，属防御性处理。

### 4.4 宽度取值：推荐**行级 tight**，不是整行宽度

| 方案 | 覆盖率 | 观感 |
|---|---|---|
| 整行宽度（`inner_w_pt`） | 与形状级同病——短行后面拖一片空白 | 高亮条整齐，但"高亮空白"仍在 |
| **行级 tight（推荐）** | 单行矩形只包住该行墨迹 | 高亮像荧光笔，跟着文字走 |

**理由**：分解实测显示覆盖率的两项主导因子是**竖向填充 0.731**（行高框比 CJK 墨迹高约 27%）与**横向填充 0.960**（tight 宽度已接近 1）；横向一旦退回"整行宽度"，末行（ragged-right，中文汇报稿的常态）必然塌到 0.1 量级。留 `pad_x_pt=2.0 / pad_y_pt=1.0` 补偿墨迹触框边（spike：**12/22 行墨迹触框边**）。

**渲染**：播放器按 `Unit.lines` 画**多个 rect**（每行一个），`Unit.rect` 只作 union bbox 用于滚动定位、命中测试与 QA 汇总。单行段落两者等价。

### 4.5 覆盖率度量与 M2 验收（**v2 重写：0.35 已证不可达**）

```python
def measure_coverage(bg_png: str, rect_px: tuple[int, int, int, int],
                     bg_rgb: tuple[int, int, int] | None = None,
                     diff_thresh: int = 30) -> float | None:
    """rect_px = (left, top, width, height) 在底图像素空间。
    bg_rgb 为 None 时取裁剪区四边 1px 边框环的中位色（**验收脚本必须显式传局部底色**，见下）。
    ink = 任一通道 |px - bg| > diff_thresh 的像素；返回值 = ink / 面积。

    v2 语义变更（R-L7）：退化矩形（w<=0 或 h<=0）与越界矩形（与图像无交集）返回 None，
    不再与"这块真的没墨迹"共用 0.0 —— 否则度量本身失去发现"定位算错"的能力。
    部分越界时先裁剪再计算，并在返回值上不做特殊标记（是否越界由调用方判 rect 合法性）。
    """
```

**口径纪律（实测驱动，必须遵守）**：
- 默认底色估计（内侧 1px 环中位）在 **rect 恰好等于墨迹块时会失明**（环本身也是墨迹色 → 无对比 → 测出 0.0）。对照图实测：真值 1.0 → 契约口径 0.0。本素材 `ring_ink ≤ 0.037`（远未到失明阈值），**中位偏差仅 0.003**，所以安全；但**验收脚本必须显式传局部底色**。构造极端下（rect 上下边被文字铺满）默认口径低估约 20 个百分点。
- 该失效行为已被用例显式锁住（`test_measure_coverage_ring_bg_blind_on_uniform_block`）。

**M2 v2 验收判据（全部可达、可复现）**：

| 判据 | 阈值 | 实测 | 说明 |
|---|---|---|---|
| **主·相对** | 行级中位 / 形状级中位 **≥ 2.5×** | **2.6×** | 同函数、同底色、同素材 → **自相对**，不受绝对尺度与机器差异影响 |
| **副·绝对** | `pad_x = pad_y = 0` 时行级中位 **≥ 0.30** | **0.321**（独立法 0.323） | 归零敏感性实测：0.0pt→0.321 / 0.5pt→0.305 / 1.0pt→0.290 / 2.0pt→0.264 |
| 溢出 | 行 rect 越出形状框的比例 **≤ 5%** | **0/27 = 0.00%** | — |
| ~~绝对 0.35~~ | **废弃** | 不可达 | 见下 |

**0.35 为什么不可达（保留证据，防止再被提出）**：分解 `0.731（竖向填充） × 0.960（横向填充） × 0.414（CJK 字形墨迹占 em 盒比例） = 0.286`，与实测中位完全自洽。**基于字体度量的任何矩形法上限 ≈ 0.414**（即把 rect 精确设成墨迹包围盒——而这需要先有底图，构成循环依赖，字体度量法拿不到）。pad 归零也只到 0.321。三个独立口径（外侧远环众数 + 欧氏距离 / Otsu 双簇 / 契约默认）分别给 0.287 / 0.235 / 0.286，量级一致，**确认不是口径 artifact**。

**作用域声明（R-F2，必须写进 M2 的结论里）**：
- 该素材上 **27 个单元 → 27 条行 rect，多行单元 0 个**；把 `word_wrap` 全强置 `True` 后仍是 27 → **没有任何段落真的需要折行**。原因是 `python-pptx` 的 `add_textbox` **默认写 `wrap="none"`**（该素材 13/22 个文本框）。
- 因此本素材上的"行级 tight"数值上**等价于"段级 tight"**，`wrap_lines` 那层**只被单元测试覆盖，未被验收素材碰过**。
- **结论**：M2 的覆盖率结论**仅对单行段落成立**。要外推到会折行的真实稿，必须另备一份长文本折行素材重测（见 §10）。

### 4.6 非文本形状的 Unit（v2 修正 D4 / D5 / D8）

| kind | 处理 |
|---|---|
| `title` | 命中三条任一（**三条 v2 均可实现**）：① 形状名含 `Title`；② `is_placeholder and ph_type in {"TITLE","CENTER_TITLE","SUBTITLE"}`（**v2 依赖 §2.3 新增的 `ph_type`**）；③ 字号 ≥ 1.3× 本页正文中位字号。全不中 → `body` |
| `body` / `bullet` | v1 的 `kind` 枚举里有 `bullet` 却**没给判定规则**（实现者自行用"body 且该形状段落数 > 1"作启发式）。v2 认可该启发式并写进契约：**`bullet` = `body` 且该形状段落数 > 1**；kind 只影响展示语义，不影响几何 |
| `picture` | 一个 Unit，`rect = lines = [形状外框]`。⚠️ **v2 删除"图片本身就是墨迹、coverage 天然接近 1"**：覆盖率取决于图片自身留白，无任何保证（见 R-D8） |
| `chart` | 一个 Unit，`rect` = 形状外框（R4：**不拆数据点**），`text` = 图表标题或空串。⚠️ **实测 coverage 0.166（全场最低）**——图表内部大片留白是常态 |
| `table` | 单元格级，每个非空单元格的每个段落一个 Unit，`kind="cell"`。x/y 由 `table_col_widths_emu` / `table_row_heights_emu` 前缀和 + **`table_cell_margins_emu`** 推出。**合并单元格（R-D5）**：只对 `is_merge_origin` 出 Unit；其 rect 宽 = **跨列宽之和**、高 = **跨行高之和**（按 `table_spans` 求和），而不是单列宽；`is_spanned=True` 的格跳过 |
| `other` | 不出 Unit（保号在 `shapes` 供 redesign 参考） |

**验收统计口径（R-D8）**：coverage 中位数**只对文本类 Unit 的行 rect 计算**；`chart` / `picture` / `table` 单元**单列报数**，不并入中位数（否则 1 个 0.166 的图表单元会无谓拉低整页中位）。

---

## 5. 产物目录、`deck.json` 与桥接

```
<out>/
├── deck.json                # DeckIR（源形状 + 页序 + 画布），**唯一真源**
├── bg/slide_1.png …         # COM 底图（faithful）或 None（--no-com / redesign）
├── units.json               # build_units 快照（QA / 调试，非真源）
└── player/index.html        # animate 产出
```

**决策（不变）**：`deck.json` 只存 `shapes`，不存 `units`。`units` 是 `shapes` + `export_width_px` 的纯函数结果，存两份必然漂移；`animate`/`video` 现场用 `build_units` 重算。

**`bg` 一律写相对 `<out>` 的 POSIX 路径**（`"bg/slide_1.png"`）；绝对路径会让 Flask 静态路由与 `file://` 打架。

**页数一致性**：`len(pages) == len(bg 文件)`；faithful 模式下不符即 `PptxError("IR_MISMATCH", …)`。

### 5.1 `page_shapes` 桥接（**v2 新增，R-D6**）

v1 没有定义"`DeckIR.pages` 的元素 → `PageShapes`"的过渡函数，而 `DeckIR.pages` 每项是 `{index, bg, shapes}`、**不含画布尺寸**，`build_units` 却需要画布宽把 EMU 换算到导出底图像素空间。v2 正式收编实现期补的桥：

```python
def page_shapes(page: dict, deck: DeckIR | dict) -> PageShapes:
    """DeckIR.pages 的元素 + 整份 deck → PageShapes（补齐 width_emu/height_emu）。
    deck 缺失或字段不全 → PptxError("IR_MISMATCH", "缺少画布尺寸，无法定位高亮：…")。"""
```

- **唯一合法入口**：`build_units` 只接受 `PageShapes`；收到裸 dict 时给明确中文错误，**不静默算错**。
- **异常类型**：抛 `PptxError` 而非裸 `ValueError`（v1 §3.1 规定 PptxError 是唯一对外异常；实测偏差已记入 R-D6）。

### 5.2 被改坏的 `deck.json` 必须走 `IR_MISMATCH`（**v2 新增，R-L3**）

`deck.json` 被称作"唯一真源"，但 `load_deck` 不做几何校验 → 几何为 `null` / 画布 `0` 时会在 `build_units` 里裸崩（`TypeError` / `ZeroDivisionError`），错误码退化成 `INTERNAL(5)`。

**v2 要求**：`load_deck`（及 CLI 侧）对每个 page 校验收：`width_emu`/`height_emu` 为正整数、每个 shape 的几何字段为 `int`（`kind == "group"` 的子形状同理）、`len(pages)` 与 `bg` 数一致、表格三组数组长度自洽 → 违反即 `PptxError("IR_MISMATCH", …)`。

---

## 6. `hl_anim.py` —— 高亮播放器

### 6.1 接口

```python
def build_player(out_dir: str,
                 bg_paths: list[str],
                 pages_units: list[list[Unit]],
                 title: str = "",
                 dim: float = 0.25,
                 auto_step_ms: int = 2000,
                 canvas_width_px: int = 1920,
                 canvas_height_px: int = 1080,
                 check_exists: bool = True) -> str:
    """写 <out_dir>/index.html，返回其路径。零截图、零 iframe、零外部依赖。"""
```

**参数校验（v2 扩展）**：

| 情形 | 行为 |
|---|---|
| `len(bg_paths) != len(pages_units)` | `PptxError("IR_MISMATCH", "底图数与讲解单元页数不一致：…")` |
| `dim` 越界（不在 `(0, 0.6)`） | `PptxError("BAD_ARGS", "降暗比例需在 (0, 0.6) 之间")` |
| **`bg_paths` / `pages_units` 非列表，或单元里含 `None`** | **v2 新增（R-F4）**：返回 `PptxError("BAD_ARGS", …)`，而不是裸 `TypeError`/`AttributeError`。实测这三类输入会漏成裸异常 |
| **`check_exists=True` 且某底图文件不存在** | **v2 新增（R-L4）**：`PptxError("IR_MISMATCH", "底图缺失：bg/slide_3.png，请重新执行 pptgen import")`。实测缺图时页面全黑而 `window.hl.ready` 照常 resolve → 步 5 的 `shot_player` 会把黑帧安静地喂给 ffmpeg |

`dim` / `auto_step_ms` / `canvas_*` 一律先 `float()`/`int()` 强转 + 范围校验，再 `json.dumps`，**不进字符串拼接**（实测：`nan` / `True` / `"abc"` / `0` / 负数全部被正确拦下）。

### 6.2 结构：**不用 iframe**

与 `anim.py` 的关键差异：`anim.py` 必须 iframe（要连设计稿 HTML 的脚本一起加载）。高亮播放器的"幻灯片"是**一张 PNG** + 绝对定位的 div，**同源限制在这里毫无收益** → 推荐**零 iframe**。审计已确认实现为真零 iframe（`test_zero_iframe` 断言），带填充等高，攻击面比 `anim.py` 更小。

若实现者出于别的原因引入 iframe，**必须** `sandbox="allow-scripts"` 且**禁止**加 `allow-same-origin`（KNOWLEDGE.md 已知风险表 + 计划硬约束 2，审计会查）。

### 6.3 转义纪律（照抄 `anim.py` 的三上下文 + 单遍替换）

| 注入点 | 上下文 | 处理 |
|---|---|---|
| `title` | HTML 文本 **且** `<title>` | `html.escape(title)` |
| 每个 `bg_path` | `src` **属性** | v1 要求 `html.escape(p, quote=True)`；实测实现改用**百分号编码 + JS `setAttribute`**，**强度等价或更强**（真 Chromium 12/12 无注入、编码路径能取回真实文件）→ v2 **追认**该写法，两种任选 |
| `CFG` 对象 | `<script>` **字符串** | `json.dumps(…, ensure_ascii=False)` 后 `.replace("</", "<\\/").replace("<!--", "<\\!--")` |
| 模板占位符 | — | **单遍** `re.sub(r"__TITLE__\|__BGJSON__\|__UNITS__\|__CONFIG__", lambda m: mapping[m.group(0)], TEMPLATE)`。**禁止链式 `str.replace`** |

**已实测的对抗结论**（真 Chromium，含阳性对照）：12/12 恶意 payload 不可执行；`document.title` 与 `h1.textContent` 均等于原始恶意串（证明被当字面文本渲染）；全文 `</script` 恰好 1 次；`<!--` 被转义成 `<\!--` 故 mXSS 的双重转义态不可达；路径逃逸对 `../`、`bg/../../x`、`..\..\x`、`C:/`、`\\srv\…`、`/etc/passwd` 一律 `IR_MISMATCH`。

**仍保留的已知边界（R-F5，已降级）**：`bg/..%2f..%2fetc.png` 不被 `_check_bg_path` 拦下，但该值在播放器里**只被引用、从不被本程序读取**，且所有出口都过 `_web_path`（`%` → `%25` 再编码），解码一次只得到字面文件名、**不构成路径分隔符** → 越界不可达。若将来有别的调用方直接拿 `bg` 值拼 URL，问题会复活 → v2 建议仍补一道 `unquote` 后再跑同一套 `normpath` 检查（防御性，非阻塞）。

### 6.4 交互与程序驱动 JS API

- 步进/回退：`Space`/`→` 下一步、`←` 上一步；点击画面下一步；`revealAll()`；页点导航。
- 自动讲解：`setInterval(stepFwd, CFG.autoStepMs)`。
- 视觉：当前 Unit 的 rect 高亮，其余降暗 `dim`（实现用 **SVG mask 挖洞**，实测两帧差异区域 bbox 非空、远离高亮处像素确实变暗）。
- 16:9 等比缩放：舞台固定 `canvas_width_px × canvas_height_px` + CSS `transform: scale()`。
- **`step == 0` 不降暗（v2 追认实现取舍）**：v1 字面推得"进页即全屏降暗"，实现改为"有高亮才进聚光模式"——否则观众进页第一眼是一片黑。**v2 采纳实现**。

```js
window.hl = { goto(page, step), next(), back(), state(), ready }
```

- `goto` **必须同步生效**（无 CSS 过渡、无 rAF 依赖）。实测实现用"多张 `<img>` 叠放 + display 切换"而非"单张换 `src`"，因为后者会触发异步图片加载、破坏同步性（`test_goto_is_synchronous_no_transition`）。
- **`goto` 的越界语义（v1 未定，v2 采纳实现）**：**钳制**——`page` 钳到 `[0, n-1]`、`step` 钳到 `[0, units]`，不取模、不抛错。
- **未使用字段（R-N2，提示）**：`CFG.title` 与每个单元的 `t`（文本）被写进 HTML 但**模板 JS 从不读**。这解释了两侧"单元文本注入在任何上下文都无落点"（是优点），但也说明这些字节纯属冗余（`units.json` 会明显变胖）。⚠️ **若将来要在播放器里显示讲解文字，`u.t` 一旦进 DOM 就是新的注入点，必须重新走一遍 HTML 文本上下文的注入审计。**

### 6.5 截图序列（供 `pptgen video --mode step`）

`shot_deck` 硬绑 `.slide` 选择器与设计稿结构，**不能用于播放器**。契约：

```python
def shot_player(player_html: str, out_dir: str,
                steps: list[tuple[int, int]],
                min_settle_ms: int = 300,
                max_settle_ms: int = 2000) -> list[str]:
    """逐 (page, step) 驱动 window.hl.goto → 截图 → 写 step_0001.png…，返回路径列表。"""
```

必须沿用 `shot.py` 的三重确定化（`window.hl.ready` → `screenshot(animations="disabled")` → 连续两帧字节一致、上限 `max_settle_ms`）。**别退回固定 sleep**。

> ⚠️ **待批准（仍未定，非 v2 能自决）**：`_launch_browser` / `_screenshot_settled` 是私有名。**建议**在 `shot.py` 加两个**纯新增**公开薄函数（`launch_browser(p)` / `shot_page(page, min_ms, max_ms)`），零行为改动、248 基线零影响；退路是 `from shot import _launch_browser, _screenshot_settled` 并注释写明"有意复用私有符号"。
> **现状**：`tools/probes/accept_m3.py` 目前走的是退路（import 私有符号），**步 5 落地时必须统一**，且**禁止抄第三份**确定化逻辑。此项已上报中枢（ppt-arch → pi）。

---

## 7. `pptx_out.py` —— 母版 + 占位符导出

### 7.1 接口

```python
@dataclass
class TemplateInfo:
    path: str
    layout_names: list[str]
    placeholders: dict[int, list[tuple[int, str, str]]]   # layout_idx -> [(ph_idx, ph_type, name)]
    has_theme_cjk: bool

def read_template(path: str) -> TemplateInfo: ...
def build_builtin_master(out_path: str, title: str = "演示文稿") -> str: ...
def build_deck_pptx(deck: DeckIR, out_path: str,
                    template: str | None = None, no_com: bool = False) -> str: ...
```

**自定义模板存放位置（R-V13 更正）**：v1 说放 `output/templates/` 会"被 `custom_templates()` 当 JSON 解析崩掉"——**实测不成立**：`custom_templates()` 只认 `.json` 且**逐文件吞异常**，`.pptx` 放进去只会被**静默跳过**。结论不变（**另放 `output/templates/pptx/`**），但理由改为"避免与 `template.py` 的 JSON 风格模板目录混淆、且不依赖它的静默行为"。

### 7.2 内置 5 版式（映射到 python-pptx 默认模板的真实 layout）

| 版式 | layout idx | 填充的占位符 idx |
|---|---|---|
| 封面 | `0` Title Slide | `0` CENTER_TITLE ← 标题；`1` SUBTITLE ← 副标题 |
| 目录 | `1` Title and Content | `0` TITLE；`1` OBJECT ← 目录条目（多段） |
| 内容 | `1` Title and Content | `0` TITLE；`1` OBJECT ← 要点（多段，`level` 做缩进） |
| 数据 | `5` Title Only | `0` TITLE；图表/表格用 `add_chart` / `add_table` 落到正文区 |
| 尾页 | `0` Title Slide | `0` CENTER_TITLE ← 结束语 |

> `layout 6 (Blank)` 只有 `DATE/FOOTER/SLIDE_NUMBER`，**不要**用它做内容页。

### 7.3 占位符类型映射

| 输入形状 | 占位符类型 | 填充方式 |
|---|---|---|
| title | `TITLE` / `CENTER_TITLE` | `slide.placeholders[0].text_frame` |
| body / bullet | `BODY` / `OBJECT` / `SUBTITLE` | 同上，逐段 `tf.add_paragraph()`，`level` 透传 |
| picture | `PICTURE` | `ph.insert_picture(abs_path)`；无该占位符则退化为 `shapes.add_picture(...)` |
| table | 无对应枚举 | `shapes.add_table(rows, cols, …)` 落在 OBJECT 占位符矩形内；表格本身可编辑 |
| chart | 无对应枚举 | `shapes.add_chart(XL_CHART_TYPE.X, …)`，类型取值对齐 `builder._CHART_TYPES` |

**清空文本**：用 `tf.clear()`（留一个空段落），再逐段 `tf.paragraphs[0]` / `tf.add_paragraph()`。**不要** `tf.text = "…"`（只能造一个无字体控制的 run）。

**占位符可能不存在**（自定义模板千奇百怪）：按 `ph_type` 查找，找不到则退化为在安全区 `Inches(0.6), Inches(0.7), Inches(12.1), Inches(6.1)` 上 `add_textbox`，并在 stderr 记 `degraded_placeholder`。

### 7.4 中文字体：必须显式写 `a:ea`

实测（§1.2）：默认模板 `theme1.xml` 里 `a:ea=""`、`<a:font script="Hans" typeface="宋体"/>`。`run.font.name` **只写 `a:latin`**，中文按 `a:ea` 匹配 → 不处理时中文渲染为**宋体**。

| 母版来源 | 策略 |
|---|---|
| **内置母版** | 双管齐下：① 每个 run 用 `builder._set_ea` 同款逻辑补 `a:ea`（**必须 append 在 `a:latin` 之后**）；② `prs.save()` 后**后处理 zip**：`<a:ea typeface=""/>` → `微软雅黑`、`<a:font script="Hans" typeface="宋体"/>` → `微软雅黑`。**② 是必要的**——用户在 PowerPoint 里新敲的字走主题字体 |
| **自定义模板** | **不覆盖字体**（只填文本、不设 `font.name`），尊重品牌模板；仅在 `has_theme_cjk is False` 时补写 run 级 `a:ea` 并 stderr 提示 |

zip 后处理：`zipfile.ZipFile(src)` 读 → 新建临时 zip **全量复制**（除 `theme1.xml`）→ `os.replace`。**必须保留 `[Content_Types].xml` 与其余全部条目**。

### 7.5 超长文本：**先缩字号，再截断**

```python
def fit_text(text: str, size_pt: float, box_width_pt: float, box_height_pt: float,
             min_size_pt: float = 12.0) -> tuple[float, str, bool]:
    """返回 (最终字号, 最终文本, 是否被截断)。用 qa.measure_text_lines 试算，
    逐档下调字号（步长 1pt，下限 12pt）直到 fits；仍放不下则二分截断并追加 "…"。"""
```

**推荐缩字号优先**：截断会**静默丢内容**——CLI 的调用方是 agent，`{ok:true}` 里少一句话它无从察觉；缩字号保住信息，只在 12pt 仍溢出时截断，且 `truncated` 必须冒泡到 stderr 与 JSON 的 `warnings`。**不支持** `MSO_AUTO_SIZE`/`normAutofit` 自动缩（Python 侧算不出 PowerPoint 的渲染结果）。

---

## 8. `cli.py` —— `pptgen` 入口（方向 B）

### 8.1 进程契约

- **stdout 恒为一行 JSON**（`ensure_ascii=False`）；人类可读进度/警告走 **stderr**。
- **永不上抛裸堆栈**：`PptxError` / `ValueError` / `RuntimeError` / 未捕获异常统一收敛成 `{"ok": false, "error": {code, message, hint}}` + 非 0 退出码。
- 中文错误必须带**可执行的修复指引**（`hint`）。
- 退出码：`0` 成功｜`2` 参数错误｜`3` 输入问题（`PPTX_*`/`IR_MISMATCH`/`TEMPLATE_INVALID`）｜`4` 缺外部依赖（`NO_POWERPOINT`/`NO_FFMPEG`/`NO_BROWSER`）｜`5` 内部错误。

```json
{"ok": true, "cmd": "import", "data": {"out": "...", "pages": 10, "units": 41, "skipped": 18, "warnings": []}}
{"ok": false, "cmd": "import", "error": {"code": "NO_POWERPOINT", "message": "…", "hint": "…"}}
```

**v2 补充**：
- `skipped` 必须**合并两个来源**：`read_pages` 的导入期过滤（实测该素材 18 条 `empty_text`）**加上** `build_units` 的布局期丢弃（`inner_w_pt <= 0`，见 R-M3）——后者在 v1 里没有任何出口。
- `warnings` 冒泡 `truncated`、`autofit_scaled`、`justify_approximated`、`is_estimated`（`qa.font_available()` 为 False）等。

### 8.2 子命令与参数

| 命令 | 必填 | 可选 | 关键行为 |
|---|---|---|---|
| `import <pptx>` | `--out DIR` | `--mode faithful\|redesign`（默认 `faithful`）、`--width {1280,1920,2560}`（默认 1920）、`--pages N`、`--no-com` | 写 `<out>/{deck.json,bg/,units.json}` |
| `animate <dir>` | — | `--dim 0.25`、`--auto-ms 2000` | 读 `deck.json` → `hl_anim.build_player` → `<dir>/player/index.html` |
| `video <dir>` | — | `--sec 4`、`--fps 25`、`--fade 0.8`、`--mode page\|step`（默认 `page`）、`-o out.mp4` | `page` 走 `video_mod.synthesize`；`step` 走 `hl_anim.shot_player` |
| `export <deck.json>` | `-o out.pptx` | `--template brand.pptx`、`--no-com` | `pptx_out.build_deck_pptx` + 自动跑 `qa.check_pptx` |
| `deck "<主题>"`（可选） | `--out DIR` | `--template`、`--mode` | 复用 `outline`→`style`→`html_gen` 流水线 |

**硬边界**：
- `--no-com` + `--mode faithful` = **非法**，报 `BAD_ARGS`：faithful 的视觉保真**完全依赖** COM 底图（D8/D9）。
- `--no-com` 的合法用途：`import --mode redesign`、`export`。
- `--width` 只在 `import` 生效并固化进 `deck.json`；`animate`/`video` 不再接受 `--width`（坐标已按该宽度算好 px）。
- `-o` 省略时默认 `<dir>/out.mp4`（`video`）/ `<dir>/<stem>.pptx`（`export`）。
- 默认 `--out`：`<项目根>/output/pptx_src/<pptx 主名>/`。
- **冷启动耗时写进帮助文本**：`import` 是冷调用形态，10 页稿实测约 **26s**（2.64s/页），且第二次**不摊销**（R-F3）。若在意体感，需另议常驻实例方案（代价是与 §3.4 的实例隔离策略纠缠）。

### 8.3 路径与输入安全

- `--template` 必须 `os.path.isfile` 且后缀 `.pptx`；**真正的红线是"不把传参路径拼进输出路径"**：所有输出路径一律 `os.path.basename()` 后拼到 `<out>` 下，`..` / 绝对路径一律拒绝。
- `deck.json` 里的 `bg` 必须校验为 `<out>` 内的相对路径（`os.path.normpath` 后不以 `..` 开头、非绝对），否则 `IR_MISMATCH`。**v2 补强（R-F5）**：先 `unquote` 再跑同一套检查。
- **zip 门禁（v2 新增）**：`read_pages` 的前置闸门（§3.2 步 1）是**唯一的**包体防线；上传路由（步 8）**不得**绕过它。
- **已核查无问题、不必重复排查的项**（审计结论）：XXE / billion laughs（库内 `resolve_entities=False` + lxml 放大系数上限，实测 0.002s 抛错）；python-pptx **零联网**（外部关系只记录不解引用）；`tasklist` 调用无 `shell=True`、无命令注入面；`app.Visible` 全仓零赋值。

### 8.4 skill 侧（D14/D15）

`skill/ppt-anim/scripts/pptgen.py` 是**薄封装**：定位项目 venv 的 python，转发到 `cli.py`，**不复制任何逻辑**。SKILL.md 正文写"何时用/怎么调/产物在哪/常见错误码含义"，错误码直接引用 §9。

---

## 9. 错误码总表

| code | 退出码 | message 示例 | hint |
|---|---|---|---|
| `BAD_ARGS` | 2 | `--no-com 与 --mode faithful 不能同时使用` | `faithful 模式需要 PowerPoint 导出底图；请去掉 --no-com，或用 --mode redesign` |
| `PPTX_NOT_FOUND` | 3 | `找不到文件：xxx.pptx` | 检查路径是否正确（建议用绝对路径） |
| `PPTX_UNREADABLE` | 3 | `PPTX 无法打开：xxx.pptx（<原因>）` / `包体异常：解压后体积 512 MB` / `PPTX 缺少幻灯片尺寸定义（p:sldSz）` | 确认是标准 .pptx（**旧的 .ppt 请先在 PowerPoint 里另存为 .pptx**）；重做的包体异常请用 PowerPoint 重新导出一份 |
| `PPTX_ENCRYPTED` | 3 | `PPTX 已加密，无法读取` | 请先用 PowerPoint 去掉打开密码再导出 |
| `PPTX_EMPTY` | 3 | `PPTX 中没有任何幻灯片` | — |
| `IR_MISMATCH` | 3 | `底图数与讲解单元页数不一致：10 vs 9` / `底图缺失：bg/slide_3.png` / `缺少画布尺寸` | 重新执行 `pptgen import` 生成 deck.json |
| `TEMPLATE_INVALID` | 3 | `模板不可用：xxx.pptx（<原因>）` | 模板需为标准 .pptx，且至少含一个版式 |
| `NO_POWERPOINT` | 4 | `未检测到 PowerPoint，无法导出保真底图` | 安装 Microsoft Office（含 PowerPoint）；或改用 `--mode redesign` |
| `COM_EXPORT_FAILED` | 5 | `第 7 页底图导出失败（已完成 6/10 页）` | 检查 PowerPoint 是否被其他程序占用/是否弹出了对话框 |
| `NO_FFMPEG` | 4 | `未找到 ffmpeg` | `winget install ffmpeg` |
| `NO_BROWSER` | 4 | `未找到可用的 Chrome/Edge` | 安装 Google Chrome，或 `playwright install chromium` |
| `INTERNAL` | 5 | `内部错误：<异常摘要>` | 附 stderr 完整堆栈；这属于 bug，请附命令与 deck.json |

**文案纪律**：`message` 说"发生了什么"，`hint` 说"下一步做什么"；二者都不出现 Python 异常类型名与英文堆栈（堆栈只写 stderr）。
**v2 纪律**：用户输入畸形（`p:sldSz` 缺失、包体超限、deck.json 被改坏）**必须**落在 `PPTX_UNREADABLE` / `IR_MISMATCH`，**不得**退化成 `INTERNAL`（实测偏差见 R-M4 / R-L3）。

---

## 10. 待实现阶段验证清单（v2 更新：已验/未验）

**已实测通过、可从清单移除**（证据见 §1.2）：

| 原编号 | 结论 |
|---|---|
| V1 / V2 | ✅ 组合仿射换算方向正确；删 xfrm 后四属性全 `None` |
| V4 | ✅ `vertical_anchor = None` 按 TOP 处理正确（实测墨迹贴框顶） |
| V5 | ✅ **首段 `space_before` 被 PowerPoint 忽略**（首行墨迹距框顶 5.9pt ≈ margin_top）→ 契约已改 |
| V6 | ✅ Length 形态行距实测 31.5pt，契约"折倍数"偏差 +19% → 契约已改（R-D7） |
| V7 | ✅ `fontScale` 被 PowerPoint 打开时重算覆盖 → 风险低于预期 |
| V9 | ✅ `is_merge_origin` / `span_width` / `span_height` 可用 |
| V13 | ✅ `custom_templates()` 只认 `.json` 且逐文件吞异常 → `.pptx` 是**静默跳过**而非崩溃 |

**仍未验（v2 重排，按风险降序）**：

| # | 不确定点 | 影响 | 验证方法 |
|---|---|---|---|
| **N1** | **PowerPoint 对"已打开文件再次 `Open`"的语义** | **直接决定 R-S1 是"严重"还是"无风险"** | 真机 + 一份**可牺牲**的稿：先手工打开（含未保存修改）→ 跑 `export_pages(同一文件)` → 确认该稿存活、是否弹保存框。**这是现在最该先做的一条** |
| N2 | 折行素材上的 coverage（R-F2） | M2 结论能否外推 | 造一份**确实折行**的长文本稿（`word_wrap=True` + 超宽段落），重跑 coverage 分解 |
| N3 | 真实手工 PPT（原 V12，**用户未提供**） | M2 结论的成立范围 | 用户提供含图片/表格/SmartArt 的真稿。现状缓解：`tools/make_fixture_pptx.py` 自造 6 页覆盖组合/表格+合并/图片/图表/行距/中英混排，但**覆盖不了**：真实字体替换、SmartArt、嵌入视频、**经 PowerPoint 打开并重存后被重排过**的稿 |
| N4 | `a:br` 的真实稿分布（R-M1） | 硬换行分叉是否被触发 | 拿一份手工汇报稿统计 `a:br` 数（仓库内三份全为 0） |
| N5 | 老 `.ppt` 的实际误判（R-L1） | 错误指引是否误导用户 | 拿一份任意 `.ppt` 喂 `read_pages` |
| N6 | `font_scale` 端到端（R-D3）与 Length 行距（R-D7） | 口径变更后的实际渲染 | 造带 `normAutofit` 与绝对行距的稿并 COM 底图对照 |
| N7 | 真实字体替换（原 V10）/ 旋转形状（原 V11） | 行宽失真 / rect 不跟随 | 非雅黑字体的稿；旋转文本框 |
| N8 | F1 的真实事故场景 | Flask worker 里"先 COM 后 export_pages" | 在 Flask 里实测（本轮只证明了 apartment 会被拆、重新 init 可恢复） |
| N9 | 多线程并发访问同一 PowerPoint 实例 | 步 6/8 的后台路径 | 实测 3 线程并发时 2/3 报 `Slide.Export : Object does not exist` → **已确认不可靠**，落地时必须靠 §3.4-7 的模块级锁串行化 |

---

## 11. 给实现者的清单

### 11.1 实现顺序与验收

| 步 | 模块 / 里程碑 | 要实现的函数 | 状态 | 验收命令 |
|---|---|---|---|---|
| 1 | **`pptx_io.py`**（M1） | `PptxError`、`powerpoint_available()`、`read_pages()`、`export_pages()`、`deck_to_dict`/`load_deck` | ✅ 已完成（`9e9c55a`） | 对 `output/b_multislide.pptx`：读出 10 页；顶层形状 **41**（**40 个文本框 + 1 个图表**），导入期过滤 **18** 个空文本框后保留 **23**（22 text + 1 chart）；坐标非 `None` 的顶层形状 0 个；COM 导出 10 张 1920×1080 PNG，**稳态**逐页中位 ≤1.5s（实测 0.125–0.146s；**冷调用端到端约 2.6s/页**，见 R-F3） |
| 2 | **`hl_layout.py` 断行层**（M2 前半） | `Line`、`wrap_lines()` + 与 `qa.measure_text_lines` 的等价测试 | ✅ 已完成（`959cb5c`） | 等价语料全绿（8 类 × 3 字号 × 8 行宽 = 192 条 + 300 条 fuzz；独立复验 864 组 + 3 条不变量） |
| 3 | **`hl_layout.py` 定位层**（M2 后半） | `Rect`、`Unit`、`build_units()`、`paragraph_lines()`、`page_shapes()`、`measure_coverage()` | ✅ 已完成（`66ed699`） | **v2 判据**：行级/形状级 **≥ 2.5×**（实测 2.6×）**或** pad 归零行级中位 **≥ 0.30**（实测 0.321）；框溢出率 ≤5%（实测 0/27）；**并产出 overlay 图人工确认**。**且必须声明作用域：仅单行段落**（R-F2） |
| 4 | **`hl_anim.py` 播放器**（M3） | `build_player()` | ✅ 已完成（`d2130bb`） | 30 条用例（含全部转义）+ 真机探针：`goto` 同步、降暗生效、自动讲解推进、页面无 JS 报错 |
| 5 | **`hl_anim.py` 截图**（M3 尾） | `shot_player()` | ⬜ 未开工 | 步进序列截图数 == Σ units；画面非空白（**v2：`check_exists=True` 的黑屏闸门必须先落地**，否则黑帧会进 MP4） |
| 6 | **`cli.py`**（M4） | 5 个子命令 + JSON 信封 + 退出码 + 错误码表 | ⬜ 未开工 | 四命令跑通；无 PowerPoint 时返回明确中文错误而非堆栈；`skipped` 合并两个来源；`export_width_px` 一路传到 `canvas_width_px`（**当前无断言保护，最易埋雷**） |
| 7 | **`pptx_out.py`**（M6） | `read_template()`、`build_builtin_master()`、`build_deck_pptx()`、`fit_text()` + theme zip 后处理 | ⬜ 未开工 | 导出 .pptx 在 PowerPoint 打开**可编辑**、改文字不破版、占位符可选中；中文渲染为**微软雅黑而非宋体**；`qa.check_pptx` 无 error |
| 8 | 工作台入口（M5）/ skill（M7） | `app.py` +2 路由、`index.html` +1 按钮、`skill/ppt-anim/` | ⬜ 未开工 | 上传 pptx → 浏览器可见播放器；pi 里按 SKILL.md 跑通一次全流程。**同时把 `app.py::_export_pdf_via_com` 统一到 §3.4 守卫版本（R-N4）** |

**回归红线**：
- `anim.py` / `builder.py` **一字不改**（实测已确认 `git diff c188b63 -- anim.py builder.py` 为空）；`qa.py` 不改签名或行为（只调用）。
- `tests/` 全绿。**基线口径要说清**：commit `c188b63` = **248** 条（契约所称"248 基线"指此）；实现步 1–4 后 545；再加独立用例共 **960 collected**。`KNOWLEDGE.md:53` 的"244"**已过时**。
- ⚠️ **`output/` 整个在 `.gitignore` 内** → 新克隆上 `needs_pptx_src` / `needs_material` 标记的用例（含 M1/M2 的核心验收）会**静默 skip**，所以"全绿"是本机产物、不是仓库可复现的事实。交接/CI 必须先跑 `tools/probes/accept_m1.py` 生成底图（R-N3）。

### 11.2 必须先验证再动笔的点（v2 重排）

1. **N1 重复 `Open` 语义** —— 唯一有"破坏用户数据"后果的未验项，也是 R-S1 的风险判据。**先做这条。**
2. **N2 折行素材重测 M2** —— 没有它，M2 的覆盖率结论只对单行段落成立。
3. **N3 真实手工 PPT**（需用户提供，仍是唯一阻塞项）。
4. **§5.2 deck.json 几何校验** —— 落地 CLI 前必须有，否则改坏的 deck.json 只会报 `INTERNAL`。
5. **`export_width_px` ↔ `canvas_width_px` 一致性** —— 目前零断言保护。

### 11.3 交付物自查

- [ ] 只读代码，**未修改任何 `.py`**（本次唯一产物是 `docs/PPTX_INTERFACE.md`）
- [ ] 每个引用的现有符号都能在 §1 找到出处（文件 + 符号名）
- [ ] 每个数据结构都写了字段名与单位（EMU / pt / px 标清）
- [ ] §10 的 9 条未验项都在文里显式标注，未混入"确定"叙述
- [ ] §12 的每条修订都写了三段式，且证据出处可追溯到具体报告小节 / 探针脚本 / 代码行
- [ ] §11.1 的顺序可执行：每步都有**可复现的验收命令**，且验收数字是**实测口径**（不是 v1 的估计值）

---

## 12. v2 修订记录（逐条三段式）

> 格式：**v1 原文 → 实测/审计结论（附证据出处）→ v2 修正**。
> 证据分级沿用审计报告口径：**已实测** = 有命令输出；**【静态推断】** = 只读代码/文档。

### R-D1 · COM 实例隔离（**前提被推翻**）

- **v1 原文**（§3.4-2）：「用 `win32com.client.DispatchEx("PowerPoint.Application")` 而非 `Dispatch` —— `Dispatch` 会**附着到用户正在用的 PowerPoint 实例**，随后的 `Quit()` 会关掉用户没保存的稿子。`DispatchEx` 强制新实例。」
- **实测/审计结论**：**本机 PowerPoint 16.0 上该前提不成立**。`tools/probes/verify_v8c_isolation.py`：启动前 `PIDs=[]` → `Dispatch` 建用户实例并 Add 两份稿后 `PIDs=['9992']` → `DispatchEx` 看到的 `Count` 等于用户的 `Count`（不是 0）、`PIDs` 仍是 `['9992']` → **没开第二个进程**；用户再加稿，`DispatchEx` 同步看到。三个创建 API 交叉对照（`Dispatch`/`DispatchEx`/`GetActiveObject`）+ 进程数 + "往 app1 加稿看 app2/app3 是否同步"三重指向同一实例（`docs/TEST_REPORT.md` §3 / `tools/probes/indep_d1_com.py`）。审计复核：`Dispatch` vs `DispatchEx` 走同一个多用途 COM server（`docs/AUDIT_REPORT.md` §4 D1）。
- **v2 修正**（§3.4）：删除"DispatchEx 可隔离"的全部表述，改写为「**前提：COM server 是共享的**」，安全完全落在守卫上。守卫升级为：`had_powerpoint`（Dispatch 前 `tasklist` 找 `POWERPNT.EXE`，探不到 → 当作 True）+ **PID 集合双探**（Dispatch 之后立刻记一次，退出时比对）+ `Presentations.Count == 0`，三者同时成立才 `Quit`。同时**明确写进契约**：`pre_count` 判据已被更严的"退出时 Count==0"取代（审计指出实现里 `pre_count` 赋值后从未被读取）。"绝不写 `app.Visible`"的理由**加强**：既然附着在用户实例上，`Visible=False` 会把用户正在看的窗口藏起来。

### R-S1 · `pres.Close()` 无守卫（**v2 新增，来自审计**）

- **v1 原文**（§3.4-5 伪码）：`finally:` → `try: pres.Close() except Exception: pass`（**无条件**），随后才对 `Quit()` 做 `pre_count == 0` 守卫。
- **审计结论**（`docs/AUDIT_REPORT.md` S1，**静态推断**——按审计纪律未启动 PowerPoint）：`Quit` 有两道守卫、**`Close` 一道也没有**（`pptx_io.py:645-656`，`grep -n "\.Quit()\|\.Close()"` 证实 `pres.Close()` 在 `finally` 里无条件执行）。既然 D1 已证实"我们附着在用户的实例上"，而**同一个威胁在 `Close()` 上一字未设防**：当**用户自己正开着这份 `deck.pptx`**（做演示时尤其常见）且 `Open` 返回的是已存在的那份 Presentation 时，被关掉的就是**用户的窗口**；有未保存修改时通常弹保存对话框 → 自动化里表现为挂起或静默丢弃。"安全的另一半（Quit）守住了，危险的一半（Close）裸露。"
- **v2 修正**（§3.4 规则 5）：`Open` **之前**记 `before_names = {p.Name for p in app.Presentations}`；`finally` 里**仅当 `pres.Name not in before_names`** 才 `pres.Close()`，否则只释放引用（我们的只读打开不产生未保存修改，不关也不丢数据）。**并把"PowerPoint 对已打开文件再次 `Open` 的语义"升为 §10 的 N1（最高优先级待验项）**——它直接决定 S1 是"严重"还是"无风险"。

### R-L5 · `had_powerpoint` 的采样窗口 TOCTOU（**v2 新增，来自审计**）

- **v1 原文**（§3.4-5）：守卫 = 「记 `pre_count = app.Presentations.Count`（`Open` **之前**），`pre_count == 0` 才 `Quit`」——v1 认为这足以保证安全。
- **审计结论**（`docs/AUDIT_REPORT.md` L5，**静态推断**）：`had_powerpoint = _powerpoint_running()` 在 `pptx_io.py:599`、`DispatchEx` 在 `:609`。用户恰在**这个窗口里启动** PowerPoint（还没开稿）→ 结束时 `had_powerpoint == False` 且 `Count == 0` → **用户的 PowerPoint 被 `Quit`**。D1 的加固堵住的是"事先就开着"，**没堵住"这期间打开"**。影响有限（起手界面无未保存数据），但与本模块已声明的立场（`pptx_io.py:20`「宁可留一个空进程，也不关用户的东西」）不一致。
- **v2 修正**（§3.4 规则 4 + 残余风险说明）：把"资格判定"从"启动前有没有进程"升级为 **PID 归属**——`DispatchEx` **之后立刻**再探一次 PID 集合，退出时只有"当前 PID 集合 == 我们记下的那一个"且 `Count == 0` 才 `Quit`，把窗口收窄到"我们 dispatch 之后用户才启动"这一小段。**残余窗口明确接受并写进契约**；同时记录彻底解法——**永不 `Quit`**（代价：常驻隐形进程），作为可选策略，默认不启用。

### R-F1 · `export_pages` 拆掉调用方的 COM apartment（**v2 新增，来自独立测试**）

- **v1 原文**（§3.4-6）：「**线程安全**：COM 在 worker 线程里必须 `pythoncom.CoInitialize()` / `finally: pythoncom.CoUninitialize()`。」——只要求配对，未规定"谁初始化谁反初始化"。
- **实测结论**（`docs/TEST_REPORT.md` F1）：`export_pages()` 成功返回后，同线程里调用方**在调用之前**持有的 PowerPoint 代理 → `RPC_E_DISCONNECTED`；该线程随后**任何** COM 调用都失败 `CO_E_NOTINITIALIZED`（**连 `Scripting.FileSystemObject` 都建不出来**，说明整个 apartment 被卸掉）。逐步插桩定位到**第 12 步（`CoUninitialize`）**；单独一对 `init/uninit` 不复现（5 组对照全存活）→ 触发条件是那对调用与 PowerPoint 打开/导出路径的交互，机制未定。**偶发**（3 线程探针 1/3 死亡）。审计复核确认代码根因成立（`docs/AUDIT_REPORT.md` §0）。影响面：CLI 自愈（实测同线程连续两次调用都成功）；**Flask / R5 的 >50 页后台线程**【推断】会拿到死代理且报错信息与真实原因无关。
- **v2 修正**（§3.4 规则 6）：改为「**只在本次真正初始化了**时才 `CoUninitialize`」——据 `CoInitializeEx` 的 `S_OK`/`S_FALSE` 或自记 `we_initialized` 标志。代码注释必须写"**调用方若还要在同一线程用 COM，需自行重新 `CoInitialize()`**"（v1 的实现注释写成"噪音，不用管"是错的）。另在 §10 增 N8：在 Flask 里实测该路径。

### R-L6 · 共享实例上无互斥（**v2 新增，来自审计**）

- **v1 原文**：§3.4 只要求线程内 `CoInitialize`，**未提并发**；计划 R5 又要求 ">50 页走后台 + 进度日志"。
- **审计结论**（`docs/AUDIT_REPORT.md` L6，**静态推断**）：`export_pages` 全程无锁。两次并发导出（Flask 工作台 + CLI，或 >50 页后台路径）会在**同一个** PowerPoint 实例上交错。守卫方向保守（`Count` 非 0 就不 Quit）故**不会破坏用户数据**，但 `Export` 可能因模态状态失败，且失败文案会把责任推给用户。独立测试另记录：3 线程同时驱动同一实例时 2/3 报 `Slide.Export : Object does not exist` / `Presentation.Close : Object does not exist` → **多线程并发访问同一 PowerPoint 实例不可靠**（`docs/TEST_REPORT.md` §7.4）。
- **v2 修正**（§3.4 规则 7 + §10 N9）：要求**模块级 `threading.Lock` 串行化 `export_pages`**（与 `app.py::_video_jobs` 同思路），CLI 与工作台共用。落地步 8 前处理。

### R-D2 · coverage 阈值 0.35 物理不可达（**阻塞验收的前提被推翻**）

- **v1 原文**（§4.4/§4.5）：「M2 的验收指标就是"高亮块内墨迹 coverage 中位数 ≥ 0.35"——这是覆盖率指标，tight 是唯一能稳定达标的取法」；§11.1 步 3 验收写「coverage **中位数 ≥ 0.35**」。
- **实测/审计结论**：**不可达**。`tools/probes/analyze_m2_ceiling.py`（`docs/IMPL_REPORT.md` D2）：实测中位 **0.286**，三项分解 `0.731（竖向填充：行高框比 CJK 墨迹高约 27%） × 0.960（横向填充） × 0.414（CJK 字形墨迹只占 em 盒，谁都改不了） = 0.286`，与实测**完全自洽**；**基于字体度量的任何矩形法上限 = 0.414**（rect 精确等于墨迹包围盒，而墨迹 bbox 必须先有底图 → 循环依赖）；把契约的 `pad(2.0/1.0pt)` **归零**也只到 **0.321**。独立验证用**完全不同的像素法**（外侧远环众数 + 欧氏距离；Otsu 双簇）复算：**0.287 / 0.235 / 0.286，上限 0.417、pad0 0.323**，差异 ≤0.004 → **确认不是口径 artifact**（`docs/TEST_REPORT.md` §2）。
- **v2 修正**（§4.5）：废弃绝对阈值 0.35，改为**可达表述**——**主判据（相对）行级/形状级 ≥ 2.5×**（实测 2.6×，同函数同底色自相对、跨机器稳）**或 副判据（绝对）pad 归零后行级中位 ≥ 0.30**（实测 0.321 / 独立法 0.323）；保留溢出率 ≤5%（实测 0/27）。**并把 0.414 的上限论证写进契约**，防止该阈值再被提出。同时新增两条口径纪律：① 验收脚本**必须显式传局部底色**（默认"内侧 1px 环中位"在 rect 恰为墨迹块时失明，对照图真值 1.0 → 0.0；本素材 `ring_ink ≤ 0.037` 不发作）；② 措辞更正为"**基于字体度量的**任何矩形法上限 0.414"（v1/实现报告的"任何矩形法上限"表述过宽——墨迹 bbox 需先有底图）。

### R-D8 · 图片/图表"coverage 天然接近 1"（**前提被推翻**）

- **v1 原文**（§4.6）：`picture` 行写「`rect` = 形状外框（**图片本身就是墨迹，coverage 天然接近 1**）」。
- **实测/审计结论**：**不成立**。该素材图表单元 coverage = **0.166**，是全场最低（`docs/IMPL_REPORT.md` D8；`docs/TEST_REPORT.md` §6 用独立像素法复算得 0.166）。图表内部本就大片留白，"整块一个单元"的覆盖率必然低。
- **v2 修正**（§4.6）：删除"天然接近 1"的断言，`picture` 只保留"rect = 形状外框"（覆盖率取决于图片自身留白，**无保证**）；并新增**验收统计口径**：coverage 中位数**只对文本类 Unit 的行 rect 计算**，`chart`/`picture`/`table` **单列报数**（否则 1 个 0.166 的图表会无谓拉低整页中位）。

### R-F2 · M2 的验收素材没考查断行层（**v2 新增，来自独立测试**）

- **v1 原文**（§4.4/§11.1）：把"行级 tight"当作已经成立的方案，验收只写 coverage 数字。
- **实测结论**（`docs/TEST_REPORT.md` F2 / `tools/probes/indep_wordwrap.py`）：该素材 **27 个单元 → 27 条行 rect，多行单元 0 个**；把 `word_wrap` 全部强置 `True` 后**仍是 27**（没有段落真的需要折行）。根因：`python-pptx` 的 `shapes.add_textbox()` **默认写 `wrap="none"`**（该素材 **13/22** 个文本框如此），`hl_layout` 遂按"不换行"处理，每段恒 1 行。→ 本素材上的"行级 tight"**数值上等价于段级 tight**，`wrap_lines` 那层**只被单元测试覆盖，未被验收素材碰过**。审计独立确认该结论（`docs/AUDIT_REPORT.md` §0）。
- **v2 修正**（§4.5 作用域声明 + §10 N2）：M2 的验收标准必须附**作用域声明——"仅对单行段落成立"**；要外推到会折行的真实稿，必须另备一份 `word_wrap=True` + 超宽段落的长文本素材重测（含折行后的竖向填充、pad 占比与 coverage 分解）。同时把 `add_textbox` 默认 `wrap="none"` 这条事实写进 §1.2。

### R-D9 · 验收数字口径（**v1 的"40 形状"未说明过滤后应剩多少**）

- **v1 原文**（§11.1 步 1 验收）：「对 `output/b_multislide.pptx`：读出 **10 页、40 形状**、坐标非 `None`；COM 导出 10 张 1920×1080 PNG」。
- **实测/审计结论**（`docs/IMPL_REPORT.md` 步 1「口径说明」；`docs/AUDIT_REPORT.md` §4 D9）：该稿**原始 41 个顶层形状 = 40 个文本框（`p:sp`）+ 1 个图表**；40 个文本框里 **18 个文本为空**，导入期按 §2.3 过滤规则剔除后保留 **23**（22 text + 1 chart），`skipped = 18` 全为 `empty_text`。**裸 XML 独立核实通过**（不经 python-pptx API）：顶层元素 41 → `{'sp': 40, 'chart': 1}`，对账 `保留 23 + 过滤 18 = 41`（`docs/TEST_REPORT.md` §6(a) / `tools/probes/indep_m1_counts.py`）。即：契约的"40"是文本框数、**不含图表**，且**未说明过滤后应剩多少** —— 数字对得上，只是口径要写清，否则实现者会误判为少读了形状。
- **v2 修正**（§11.1 步 1）：验收数字改写为「读出 10 页；顶层形状 **41**（**40 个文本框 + 1 个图表**），导入期过滤 **18** 个空文本框后保留 **23**（22 text + 1 chart）；坐标非 `None` 的顶层形状 **0** 个」。同批一并修正的还有该行的耗时口径（见 R-F3）。

### R-F3 · "≤1.5s/页"的口径（**v2 新增，来自独立测试**）

- **v1 原文**（§11.1 步 1 验收）：「COM 导出 10 张 1920×1080 PNG，**`≤1.5s/页` 量级**」。
- **实测结论**（`docs/TEST_REPORT.md` F3 / `tools/probes/indep_export_cost.py`；`docs/IMPL_REPORT.md` 耗时口径）：**稳态**逐页中位 **0.125s**（最大 0.343s），**达标且优于 spike 的 0.82s/页**；**附着调用** 0.23s/页；但**冷调用端到端 = 2.64s/页**（10 页 26.4s）——而冷调用正是 CLI 每次 `import` 的真实形态，且**第二次连续冷调用仍 26.45s、不摊销**（每次自起自 Quit）。成本构成：`DispatchEx` 启动 5.77s + `Open` 0.18s + 10 页 `Export` 1.47s + `Close` 0.07s + `Quit()` 立即返回但进程延迟退出（实测一次 >30s）+ `CoUninitialize` 2.67s + `tasklist` 0.47s。**实现报告的"我不确定的地方 #7"被证实为真。**
- **v2 修正**（§11.1 步 1 + §8.2）：验收口径写明为「**稳态**逐页中位 ≤1.5s」，并把冷调用成本单列（约 **2.6s/页**、10 页约 26s，且不摊销）；CLI 帮助文本要写清这一点，是否需要常驻实例方案留待步 6 决定（代价是与 §3.4 的实例隔离策略纠缠）。

### R-D3 · `font_scale` 口径不自洽（**v1 自相矛盾**）

- **v1 原文**：§2.3「`font_scale = a:normAutofit/@fontScale ÷ 1000`（`"60000"` → **60.0**，即百分数）」；§4.3「`font_scale is not None` → 所有 `size_pt *= font_scale`」。
- **实测/审计结论**：**不自洽**（`docs/IMPL_REPORT.md` D3）。直接相乘会把字号放大 100 倍（实测得 **1200pt** 而非 12pt）。实现按**存储口径**在 `hl_layout.py:287` 里 `/100` 取真实倍率，存储照 v1 不动。审计确认与报告一致（`docs/AUDIT_REPORT.md` §4 D3，用例 `test_build_units_font_scale_scales_size_and_warns` 锁定）。另：v1 的"÷1000"措辞本身也不准——`fontScale` 的单位是千分之一**百分点**，真实倍率是 `÷100000`（`60000 → 0.6`），`÷1000` 得到的是**百分数 60.0**，不是倍率。
- **v2 修正**（§2.3 + §4.3）：字段语义改为**真实倍率**——`font_scale = @fontScale ÷ 100000`（`"60000"` → `0.6`），§4.3 用法保持 `size_pt *= font_scale`（直接相乘）。**注**：这是一处**口径变更**，需在两侧之一挪动除法（现状是"存百分数 + 用时 /100"，净效果本来就对）；实现方按 v2 落字时只需改一处。同时把 V7 的实测补充写进 §4.3：注入的 `fontScale` 被 PowerPoint 打开时**重算覆盖**，该风险现实概率低于 v1 预期，属防御性处理。

### R-D4 · 标题判定第 ② 条无法从 IR 实现（**IR 缺字段**）

- **v1 原文**（§4.6）：标题命中三条任一：① 形状名含 `Title`；② `is_placeholder` 且 **`ph type ∈ {TITLE, CENTER_TITLE}`**；③ 字号 ≥ 1.3× 本页正文中位字号。
- **实测/审计结论**（`docs/IMPL_REPORT.md` D4，读代码；`docs/AUDIT_REPORT.md` §4 确认）：§2.3 的 `ShapeInfo` **只带 `is_placeholder: bool`，没有 ph type 字段** → 第 ② 条**不可实现**，实现只做了 ①③。影响面：本项目素材上第 ② 条本就不会命中（`builder.py` 全用 `add_textbox`，`is_placeholder` 全 False），**接真实 PPT（含占位符）时会有影响**（`docs/TEST_REPORT.md` §6 确认为真，影响未经实测）。
- **v2 修正**（§2.3 + §4.6）：`ShapeInfo` **新增 `ph_type: str | None`**（`shape.placeholder_format.type` 的名字，实测取值 `TITLE`/`CENTER_TITLE`/`BODY`/`OBJECT`/`PICTURE`/`SUBTITLE`…，非占位符为 `None`）；§4.6 第 ② 条改为可落地（并把 `SUBTITLE` 一并纳入）。实现方需在 `pptx_io._shape_from` 回填该字段（纯新增，零行为改动）。

### R-D5 · 合并单元格 rect 只能按单列宽算（**IR 缺 span**）

- **v1 原文**（§4.6）：`table` 行「**合并单元格**：只对 origin 单元格出 Unit、被并格跳过；x/y 由 `table_col_widths_emu` / `table_row_heights_emu` 前缀和推出」。
- **实测/审计结论**（`docs/IMPL_REPORT.md` D5；`docs/TEST_REPORT.md` §5 数值证实）：§2.3 的 `ShapeInfo` **不带 span 信息**，所以 origin 格的 rect **只能按单列宽**算（应为跨列宽之和）。实测：合并 origin 格按单列宽（201.6pt）断行得 **2 行**，按合并宽（417.6pt）应为 **1 行** → 会**多出高亮行**；若该合并格居中/右对齐，水平定位也会错。审计确认（`docs/AUDIT_REPORT.md` §4 D5）。被并格因文本为空被"非空单元格"规则天然跳过这点没问题。**可用性已实测**：`is_merge_origin`/`span_width`/`span_height` 均可用（V9 ✅）。
- **v2 修正**（§2.3 + §4.6）：`ShapeInfo` **新增 `table_spans: list[list[tuple[int,int]]]`**（`[row][col] → (span_width, span_height)`，默认 `(1,1)`）与 **`table_cell_margins_emu`**（`TableCell` 内边距未必等于形状级默认值——实现者已申报该未验证点）；§4.6 的 origin rect 改为**跨列宽之和 × 跨行高之和**。

### R-D6 · `DeckIR.pages` 与 `PageShapes` 缺桥

- **v1 原文**（§2.3/§5）：`DeckIR.pages` 每项是 `{index, bg, shapes}`（**不含画布尺寸**），而 `build_units(page: PageShapes, …)` 需要画布宽把 EMU 换算到导出底图像素空间——**v1 没有定义两者之间的过渡函数**。
- **实测/审计结论**（`docs/IMPL_REPORT.md` D6；`docs/TEST_REPORT.md` §6 确认"是必需补丁"）：实现补了 `hl_layout.page_shapes(page_dict, deck) -> PageShapes`，`build_units` 收裸 dict 且无 deck 时给明确中文报错（不静默算错）。**但审计指出**：该函数抛的是 `ValueError` 而非 `PptxError`，与 v1 §3.1"唯一的对外异常"不严格一致（`docs/AUDIT_REPORT.md` §4，低）。
- **v2 修正**（§5.1 新增专节）：正式定义 `page_shapes(page: dict, deck: DeckIR | dict) -> PageShapes` 为**唯一合法桥接入口**，写明签名、行为与"deck 缺失/字段不全 → `PptxError("IR_MISMATCH", …)`"；并要求 `hl_layout` 的对外函数统一抛 `PptxError`（不裸 `ValueError`）。

### R-D7 · `Length` 形态行距重复计入 `LINE_HEIGHT_FACTOR`

- **v1 原文**：§2.3「`line_spacing`：`Length` 绝对值需除该段字号折算成倍数再存」；§4.3「`line_h_pt = size_pt × qa.LINE_HEIGHT_FACTOR × (line_spacing or 1.0)`」。
- **实测/审计结论**（`docs/IMPL_REPORT.md` D7，实测 V6）：对"固定 30pt 行距、16pt 字号"的段落，契约算出 `16 × 1.25 × (30/16) = 37.5pt`，而 PowerPoint 实测渲染约 **31.5pt** → 偏差约 **+19%**。审计确认（`docs/AUDIT_REPORT.md` §4 D7）。根因：PowerPoint 的"固定值行距"**整体取代**行高，而不是"在行高基础上再乘系数"。本素材无显式行距故不影响 M2，**真实手工稿会踩到**。
- **v2 修正**（§2.3 + §4.3）：废弃单一 `line_spacing` 字段，拆成 **`line_spacing_mult: float | None`**（倍数）与 **`line_spacing_pt: float | None`**（Length 绝对 pt），两者互斥；§4.3 行高改为：`line_h_pt = line_spacing_pt`（**直接用，不经任何系数**）**或** `size_pt × LINE_HEIGHT_FACTOR × (line_spacing_mult or 1.0)`。并要求 `read_pages` **先判 `isinstance(ls, Length)`** 再判数值（`Length` 是 `int` 子类，先判数值会把 `Pt(30)` 存成 381000 倍行距——实现期抓到的真 bug，用例 `test_line_spacing_multiple_and_length` 锁定）。

### R-V5 · 首段 `space_before` 被 PowerPoint 忽略

- **v1 原文**（§4.3）：逐段落累加里写「段前距 | `cursor_pt += space_before_pt or 0`」——**未区分首段**。
- **实测结论**（`docs/IMPL_REPORT.md` 开工前验证 V5 ✅）：**首段 `space_before` 被 PowerPoint 忽略**——实测首行墨迹距框顶仅 **5.9pt ≈ margin_top**（3.6pt + 抗锯齿），未计入 `space_before`。该项目 `builder.py` 大量使用 `space_before=Pt(20)`，若计入则首个 rect 会系统性下偏 20pt。
- **v2 修正**（§4.3）：明确「**仅非首段生效**」，并把它从 §10 待验清单移到"已实测通过"。

### R-M1 · `a:br` 几何与 QA 分叉（**契约内部矛盾**）

- **v1 原文**：§2.3「段落文本必须走 XML：… `a:br` → `"\n"`」+ §4.3「`word_wrap is False` → `text.split("\n")`；**否则 `wrap_lines(text, size_pt, inner_w_pt)`**」+ §4.2「`assert len(wrap_lines(t, s, w)) == qa.measure_text_lines(t, s, w)`——这正是 D2「共享底层」的价值」。
- **实测/审计结论**（`docs/AUDIT_REPORT.md` M1，**已实测复现**）：`qa._tokenize` 把 `"\n"` 当作**一个 1.0em 宽的 CJK 字形**（既非空格、也非 alnum → 落 `else` 分支），**它不构成硬换行**。于是 v1 的两条要求**不可兼得**：照字面做 → **软换行被吃掉**（渲染错）；实现里加了一条"先硬拆再各自贪心"→ 渲染对，但与 qa 对同一段**差 +1 行**（实测 `'第一行\n第二行'`：`_paragraph_lines=2`、`wrap_lines=1`、`qa=1`）。关键：**§4.2 的等价断言测的是 `wrap_lines`（它确实与 qa 一致，独立复验 864 组 + 350 组 + 4000 组 fuzz 全绿），但真正产出几何的是 `_paragraph_lines`，它没有任何等价测试** → v1 声称的"防漂移闸门"**盖不住真正被调用的那个函数**。现实性边界：仓库内三份稿 `a:br` **全为 0**（44/25/51 段），**当前素材不发作**，但手工汇报稿很常见。
- **v2 修正**（§4.2 整体重写）：把硬换行**显式写进契约**，并**不修改 `qa._tokenize`**（改它会动到既有 248 基线的溢出告警行为，代价大于收益）。具体：① 折行层保持与 qa 严格同源；② 新增 `paragraph_lines(para, size_pt, box_width_pt)` 为**几何真正的入口**（先按 `"\n"` 硬拆、每段各自 `wrap_lines`）；③ 等价断言改为**两层**——`wrap_lines == qa`（无硬换行时）+ **`paragraph_lines == qa + text.count("\n")`**（把分歧写成可断言的契约，而不是让它在暗处漂移）；④ **明写已知后果**：对含软换行的段落，`qa` 的溢出告警会**少算行数**——高亮位置与溢出告警在 `a:br` 上**刻意不同**，是两侧各自的正确性优先，不是 bug。

### R-M2 · zip bomb 无闸门（**契约点名要审、但未给闸门**）

- **v1 原文**（§8.3）：「审核项（交给审计 agent）：pptx 上传的 zip bomb / 恶意 XML、`--template` 的路径穿越、CLI 参数注入。」——只把 zip bomb **列为审核项**，没有规定任何闸门。
- **审计结论**（`docs/AUDIT_REPORT.md` M2，**已实测复现**）：把 64 MiB 全零塞进 `[Content_Types].xml`（包体 111,679 B = 109.1 KB，压缩比 ≈ **601:1**），`Presentation()` 照单全收，**67,108,864 字节在内存里展开**；即使 XML 立刻解析失败（`XMLSyntaxError`），字节已展开——pptx 侧只做 `read`，**没有任何大小/比率校验**。触发点在 `[Content_Types].xml`（`Presentation()` 打开时**必读**）。可利用性边界：当前入口是本地文件/CLI，**还不是远程可达**；一旦 `app.py` 的 pptx 上传路由上线，就变成"上传一个 100 KB 文件把服务打挂"。**同批已核查无问题、不必重复劳动**：XXE / billion laughs 双重防护（库内 `resolve_entities=False` + lxml 放大系数上限，实测 0.002s 抛错）；python-pptx **零联网**。
- **v2 修正**（§3.2 步 1 + §8.3）：`read_pages` **前置一道廉价门禁**——遍历 `ZipFile.infolist()` 累计 `file_size`（zip 目录里的**声明值，不需要解压**），总量超阈值（建议 **200 MB**）或**单条** `compress_size/file_size` 比率异常（建议 **> 200:1**）→ `PptxError("PPTX_UNREADABLE", "包体异常：解压后体积 … MB")`。**零新依赖、零解压成本**。§8.3 同步规定：这是**唯一的**包体防线，上传路由**不得**绕过。

### R-M3 · `inner_w_pt <= 0` 静默丢弃形状

- **v1 原文**（§4.3）：「`inner_w_pt <= 0` → **跳过该形状（记 warning**，不产出 Unit）」——但**没给 warning 的承载通道**。
- **实测/审计结论**（`docs/AUDIT_REPORT.md` M3，**已实测复现**）：实测形状宽 100000 EMU、左右内边距各 91440 → `inner_w_pt = -6.526pt`，`build_units` 返回 0 个单元——实现做到了"不产出 Unit"，但**连一条诊断都没有**，`build_units` 也没有 `skipped` 出口。`read_pages` 会把导入期过滤写进 `deck.skipped`，**布局期丢弃一条都进不去**（CLI 的 JSON 信封里没有）。用户在播放器里只看到"某些文字没有高亮"，排查时无任何线索。
- **v2 修正**（§4.3 + §8.1）：`build_units` 增加可选收集器 `skipped: list[dict] | None = None`，追加 `{"page_index","shape_name","shape_id","reason"}`，**与 `read_pages.skipped` 同构**；§8.1 规定 CLI 的 `skipped` 字段**必须合并两个来源**（导入期过滤 + 布局期丢弃）。

### R-M4 · 缺 `p:sldSz` → 错误码从 3 退化成 5

- **v1 原文**（§3.2 步 1）：`prs = Presentation(os.path.abspath(pptx_path))`，`width_emu = int(prs.slide_width)`——**未规定任何失败处理**。
- **审计结论**（`docs/AUDIT_REPORT.md` M4，**已实测复现**）：从 `presentation.xml` 移除 `<p:sldSz>` 后 `prs.slide_width = None` → `int(None)` 抛裸 `TypeError`。该行在 `read_pages` 的 `try` **之外**（try 只包住 `Presentation()`）→ 错误码从 `PPTX_UNREADABLE(3)` **退化成 `INTERNAL(5)`**。而 §9 的文案纪律明确 `INTERNAL(5)` 的语义是"这属于 bug，请附命令与 deck.json"——对一个**用户稿子畸形**的情形给出"这是我们的 bug"，会把排查引向错误方向。同类风险还有 `prs.slides`。
- **v2 修正**（§3.2 步 2 + §9）：把 `width_emu/height_emu` 的读取**并入同一个 `try`**，或读后判空并抛 `PPTX_UNREADABLE("PPTX 缺少幻灯片尺寸定义（p:sldSz）")`；§9 增加纪律：「用户输入畸形**必须**落在 `PPTX_UNREADABLE` / `IR_MISMATCH`，**不得**退化成 `INTERNAL`」。

### R-L1 · 老 `.ppt` 被判成"已加密"（**指引不可执行**）

- **v1 原文**（§9 错误码表）：`PPTX_ENCRYPTED` 的 hint 写「请先用 PowerPoint 去掉打开密码再导出」；`PPTX_UNREADABLE` 的示例写「确认是 `.pptx`（非 `.ppt`/`.pdf`）」。
- **审计结论**（`docs/AUDIT_REPORT.md` L1）：实现的 `_classify_bad_package` 见到 OLE2 魔数（`D0 CF 11 E0 A1 B1 1A E1`）就一律判 `PPTX_ENCRYPTED`。但 **OLE2 复合文档魔数不专属"加密 pptx"**——老 `.ppt`/`.doc`/`.xls` 全是 OLE2。用户拿一份 `.ppt` 过来会被告知"已加密，请去掉打开密码"，而**那份文件根本没有密码**（【静态推断】：本机无真实 `.ppt` 样本，按格式事实推断同一分支会命中）。**更值得注意的是测试质量**：`tests/test_pptx_io.py::test_encrypted_code` 自己拼 `OLE2 头 + 512 个零` 并断言 `code == "PPTX_ENCRYPTED"`，**把这个误判锁成了期望行为**。
- **v2 修正**（§9）：`PPTX_UNREADABLE` 的 hint 改为「确认是标准 .pptx（**旧的 .ppt 请先在 PowerPoint 里另存为 .pptx**）」；并要求区分手段用**扩展名 + CFB 目录里是否存在 `EncryptedPackage` 流**（加密 OOXML 的特征），或至少 `"EncryptedPackage" in head_bytes`；改不动就退化为 `PPTX_UNREADABLE`。同时记入 §10 N5 待验（拿一份真实 `.ppt` 定论）。**注**：这条修的是契约文案与判据方向，**测试的"锁定误判"要由实现/测试角色同步修**，不在本文范围。

### R-L2 · 表格列宽/行高数组短于单元格 → 静默截断

- **v1 原文**（§2.3）：`table_cells` / `table_col_widths_emu` / `table_row_heights_emu` 三个数组，**未规定长度不一致时的行为**。
- **审计结论**（`docs/AUDIT_REPORT.md` L2，**已实测复现**）：列宽数组截短（1 列而 2 列单元格）→ 单元数从 6 变 **2**、文本只剩 `['A1','A2']`；行高数组截短 → 从 6 变 3。**无异常、无 warning**。可达性：`read_pages` 用 `len(table.columns)`/`len(table.rows)` 生成数组，长度**必然对齐**，所以正常导入路径不触发；可达路径是**被手改/被别的工具生成的 `deck.json`**。
- **v2 修正**（§5.2 新增的 `load_deck` 校验）：把"表格三组数组长度自洽"列入 `IR_MISMATCH` 校验项（见 R-L3 的同一节）。

### R-L3 · 被改过的 `deck.json` → 裸崩而非 `IR_MISMATCH`

- **v1 原文**（§5）：「`deck.json` 是**唯一真源**」——**未规定真源被改坏时的行为**；§2.1 只强调"EMU 是源，不得反向由 px 推 EMU"。
- **审计结论**（`docs/AUDIT_REPORT.md` L3，**已实测复现**）：`picture(left=None)` → `TypeError: unsupported operand type(s) for /: 'NoneType' and 'int'`；画布 `width_emu=0` → `ZeroDivisionError`。`_shape_from` 对几何字段不校验、`None` 被原样装进 `ShapeInfo`；`load_deck` 的 `except (KeyError, TypeError, ValueError)` 只包住**构造期**，而构造期不报错 → 崩在 `build_units`，最终归 `INTERNAL(5)`。
- **v2 修正**（§5.2 新增专节）：`load_deck` 与 CLI 侧对每个 page 校验：`width_emu`/`height_emu` 为正整数、每个 shape（含 `children`）的几何字段为 `int`、`len(pages)` 与底图数一致、表格三组数组长度自洽 → 违反即 `PptxError("IR_MISMATCH", …)`。

### R-L4 · 播放器不校验底图存在/新鲜 → 黑屏静默通过

- **v1 原文**（§6.1）：`build_player` 的校验只点名「页数不一致 → `IR_MISMATCH`」与「`dim` 越界 → `BAD_ARGS`」。
- **审计结论**（`docs/AUDIT_REPORT.md` L4，**【静态推断】**）：`_check_bg_path` 只回答"这个字符串是不是 out_dir 内的相对路径"，**不回答"文件在不在"**。缺图时 `<img>` 触发 `onerror` → `hl.ready` 里 `im.onload = im.onerror = res` **同样 resolve** → `fit()`/`render()` 照跑 → 截图得到 `.frame{background:#000}` 的**黑图，没有任何错误**，而步 5 的 `shot_player` 会把它直接喂给 ffmpeg。
- **v2 修正**（§6.1）：`build_player` 增加 `check_exists: bool = True`，对每个 bg 做 `os.path.isfile(os.path.join(out_dir, p))`，缺失即 `PptxError("IR_MISMATCH", "底图缺失：bg/slide_3.png，请重新执行 pptgen import")`。§11.1 步 5 验收注明"**该闸门必须先落地**，否则黑帧会进 MP4"。

### R-L7 · `measure_coverage` 用 `0.0` 重载三种语义

- **v1 原文**（§4.5）：「`return ink_count / (width*height)`；**空矩形返回 0.0**」。
- **审计结论**（`docs/AUDIT_REPORT.md` L7，**已实测复现**）：空宽矩形 → 0.0、负宽矩形 → 0.0、完全在图外 (1000,1000,10,10) → 0.0，**与"rect 里真的没有墨迹"不可区分**。若高亮定位算错到图片外（正是 M1/L3 那类 bug 的后果），度量给出的证据与"这块是空白"完全一样 → **度量本身失去了发现"定位算错"的能力**。前一轮 `docs/TEST_REPORT.md` F6 已量化底色盲点；这条补的是"退化/越界 → 0.0"。
- **v2 修正**（§4.5）：返回类型改为 `float | None`——**退化矩形与越界矩形返回 `None`**，让调用方必须显式处理，从而能把"定位算错"与"本来没墨迹"分开；部分越界先裁剪再算。

### R-F4 · `build_player` 对结构类型错误抛裸异常

- **v1 原文**（§6.1）：只规定了"页数不一致"与"`dim` 越界"两种校验。
- **实测结论**（`docs/TEST_REPORT.md` F4）：`bg_paths=None` → 裸 `TypeError: object of type 'NoneType' has no len()`；`pages_units=None` → 同；`units` 里含 `None` → 裸 `AttributeError: 'NoneType' object has no attribute 'kind'`。缓解：CLI 路径上 `load_deck` 已把结构校验前置（`except (KeyError, TypeError, ValueError) → IR_MISMATCH`），实际风险低。
- **v2 修正**（§6.1 校验表）：把"`bg_paths`/`pages_units` 非列表、或单元含 `None`"明确列为 `PptxError("BAD_ARGS", …)`，与 §8.1"永不上抛裸堆栈"一致。

### R-F5 · `bg/..%2f..%2fetc.png` 放行（**已被审计降级**）

- **v1 原文**（§6.3）：`bg` 路径的校验即 `_check_bg_path` 的"必须位于 out_dir 内"。
- **实测/审计结论**：独立测试报"未拦截（放行）"（`docs/TEST_REPORT.md` F5）；**审计复核后修正/降级**（`docs/AUDIT_REPORT.md` §3）：`_check_bg_path` 确实不解码百分号，但**所有出口都经 `_web_path`**（`%` → `%25` 再编码），解码一次只得到**字面文件名** `..%2f..%2fetc.png`，**不构成路径分隔符** → 越界不可达。**保留意见**：若将来有别的调用方直接拿 `bg` 值拼 URL，问题会活过来。
- **v2 修正**（§6.3 + §8.3）：确认为**非阻塞**，但为防御性保留一条：`_check_bg_path` **先 `unquote` 再跑同一套 `normpath` 检查**。§8.3 同步写进"deck.json 的 `bg` 校验"。

### R-N1 · §2.3 "已展开组合" 与 `children` 冲突

- **v1 原文**（§2.3）：`PageShapes.shapes` 注释写「**已展开组合**、已过滤」，但同一节的 `ShapeInfo.children` 又定义「kind == "group" 时有效（子坐标已换算为绝对 EMU）」——**自相矛盾**。
- **审计结论**（`docs/AUDIT_REPORT.md` N1）：实测 `read_pages` **保留 group 节点**（如 `fixtures_cover.pptx` 第 0 页顶层形状 = `[('text',…), ('group', 2)]`），`hl_layout.walk` 依赖这一读法。实现选了 `children` 一侧。
- **v2 修正**（§2.3）：`PageShapes.shapes` 注释改为「**组合以 `children` 嵌套表示**（不是已展开为平铺列表）；已按过滤规则过滤」。实现侧无需改动。

### R-N2 · payload 里有从未被消费的字段（**提示**）

- **v1 原文**（§6.3/§6.4）：把 `title` 写进 `CFG`，并把每个单元的文本写进 payload。
- **审计结论**（`docs/AUDIT_REPORT.md` N2）：实测模板 JS **从不读** `CFG.title` 与单元的 `t`（只读几何 `l`）。这解释了为什么"单元文本注入"在任何上下文都无落点（**是优点**），但也意味着这些字节纯属冗余（`units.json` 会明显变胖）。
- **v2 修正**（§6.4 增补警示）：保留现状（不删，零风险），但**加一条红线**——「**若将来要在播放器里显示讲解文字，`u.t` 一旦进 DOM 就是新的注入点，必须重新走一遍 HTML 文本上下文的注入审计**」。

### R-N3 · 测试数字三处不一致 + 素材被 gitignore

- **v1 原文**（§11.1）：「回归红线：每步结束跑 pytest，**248 条保持全绿**，新增用例约 20–25 条。」
- **审计结论**（`docs/AUDIT_REPORT.md` N3）：`KNOWLEDGE.md:53` 写"当前 **244** 条"（**已过时**）；`IMPL_REPORT` 报 545 passed，实际 `--collect-only` 是 **960 collected**。更关键：`output/` **整个在 `.gitignore` 内**（`git ls-files output/` 为空）→ **新克隆上** `needs_pptx_src` / `needs_material` 标记的用例**全部静默 skip**，其中就包括 M1 的两条核心验收与 M2 的两条验收 → "545 全绿"是**本机产物**，不是仓库可复现的事实。
- **v2 修正**（§11.1）：把口径写清——**248 指 commit `c188b63` 的基线**；实现步 1–4 后 545；再加独立用例共 960。并新增交接纪律：「**交接/CI 必须先跑 `tools/probes/accept_m1.py` 生成底图**，否则 M1/M2 的核心验收会被静默 skip」。`KNOWLEDGE.md` 的"244"由后续角色修正（不在本文范围）。

### R-N4 · `app.py::_export_pdf_via_com` 仍是老写法（**同仓库同威胁**）

- **v1 原文**：§3.4 的 COM 守卫只针对新模块 `pptx_io.py`，**未提**仓库内既有的 COM 调用点。
- **审计结论**（`docs/AUDIT_REPORT.md` N4）：`app.py:753-757` 用 `Dispatch` + **无条件** `pres.Close()` + **无条件** `app.Quit()`，失败一律 `return False`——正是 `pptx_io` 花两道守卫防的事。既然 D1 已证实"COM server 是共享的"，这条路径**可能关掉用户的 PowerPoint**。
- **v2 修正**（§3.4 末条 + §11.1 步 8）：写入契约作为**已知同威胁点**，要求**步 8 落地时统一**到 §3.4 的守卫版本（或至少加 `had_powerpoint` 守卫）。**本次不越权改它**。

### R-V13 · `output/templates/` 放 `.pptx` 的后果（**理由被更正**）

- **v1 原文**（§10 V13）：「建议自定义 pptx 模板改放 `output/templates/pptx/`，**避免 `custom_templates()` 把 `.pptx` 当 JSON 解析**（会崩）。」
- **实测结论**（`docs/IMPL_REPORT.md` 开工前验证 V13 ✅）：`custom_templates()` **只认 `.json` 且逐文件吞异常** → 放 `.pptx` 只会被**静默跳过**，v1 说的"当 JSON 解析崩掉"**实际不会发生**。（另：`output/templates/` 已被 `template.py` 用作 JSON 风格模板目录，实测有 `d24ceec9.json`。）
- **v2 修正**（§7.1）：**结论保留**（仍放 `output/templates/pptx/`），但**理由改为**"避免与 `template.py` 的 JSON 风格模板目录混淆，且不依赖它的静默行为"。

### R-F3-附 · 冷启动成本未拆分量（并入 R-F3，不另立修订）

- **v1 原文**（§11.1）：验收"≤1.5s/页"，未区分冷热。
- **实测结论**（`docs/IMPL_REPORT.md` §「我不确定的地方」7 + `docs/TEST_REPORT.md` F3）：首轮整轮 27.03s / 10 页含 PowerPoint 冷启动与首次打开，**未拆分**；独立测试补拆：`DispatchEx` 启动 5.77s + `CoUninitialize` 2.67s 是大头，逐页 `Export` 合计仅 1.47s。
- **v2 修正**：并入 R-F3 的 §11.1 口径改写，不另立条目。
