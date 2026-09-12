# 任务卡 · 架构 agent（ppt-arch）

> 你是本计划的**架构角色**。只读代码仓库，**不修改任何代码文件**，唯一产物是 `docs/PPTX_INTERFACE.md`。

## 必读（按顺序）

1. `PLAN_PPTX_ANIM.md` —— 本期总计划（16 条已定决策 D1–D16 / spike 实测结论 / 7 个里程碑 M1–M7 / 9 条风险）
2. `KNOWLEDGE.md` —— 项目执行手册（模块表、状态机、路由、已知风险）
3. 你要设计的模块所依赖的**现有实现**（必须真读，不许猜 API）：
   - `qa.py` —— **已有 FontTools 字体度量与贪心换行实现**，是你行级高亮定位的核心复用对象
   - `shot.py` —— Playwright 截图管线（三重确定化：IO shim + `animations="disabled"` + 稳定性轮询）
   - `video.py` —— ffmpeg 配方层（`build_page_args` / `build_final_args` / `synthesize`）
   - `anim.py` —— 现有逐元素揭示播放器（**不要改它**，你设计的是并列的高亮播放器）
   - `builder.py` —— 现有 pptx 导出（保留，你设计的是并列的母版导出）
   - `llm_util.py` —— 模型/渠道运行时配置（`chat`/`vision`/`image` 三档）

## 你的任务：定清 5 个新模块的接口契约

产出 `docs/PPTX_INTERFACE.md`，必须逐项给出**函数签名 + 输入输出数据结构 + 错误语义**：

### 1. `pptx_io.py`（pptx 读取 + COM 底图导出）
- `read_pages(pptx_path) -> list[PageShapes]`：逐页形状清单
  - 必须处理：**组合形状递归展开**（`GroupShape.left` 可能为 `None`）、空文本框过滤、零尺寸过滤、隐藏形状
  - 每个形状要带：坐标(EMU)、文本、**每个 run 的字号**、段落结构、类型标记（text/picture/table/chart/group/other）
  - 说明 `GroupShape` 内子形状坐标的转换规则（子坐标相对组原点）
- `export_pages(pptx_path, out_dir, width) -> list[str]`：COM 导底图
  - **必须用绝对路径**（spike 实证：相对路径会被 PowerPoint 的 cwd 解析而报"找不到文件"）
  - 无窗口 + 只读打开；**必须保证 `pres.Close()` 与 `app.Quit()` 在异常路径也执行**
  - `ffmpeg_path()` 式的前置探测：`powerpoint_available() -> bool`，不可用时给明确中文错误与安装指引
- 定义错误类型（建议单一 `PptxError` 带中文 message，供 CLI 直接打印）

### 2. `hl_layout.py`（**行级**高亮矩形计算，本期最高风险点）
- spike 实测：文本框矩形内墨迹覆盖率中位数仅 **0.087**（p10=0.005）→ **禁止直接用 `shape.left/top/width/height` 当高亮块**
- 必须复用 `qa.py` 的度量能力（`_load_font` / `_char_width_pt` / `_measure_lines_ex` / `EMU_PER_PT` / `LINE_HEIGHT_FACTOR`）
- 接口建议：`build_units(page: PageShapes) -> list[Unit]`，`Unit` 含 `rect`（EMU 与 px 两种单位都给出）、`text`、`page_index`、`order`
- 需明确定义：文本框内边距（`margin_left/top/right/bottom`）如何计入、段落间距、行高、中英文混排的字宽取值、字号缺失时的回退值、**行矩形宽度取"整行宽度"还是"实际文字宽度"**（影响观感，给出推荐并说明理由）
- 目标：高亮块内墨迹 coverage 中位数 ≥ 0.35

### 3. `hl_anim.py`（高亮播放器）
- 与 `anim.py` 并列的独立播放器，**不改 `anim.py`**
- 交互：逐条高亮 + 其余降暗（降暗比例可配）、步进/回退/自动讲解、16:9 等比缩放
- 安全约束（**来自 KNOWLEDGE.md 已知风险表，不得违反**）：`iframe` 保持 `sandbox="allow-scripts"`，**禁止**加 `allow-same-origin`
- 必须沿用 `anim.py` 已有的**注入转义纪律**：title / deck 路径 / config 三者分属不同上下文（HTML 文本、属性、`<script>` 字符串），`</` 与 `<!--` 都要转义、占位符**单遍替换**
- 输入：底图路径列表 + 每页 `Unit` 列表 + 参数

### 4. `pptx_out.py`（母版 + 占位符导出）
- 与 `builder.py` 并列，**不改 `builder.py`**
- 内置 5 版式（封面/目录/内容/数据/尾页）的母版生成方式；自定义模板上传后的读取与填充策略
- 明确：占位符类型映射（title/body/picture/table/chart）、超长文本如何处理（截断 or 缩字号）、中文字体如何写 `a:ea`

### 5. `cli.py`（`pptgen` CLI，方向 B 的入口）
- 4 个子命令：`import` / `animate` / `video` / `export`（另加可选 `deck`）
- **stdout 输出 JSON（机器可读），人类可读信息走 stderr** —— agent 通过 bash 调用，必须易于程序化解析
- 失败一律返回**明确中文错误 + 修复指引**，不得抛裸堆栈
- 参数设计：`--mode faithful|redesign`、`--width`、`--dim`、`--sec`、`--template`、`--no-com`

## 硬性要求

- **只读代码**：不得修改任何 `.py` 文件；只写 `docs/PPTX_INTERFACE.md`
- **不许猜 API**：每个引用的现有函数/常量都必须是你真读过的（写出文件与符号名）
- 契约要**可直接交给实现者照着写代码**：数据结构给出字段名与单位（EMU / pt / px），不要只给概念
- 明确标注**你不确定的地方**（例如组合形状的坐标换算、图表能否拆数据点），列成"待实现阶段验证"清单
- 报告要短而实，不要复述计划全文

## 完成后

在 `docs/PPTX_INTERFACE.md` 末尾附一节「给实现者的清单」：按模块列出要实现的函数与顺序，以及哪些点必须先验证。
