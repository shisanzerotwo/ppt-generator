# KNOWLEDGE · 执行手册（agent 读）

> 供 AI agent 当作项目记忆读。规则：**符号锚点**（函数名/路由名），不用行号（代码在变动）。
> 本文件取代仓库根目录曾出现的同名草稿（那份把安全项写反、模块路径写成 `output/*.py`）。

## 模块（均位于**项目根**，非 output/）
| 模块 | 职责 | 关键符号 |
|---|---|---|
| app.py | Flask + 状态机（state+lock+worker 线程） | `_start_generation` `_pause_gate` `_gen_images` `_design_and_save` |
| outline.py | 大纲生成 + 清洗 | `generate_outline` `_normalize` `_normalize_chart` |
| image_gen.py | 逐页配图 | `generate_image` |
| critic.py | 视觉校验 + 定点修改 | `review_image` `revise_target` `refine_outline` |
| html_gen.py | LLM 自主设计 HTML + 打印分页兜底 | `generate_html_deck` `_ensure_print_css` |
| builder.py | pptx 多版式导出 | `build_ppt` |
| template.py | 模板库 + 参考稿识别 + 自定义模板持久化 | `builtin_templates` `get_template_style` `analyze_reference` `_custom_path` |
| style.py | 6 套风格库 | `STYLE_LIBRARY` `decide_style` `style_guidance` |
| llm_util.py | 统一 LLM client / 运行时切模型 / 多渠道 / 配图总开关 | `llm_client` `get_model` `add_channel` `select_channel` `mask_key` `images_enabled` |
| video.py / shot.py / anim.py | 视频合成 / Playwright 截图 / 教学动画 | `build_page_args` `build_final_args` |
| pptx_io.py | pptx 读取（python-pptx）+ **COM 无窗口导出保真底图** | `read_pages` `export_pages` `powerpoint_available` `PptxError` `load_deck` |
| hl_layout.py | pptx 稿的**行级**高亮矩形（复用 qa.py 的字体度量） | `wrap_lines` `build_units` `measure_coverage` `page_shapes` |
| hl_anim.py | **高亮**播放器（与 anim.py 的逐元素揭示并存，勿混用）+ 步进截图 | `build_player` `shot_player` |
| pptx_out.py | deck.json → **可编辑** pptx（母版 + 占位符，与 builder.py 并存） | `read_template` `build_builtin_master` `build_deck_pptx` `fit_text` |
| cli.py | `pptgen` CLI（方向 B 的入口）：import/animate/video/export/deck | `main` |
| qa.py / quality.py / uploads.py / main.py | 质检 / 文档导入 / CLI | — |
| skill/ppt-anim/ | 给 agent 用的 skill（SKILL.md + 薄封装 scripts/pptgen.py） | 逻辑全在 `cli.py`，不复制 |

## 状态机
`state` 全局 dict + `lock` + worker 线程。`phase`：`idle → outline → images → designing → ready`，分步确认时出现 `review`（`state["await_step"]` 记当前暂停点）。前端每秒轮询 `/api/status`。

## 路由（要点）
- 生成：`POST /api/generate`、`POST /api/import_file`
- 状态：`GET /api/status`
- 分步：`POST /api/stepwise`（开关）、`POST /api/continue`（放行）
- 修改：`POST /api/refine`（带 `target` 为定点，不带为整篇）、`POST /api/redesign`、`POST /api/slide/<i>/text`、`POST /api/slide/<i>/image`、`POST /api/slide/add`、`POST /api/slide/reorder`、`DELETE /api/slide/<i>`
- 风格/模板：`POST /api/brand`、`GET /api/templates`、`POST /api/template/select`、`POST /api/template/analyze`、`POST /api/templates/save`、`POST /api/templates/delete`
- 模型：`GET|POST /api/models`、`/api/channels*`
- 导出：`/api/export`（pptx）、`/api/export_pdf`、`/api/export_html`、`/api/export_txt`、`/api/export_animation`、`/api/export_video`
- 历史/产物：`/api/decks`、`/api/projects`、`/api/projects/load`、`/api/artifacts`
- **pptx 高亮讲解（与本文件上面那条"AI 生成"流水线独立，状态在 `state["pptx"]`）**：
  `POST /api/pptx/import`（上传 .pptx → 后台 worker → 播放器）、`GET /pptx/<path>`（服务产物）
- 静态：`/decks` `/images` `/files` `/videos` `/animation` `/pptx`

## 生成流水线
`outline` → AI 选风格（`state["tpl_style"]` 优先，否则 `style.decide_style`）→ `_gen_images`（含 `critic.review_image` 校验闭环）→ `_design_and_save`（`html_gen`）→ 快照 `output/projects/`。

**配图总开关**：`runtime_config.json` 的 `images` 键（缺省 = 开）。关掉后 `app._without_images` 会清空 `image_prompt` 并把 `image-*` 布局降级（否则设计稿/PPTX 会空出半页图位），`_gen_image_page` 另有一道兜底闸挡住 refine 新加的提示词——即手改配置也绕不过去。生图上游不给力时（慢/质量差）用它一键切纯排版，其余流程不变。

## 已知风险
| 风险 | 现状 | 说明 |
|---|---|---|
| iframe 预览安全 | **已正确加固** | 设计稿 iframe 用 `sandbox="allow-scripts"`（**禁** `allow-same-origin`）。这是审计 M2 的**有意设计**：LLM 产物可用脚本翻页，但**不能**同源 fetch 调 API/读 cookie。**切勿按"移除限制/加 allow-same-origin"去改。** |
| 截图截到空图表 | **已修复** | 设计稿的图表柱多由 `IntersectionObserver` 驱动，纯等待不可靠（`shot.py` 曾固定 300ms → 空图表进图片/视频）。现为三重确定化：注入 IO shim 立即派发 + `screenshot(animations="disabled")` + 稳定性轮询（上限 2s）。回归测试 `tests/test_shot_settle.py`（真实浏览器，无 Chrome 则跳过）。改 `shot.py` 时别退回固定 sleep。 |
| 设计稿 HTML 布局溢出 | **无门禁** | LLM 生成的 HTML 可能标题出界、内容超出 720px（实测某时间线页：标题顶端 -9px、内容溢出 45px），而 `qa.check_pptx` 只校验 PPTX，**HTML 侧没有任何校验**。视频/截图会把裁切原样带出去。需要时按 `getBoundingClientRect` vs section 矩形做检查。 |
| HTML 导出图片引用 | 仍存在 | 导出 HTML 引用共享 `output/images/slide_N.png`；同主题重新生成会覆盖，导致旧导出稿配图错位 |
| LLM 调用超时 | 已处理 | `llm_util.llm_client(timeout)` 按用途 60/120/300s；分步确认另有 `REVIEW_TIMEOUT`（等待放行超时，**非** LLM 超时） |
| phase 竞态 | 已处理 | `_start_generation` 在锁内初始化并置 phase，避免前端早于 worker 停止轮询 |
| chart 数据 | 已处理 | `outline._normalize_chart` 会数值化 + labels/values 等长 + type 校验 |
| 定点修改并发 | 已处理 | `review` 态**不放行** `/api/refine`（会另起 worker 与阻塞的闸门竞争）；暂停期改文字用 `/api/slide/<i>/text` |

## 已知风险 · pptx 双向链路（本轮**不修**，逐条标注证据出处）

> 来源：独立测试 agent 的 `docs/TEST_REPORT.md`（F 系列）与审计 agent 的
> `docs/AUDIT_REPORT.md`（S/M/L 系列）。实现报告 `docs/IMPL_REPORT.md` 记录的是
> 实现者的自述与契约缺陷（D 系列），三者可对照读。
> **除下表外，S1（`pres.Close()` 无守卫）已修复**（commit `9833b85`，真机阳性/阴性
> 对照见 `tools/probes/s1_probe_v2.py`，防回归用例 `tests/test_s1_close_guard.py`）。

| 现状 | 风险 | 证据出处 | 说明与影响 |
|---|---|---|---|
| **未修** | F1 `export_pages` 会拆掉调用线程的 COM apartment | TEST_REPORT §6 F1；AUDIT_REPORT §2（复核代码根因） | 调用后调用方**事先持有的 COM 代理失效**，本线程后续 COM 调用报 `CO_E_NOTINITIALIZED`（连 `Scripting.FileSystemObject` 都建不出来）；重新 `CoInitialize()` 可恢复。**偶发**（频次实测 1/3）。CLI 自愈（它自己 init/uninit），**Flask worker 里"先 COM 后 export_pages"会踩**。本轮规避：`pptgen import` 在导出后不再碰 COM |
| **未修** | F2 真实素材上行级高亮退化为**段级** | TEST_REPORT §6 F2 | `b_multislide.pptx` 里 27 单元 → 27 行（多行单元 0 个），因为 python-pptx 的 `add_textbox` 默认写 `wrap="none"`（该稿 13/22 个文本框如此）。M2 的 coverage 结论**只覆盖单行段落**，不能外推到会折行的真实稿；断行层只被单元测试覆盖，没被验收素材碰过 |
| **未修** | F3 冷调用端到端 **2.6s/页**（"≤1.5s/页"只在稳态成立） | TEST_REPORT §6 F3 | 单页稳态 0.125s、附着调用 0.23s/页，但**冷调用 2.64s/页**（10 页 26.4s）且不摊销（每次调用自己起 PowerPoint、自己 Quit）。成本：DispatchEx 启动 5.77s + `CoUninitialize` 2.67s + `Quit` 延迟退出。用户体感是"每次 import 约 26 秒" |
| **未修** | M1 `a:br`（段内软换行）几何与溢出 QA **分叉** | AUDIT_REPORT §2 M1 | `qa._tokenize` 把 `"\n"` 当 1.0em 的 CJK 字形（不是硬换行），契约 §4.3 让整段交给 `wrap_lines` → 软换行被吃掉。实现加了 `if "\n" in text: 先硬拆` 修正渲染，但 `_paragraph_lines` 与 qa **差 +1 行**。契约 §4.2 那条"防漂移"等价断言只盖 `wrap_lines`，**盖不住真正产出几何的那个函数**。仓库内三份稿实测 `a:br=0`，故当前不发作；手工稿会踩 |
| **未修** | M2 **zip bomb 无闸门** | AUDIT_REPORT §2 M2 | 109 KB 的包实测让 **64 MiB 进内存**（压缩比 601:1，`Presentation()` 照单全收）。契约 §8.3 点名要审，但没规定闸门。本轮**只在上传路由加了 60 MB 上传体积兜底**（`app.MAX_PPTX_BYTES`）——那是"上传大小"，挡不住高压缩比小包；真正的闸门（`infolist()` 累计 `file_size` + 压缩比上限）应落在 `read_pages` 前置 |
| **未修** | M3 `inner_w_pt <= 0` 的形状被**静默丢弃** | AUDIT_REPORT §2 M3 | 契约 §4.3 要求"记 warning"，实现只做到"不产出 Unit"，`build_units` 也没有 `skipped` 出口 → "这里为什么没有高亮"无从排查。修法与 `read_pages.skipped` 同构（加可选收集器），成本极低 |
| **未修** | M4 缺 `p:sldSz` → 裸 `TypeError`，错误码从 3 退化成 5 | AUDIT_REPORT §2 M4 | `prs.slide_width` 可为 `None`，而 `int(...)` 那行在 `read_pages` 的 `try` **之外** → 用户稿畸形却被告知"这是我们的 bug"（`INTERNAL`），把排查引向错误方向 |
| **未修**（调用侧已规避） | L4 播放器不校验底图存在 → 黑屏静默通过 | AUDIT_REPORT §2 L4 | 缺图时 `<img>` 触发 `onerror`，`hl.ready` **照样 resolve** → 截出黑帧喂给 ffmpeg。本轮规避：`shot_player`（`video --mode step` 走它）与 `cli animate` 都先校验底图存在/`naturalWidth>0`；**播放器本身仍未校验**，自己写脚本要记得 |
| **未修**（实测印证） | F3 补充：**工作台路径实测 27s / 10 页** | 2026-09-13 复查实测 | 起 Flask 后 `POST /api/pptx/import` 上传 `b_multislide.pptx`：12:44:34 → 12:45:01 共 **27 秒**（与 F3 的冷调用 2.64s/页同级）；而 CLI 单跑 `import` 8.22s 是同进程复用下的表现。→ **不是 Flask 特有 bug**，而是"每次 import 都新建并 Quit PowerPoint"的代价；优化方向是复用实例，但需与 S1（Close 关掉用户稿）/ D1（实例共享）的风险权衡 |
| **未修** | L5 `had_powerpoint` 采样窗口 TOCTOU | AUDIT_REPORT §2 L5 | 采样在 `DispatchEx` **之前**：用户恰好在这个窗口里启动 PowerPoint（还没开稿）→ 结束时 `had_powerpoint=False` 且 `Count==0` → Quit 掉用户刚启动的 PowerPoint。D1 的加固堵住了"事先开着"，没堵住"这期间打开" |
| **未修** | L6 共享 PowerPoint 实例上**没有互斥** | AUDIT_REPORT §2 L6 | `export_pages` 全程无锁。两次并发导出（工作台 + CLI，或 >50 页后台路径）会在**同一个**实例上交错。守卫方向偏保守（少 Quit）故不破坏数据，但 `Export` 可能因模态状态失败，且失败文案把责任推给用户 |
| **未修** | L7 `measure_coverage` 用 `0.0` 重载三种语义 | AUDIT_REPORT §2 L7 | 矩形退化 / 完全在图外 / 真的没墨迹，三者都返回 `0.0`。若高亮定位算错到图片外，度量给出的证据与"这块是空白"**完全一样** → 度量丧失了发现"定位算错"的能力。建议越界/退化返回 `None` |

**同族但更轻、未修**（详见 AUDIT_REPORT §2 提示节与 L1/L2/L3）：
`app.py::_export_pdf_via_com` 已在步 8 统一到守卫版本；L1 老 `.ppt` 被判成"已加密"
（指引不可执行，且被 `test_encrypted_code` 锁成期望行为）、L2 表格列宽/行高数组短于
单元格时静默截断、L3 被改坏的 `deck.json` 裸崩而非 `IR_MISMATCH`、N3 测试数字口径
（`docs/PPTX_INTERFACE.md` 的"248 基线"与本文旧写的"244 条"均已过时；**素材在
`.gitignore` 内，新克隆上 M1/M2 的几条验收会静默 skip**，需先跑 `tools/probes/accept_m1.py`）。

## 验收
- `.venv/Scripts/python.exe -m pytest tests/ -q` 全过（当前 **1028** 条；`tests/conftest.py` 有 autouse 夹具把 `runtime_config.json` 隔离到临时目录，用例不受本地渠道/开关影响）
- **pptx 链路的验收要先有素材**：`output/b_multislide.pptx` 与 COM 底图在 `.gitignore` 内，
  新克隆上 `needs_pptx_src` / `needs_material` 标记的用例会 skip —— 跑
  `tools/probes/accept_m1.py` 生成底图即可恢复（`accept_m2/m3/m3_shot/m6` 依此类推）
- `curl http://127.0.0.1:5000/` → 200
- 端到端：选内置模板生成→设计稿含模板主色；传参考图→生成色贴近识别结果
- pptx 端到端：`pptgen import → animate → video`（CLI）或工作台「导入 PPTX 高亮讲解」
