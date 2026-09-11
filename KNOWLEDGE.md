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
| llm_util.py | 统一 LLM client / 运行时切模型 / 多渠道 | `llm_client` `get_model` `add_channel` `select_channel` `mask_key` |
| video.py / shot.py / anim.py | 视频合成 / Playwright 截图 / 教学动画 | `build_page_args` `build_final_args` |
| qa.py / quality.py / uploads.py / main.py | 质检 / 文档导入 / CLI | — |

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
- 静态：`/decks` `/images` `/files` `/videos` `/animation`

## 生成流水线
`outline` → AI 选风格（`state["tpl_style"]` 优先，否则 `style.decide_style`）→ `_gen_images`（含 `critic.review_image` 校验闭环）→ `_design_and_save`（`html_gen`）→ 快照 `output/projects/`。

## 已知风险
| 风险 | 现状 | 说明 |
|---|---|---|
| iframe 预览安全 | **已正确加固** | 设计稿 iframe 用 `sandbox="allow-scripts"`（**禁** `allow-same-origin`）。这是审计 M2 的**有意设计**：LLM 产物可用脚本翻页，但**不能**同源 fetch 调 API/读 cookie。**切勿按"移除限制/加 allow-same-origin"去改。** |
| HTML 导出图片引用 | 仍存在 | 导出 HTML 引用共享 `output/images/slide_N.png`；同主题重新生成会覆盖，导致旧导出稿配图错位 |
| LLM 调用超时 | 已处理 | `llm_util.llm_client(timeout)` 按用途 60/120/300s；分步确认另有 `REVIEW_TIMEOUT`（等待放行超时，**非** LLM 超时） |
| phase 竞态 | 已处理 | `_start_generation` 在锁内初始化并置 phase，避免前端早于 worker 停止轮询 |
| chart 数据 | 已处理 | `outline._normalize_chart` 会数值化 + labels/values 等长 + type 校验 |
| 定点修改并发 | 已处理 | `review` 态**不放行** `/api/refine`（会另起 worker 与阻塞的闸门竞争）；暂停期改文字用 `/api/slide/<i>/text` |

## 验收
- `.venv/Scripts/python.exe -m pytest tests/ -q` 全过（当前 234 条）
- `curl http://127.0.0.1:5000/` → 200
- 端到端：选内置模板生成→设计稿含模板主色；传参考图→生成色贴近识别结果
