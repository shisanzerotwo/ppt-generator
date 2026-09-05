# ppt-generator 补充优化 · 计划开发文档

> 版本：v3.1（2026-09-05）｜ 基线：commit `bd9ef5c`（142 tests passed, 4.87s）
> 原则：**不过度修改**——每项改动外科手术式、独立可验证、保持现有 142 条测试全绿、不动 Claude 已落的模板导入/两栏布局语义。

## 一、背景

对参考项目 `yuyuanweb/ai-ppt` 的三份源码级研究笔记（工作流/LLM 层、导出门禁、解析与大纲）确认了四个架构级差距：页级并发、导出质量门禁、内容自纠、文档溯源。本项目走「LLM 直接设计 HTML」路线，不搬它的重型实现（React 编辑器/布局 solver/三分离模型），只移植**与框架无关的零件**，并按本项目尺度做轻量适配。

## 二、本轮实施范围（本次开发）

| # | 阶段 | 内容 | 改动文件 | 分工 |
|---|---|---|---|---|
| P1 | 公共件抽取 | `llm_util.py`：ZhipuAI client 构造 + JSON 提取，critic/template 去重 | llm_util.py(新), critic.py, template.py | 主线 |
| P2 | 低垂果实 | ① 各阶段耗时打点进日志；② deck 级重复页检测（bigram Jaccard，warning 不阻断）+ `/api/quality`；③ 配图结果缓存（md5(prompt+title)，同 key 不再花生图钱）；④ 内容过瘦检测 | app.py, quality.py(新) | 主线 |
| P3 | 页级并发生图（轻量版） | `_gen_images` 改 ThreadPoolExecutor(3)，单页逻辑抽函数，锁语义不变。**决策：不引 ARQ/Redis**——单进程工具引独立 worker 进程+消息队列属过度工程；SSE 端点暂缓（前端轮询已够用，避免造无人消费的接口） | app.py | 主线 |
| P4 | 主题 CSS 变量 + 导出同源 | html_gen 提示词硬性要求 `:root{--bg/--fg/--accent/--muted}` 设计令牌并全程引用；新增 `POST /api/theme` 直改设计稿 :root 块实现**毫秒级换色**（无 :root 变量的旧稿返回 409 提示重生成）；`/api/export_html` 设计稿存在时直接同源返回，不再走第二套模板 | html_gen.py, app.py, templates/index.html(轻) | 主线 |
| P5 | PDF 入口 | `uploads.py`：pdfplumber 逐页提取，加密件试空密码、扫描件明确报错；`_parse_upload` 加 .pdf 分支 | uploads.py(新), tests/test_uploads.py, app.py(接线) | 子agent A |
| P6 | 导出质量门禁 | `qa.py`：① FontTools 字形宽度 + 贪心换行模拟（字体缺失降级估算），检测文本框溢出（warning）与形状越界/页数不符（error）；② 导出后回读验证；`builder.py`：所有 run 写 `a:ea` 中文字体 + 关闭 spAutoFit；导出路由接线：error → 422 拒发并附报告，warning → 放行提示 | qa.py(新), builder.py, tests/test_qa.py, app.py(接线) | 子agent B |

**明确不做（防过度修改）**：ARQ/Redis/SSE 完整版（P3 决策）、LangGraph 自纠环（LLM 成本，见路线图）、refine 工具调用化、预览内点击编辑、Docker/多用户。这些进路线图不进本轮。

## 三、后续路线图（本次不做，按需启动）

1. **B 单页自纠环**：critic.review_page（vision/文本审单页内容）+ issues 输入字段喂回改写，封顶 2 轮——等 Agnes 余额恢复后再开（省钱优先）。
2. **SSE + ARQ 完整版**：若将来多任务/断线重连成为真实痛点，再升级 P3。
3. **F3 refine 工具调用化**（副本+diff、locked 剔除、错误回灌、轮次封顶）、**F4 预览内点击编辑**、**演讲者备注**、**revision 回滚**、**vision 全稿审稿**（与逐页 PNG 导出共用截图基建）。

## 四、验证标准

- 全量 pytest 保持通过（142 + 新增用例：uploads/qa/quality/theme/concurrency）。
- 并发生图：日志显示页完成顺序乱序（非 1→n 串行）即生效；`/api/quality` 对重复页样例返回 warning。
- 主题换色：生成含 :root 变量的稿后调 `/api/theme`，文件内容变化且预览即时生效。
- 导出门禁：构造超长标题页导出返回 422+报告；正常稿 0 error 放行；导出文件 python-pptx 可重开、页数一致。
- PDF：文本型 PDF 走通到大纲；扫描件/加密件返回明确中文错误。

## 五、风险与回退

- 每阶段独立 commit，出问题按 commit 回退。
- builder.py 改动仅限 `_set_text` 系（a:ea + auto_size），版式几何零改动。
- 锁语义：并发生图沿用现有 `lock` 粒度（页内写状态均持锁），不新增共享可变结构。
- 缓存 key 含 title，避免同 prompt 异页文案错配；缓存目录 `output/images/cache/`，gitignore 已覆盖 `output/`。

## 六、流程（Loop 三段式）

实现（主线 + 子agent A/B 并行）→ 集成 → 全量 pytest → commit → 三个分析子agent 并行（测试成果分析 / 安全测试 / 代码审计）→ 处理发现项 → 收尾汇报（微信机器人通道同步）。

## 七、执行记录（2026-09-05 实际结果）

- **实现**：P1–P6 全部落地，测试 142 → 182 全绿（+40：qa 15 / quality 5 / uploads 3 / theme 4 / export_sec 2 / wiring 11）。
- **安全测试**（101 项离线用例）：发现并修复 4 项——简版导出 img src 注入（中，白名单+转义）、html_path 编码名不一致（unquote 三处）、项目名盘符逃逸（拦冒号）、PDF 提取异常未包装（统一 ValueError）。
- **代码审计**：发现并修复 3 中 + 7 低——qa 漏报段间距、缓存目录导入期固化（测试污染）、projects_load 漏 unquote、缓存原子写/读降级/写降级、theme 线程唯一 tmp + hex 右边界、缓存 key 掺 points、打点排除 gate 等待。
- **测试成果分析**：确认模块层覆盖扎实；接线层缺口已按清单补 11 条薄用例闭合（test_wiring.py）。
- **未做（记录在案）**：L4 上传大小上限、L5 错误信息脱敏、L8 outline 收敛第 4 份 client、L9 review 态换图竞态（既有）、缓存 LRU 上限、图表中文 a:ea——均为低危，见两份分析报告。
- 提交序列：`dab55cf` 优化批次 → `565c170` theme 测试 → `d4cdf29` 安全修复 → `8f5d052` 审计修复 → 本条（wiring 测试 + 缓存写降级）。
