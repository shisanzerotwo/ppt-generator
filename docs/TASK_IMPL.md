# 任务卡 · 实现 agent（ppt-impl）

> 你是本计划的**实现角色**。上游契约已由架构 agent 产出并经独立核验（717 行，含实测数据）。
> 你的职责：**严格按契约把代码写出来**，不改契约本身（发现契约定错，写进报告的"契约缺陷"节，不要自行改契约）。

## 必读

1. `docs/PPTX_INTERFACE.md` —— **接口契约（唯一权威）**。§1 可复用能力表、§2 数据契约、§3–§8 各模块接口、§9 错误码、§10 待验证、§11 实现顺序
2. `PLAN_PPTX_ANIM.md` —— 总计划（D1–D16 决策、M1–M7 验收、R1–R9 风险）
3. `KNOWLEDGE.md` —— 项目手册（**已知风险表里的红线必须遵守**）

## 执行顺序：严格照 §11.1，一步一验，不许跳步

| 步 | 模块 | 验收 |
|---|---|---|
| 1 | `pptx_io.py`（M1） | 对 `output/b_multislide.pptx`：读出 10 页 / 40 形状、坐标非 None；COM 导 10 张 1920×1080 PNG，≤1.5s/页 |
| 2 | `hl_layout.py` 断行层（M2 前半） | 与 `qa.measure_text_lines` 的**等价测试**全绿（8 类语料） |
| 3 | `hl_layout.py` 定位层（M2 后半） | 同素材 coverage **中位数 ≥ 0.35**、框溢出率 ≤ 5%，并产出 overlay 图 |
| 4 | `hl_anim.py` 播放器（M3） | 浏览器可开；页数 == 底图数；转义用例（title 含 `</title><script>`、`__CONFIG__`、路径含 `#/%`） |
| 5 | `hl_anim.py` 截图（M3 尾） | 步进序列截图数 == Σ units，画面非空白 |
| 6 | `cli.py`（M4） | `import/animate/video/export` 跑通；**无 PowerPoint 时明确中文错误而非堆栈** |
| 7 | `pptx_out.py`（M6） | 导出 .pptx 可编辑；中文渲染为**微软雅黑而非宋体**；`qa.check_pptx` 无 error |
| 8 | 工作台入口 + skill（M5/M7） | 上传 pptx → 播放器可见；`skill/ppt-anim/` |

**本次会话目标：至少完成步 1–4**（核心硬链路）。若上下文将尽，**完成当前步并 commit 后停下**，在报告里写清停在哪、下一步从哪开始 —— 不要为了"做完"而跳过验证。

## 开工前必须先做的验证（§11.2，顺序照抄）

1. **V8 COM 实例隔离**（最高危：写错会关掉用户没保存的 PowerPoint）—— 先手动开一份未保存的稿 → 跑 `export_pages` → 确认该稿仍在。先验证再写代码。
2. **V1/V2 组合形状换算 + `GroupShape.left is None`**
3. **V4/V5 垂直锚点继承 + 首段 `space_before`**
4. **V7 `a:normAutofit/@fontScale`**
5. **V13 `output/templates/` 命名冲突** —— 契约建议自定义 pptx 模板放 `output/templates/pptx/`，按此实现

### 关于 V12（真实稿素材）—— 你可以自己缓解

契约把"用户提供手工真实 PPT"列为阻塞项。**你不必等**：请**用 python-pptx 自己造**测试素材，覆盖契约未验证的形状类型：

- 含**组合形状**（`add_group_shape`）的一页
- 含**表格 + 合并单元格**的一页
- 含**图片**的一页
- 含**图表**的一页
- 含**显式行距 / 首段 `space_before` / 自动缩排（normAutofit）**的一页
- 一页**中英混排 + 超长无空格英文词**

素材生成脚本放可复用位置（如 `tools/make_fixture_pptx.py`），产物进 `output/fixtures/`，并把"自造稿能覆盖什么、覆盖不了什么（真实世界字体替换/复杂版式）"写进报告。

## 硬性约束

- **`anim.py` / `builder.py` 一字不改**（契约 §11.1 回归红线）
- **不加新依赖**（`python-pptx` / `pywin32` / `fonttools` / `playwright` / `Pillow` 都已装）
- **每步结束跑全量测试**：`.venv/Scripts/python.exe -m pytest tests/ -q` —— **248 条必须保持全绿**
- 现有模块的**公共行为不得改变**（`qa.py` 的度量函数只能调用，不要改其签名或行为；如必须改，先停下写进报告）
- 每个里程碑**独立 git commit**（信息写清是第几步）
- 测试用例放在 `tests/`，新增约 20–25 条；**COM / 浏览器相关用例要能在无相应环境的机器上 skip**（照抄 `tests/test_shot_settle.py` 的 skipif 模式）
- 中文错误信息优先；CLI 不得抛裸堆栈

## 交付物

- 代码：`pptx_io.py`、`hl_layout.py`、`hl_anim.py`、`cli.py`、`pptx_out.py`（按实际完成到哪步）
- 测试：`tests/test_pptx_io.py`、`tests/test_hl_layout.py`、`tests/test_hl_anim.py` 等
- 素材脚本：`tools/make_fixture_pptx.py`
- **进度报告：`docs/IMPL_REPORT.md`**

## 报告要求（`docs/IMPL_REPORT.md`）

Loop 工程纪律：**模型不能批准自己的完成**，每条声明必须附可复现证据。

每步写：① 做了什么（文件 + 函数）② 跑了什么命令 ③ **命令的真实输出**（关键数字，如 coverage 中位数、耗时）④ 未通过/存疑的点。

末尾固定三节：
- **契约缺陷**：契约哪里写错了或不可实现（不要自行改契约）
- **遗留与下一步**：停在哪、下一步从哪继续
- **我不确定的地方**：任何"看起来对但没验证"的断言，明确标注

不要复述计划全文，报告要短而实（每条断言都要有命令输出支撑）。
