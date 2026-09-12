# 任务卡 · 功能测试 agent（ppt-test）

> 你是本计划的**独立验证角色**。实现者已完成步 1–4（`pptx_io.py` / `hl_layout.py` / `hl_anim.py`）并自写了 297 条测试（全量 **545 passed**）。
>
> **你的职责不是复跑它的测试，而是独立地试图证伪。** 实现者自测往往覆盖"对自己有利的语料"；你要用不同方法、不同素材、对抗性输入去撞它。

## 必读

1. `docs/PPTX_INTERFACE.md` —— 接口契约（权威定义，含 §10 待验证清单）
2. `docs/IMPL_REPORT.md` —— 实现报告，**尤其「契约缺陷」节 D1–D9**（含一条高危 D1、一条阻塞验收 D2）
3. `PLAN_PPTX_ANIM.md` —— M1–M3 的验收标准
4. `KNOWLEDGE.md` —— 项目已知风险红线

## 硬约束

- **只新增 `tests/` 下的文件**（建议 `tests/test_pptx_independent.py` 等）
- **不改实现代码**（`pptx_io.py` / `hl_layout.py` / `hl_anim.py`）；发现 bug 写进报告，**不要动手修**（修是别的角色的事）
- **不改** `anim.py` / `builder.py` / `qa.py` / `video.py` / `shot.py` / `template.py`
- **不改实现者已有的测试文件**（`tests/test_pptx_io.py` / `test_hl_layout.py` / `test_hl_anim.py`）
- 你新增的用例必须**自己跑通**；全量 `pytest tests/ -q` 结束时仍须全绿
- 依赖 COM / 浏览器的用例，必须在缺失环境下 **skip**（照抄 `tests/test_shot_settle.py` 的 skipif 模式）

## 必须独立验证的点（按价值排序）

### 1. M2 的 coverage 结论要用**独立实现的方法**交叉验证
- 实现者报"行级 coverage 中位 0.286、理论上限 0.414"，并据此判定契约阈值 0.35 不可达
- **不要复用 `hl_layout.measure_coverage`**：自己写一套像素统计（例如：取墨迹包围盒的紧致矩形 / 直接对所有前景像素求凸包或 bbox / 换一种底色估计算法）重算，看能否复现 0.286 与 0.414
- 若复现不出 → 它的结论可能是口径artifact，必须报告

### 2. D1（高危）· COM 实例隔离
- 独立复现"`DispatchEx` 与 `Dispatch` 拿到同一实例"（不同探针写法，别照抄它的）
- 验证它加固后的守卫（`pre_count` + POWERPNT.EXE 进程数）真能保护这个场景：**用户开着 PowerPoint 但没打开任何稿**（此时 `pre_count == 0`，只按契约会关掉用户的 PowerPoint）
- 验证"绝不触碰 `app.Visible`"这条红线在代码里确实成立

### 3. M3 的转义安全（对抗性）
自己构造恶意输入去撞播放器，至少覆盖：
- `title` = `</title><script>alert(1)</script>`
- `title` 含 `__CONFIG__` / `__DECK__` 字样（针对单遍替换）；含 `<!--`
- 底图路径含 `#` `%` 空格 中文（URL 编码）
- 断言产物 HTML 里不出现可执行注入；能用真实浏览器则验证无 `pageerror`

### 4. 边界与异常路径
- 0 页 pptx、非 pptx 文件（.docx/.txt 改名）、损坏 zip、超大文件
- 空页（无任何形状）、全空文本框、单字宽度超过行宽的极端文本、>1000 字超长段落
- 无 PowerPoint 环境下的错误路径（用打桩/monkeypatch 模拟 `powerpoint_available()` 返回 False，断言给的是**明确中文错误**而非裸堆栈）

### 5. 一致性与回归
- `hl_layout.wrap_lines` 与 `qa.measure_text_lines` 的等价性：用**你自己挑的**语料（含实现者可能没测的：连续空格、全角标点行首、emoji、制表符、`\r\n`、纯数字、超长英文单词+中文混排）
- 确认 `anim.py` / `builder.py` 的既有行为未受影响

### 6. M1 的验收数字
- `pptx_io.read_pages` 对 `output/b_multislide.pptx` 的数量口径（实现者报告 D9：原始 41 形状 = 40 文本框 + 1 图表，过滤空文本后保留 23）——独立核对这个口径
- COM 导出：页数、尺寸、每页耗时

## 交付物

`docs/TEST_REPORT.md`，要求：

- 每条验证写：**测什么 / 命令 / 真实输出 / 判定**（通过 / 失败 / 存疑）
- **发现的问题**按严重度排序，每条附**最小复现步骤**
- 单列一节「**我无法验证的**」（环境限制、权限、需要用户素材等）
- 明确区分「我实测到的」与「我从代码推断的」—— 后者必须标为推断

## 纪律

- **不许只看代码就下结论**：每条断言必须有命令输出支撑
- 发现实现有问题 → **写报告，不改实现**
- 你的价值在于找到实现者没想到的失败路径，而不是确认它对
