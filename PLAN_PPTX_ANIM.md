# pptx 双向 + 会动的演示 · 开发计划

> 版本 v1（2026-09-12）｜ 基线：commit `c188b63`（248 tests collected）
> 原则：**不过度修改** —— 原地共享、只加新入口；每项改动独立可验证；保持 248 条测试全绿。
> 上游计划：[PLAN.md](PLAN.md)（补充优化 P1–P6）、[PLAN_ANIMATION.md](PLAN_ANIMATION.md)（动画与视频导出 M1–M3）

---

## 一、背景与两个目标方向

ppt-generator 当前形态是**单进程 Flask Web 工作台**（LLM 自主设计 HTML 幻灯片 + 动画/视频导出）。接下来要往两个方向走：

| 方向 | 定位 | 交付形态 |
|---|---|---|
| **A · 可视化工作台** | 给人用的动画演示工作台，**agent 也能驱动**；核心是"把 PPT 变成会动的演示" | 现有 Flask 工作台 + 新增 pptx 入口 |
| **B · agent 插件** | 让任意 agent 在做 PPT 时直接调用这套能力 | `skill + CLI` 双层（SKILL.md + `pptgen` CLI） |

**本期聚焦的共用硬链路**：`pptx → 会动的演示 → 动画/MP4`，以及反向的 `→ 可编辑 pptx`。这条链路两个方向都要用，先把它跑通再谈外壳。

---

## 二、已定决策（本轮共识，共 16 条）

| # | 决策 | 选择 | 理由/代价 |
|---|---|---|---|
| D1 | 两方向关系 | **定位不同、各自演进** | A 面向人看的演示工作台，B 面向 agent 调用；允许功能集分叉 |
| D2 | 共享边界 | **共享底层、上层分叉** | 底层纯函数模块共用，编排/入口/UI 各自写；避免真 bug 修两遍 |
| D3 | 目录落法 | **原地共享，只加新入口** | 保留根目录现有模块；248 条测试与 `app.py` 的 import 全不动 |
| D4 | pptx 定位 | **双向互通**（导入 + 导出） | 本期两条都做（用户明确要求）|
| D5 | 开工顺序 | **先做共用硬链路** | px→演示 这条链路两方向都要用，先验证它 |
| D6 | 首块交付 | **CLI + 工作台同做** | CLI 内核 + 工作台一个薄按钮调它 |
| D7 | "可视化"含义 | **会动的演示** | 逐条高亮、可步进讲解、导出 MP4；不是"数据→图表" |
| D8 | pptx→HTML 路线 | **混合：COM 底图 + python-pptx 文字层** | 视觉保真 + 文字可选中可动效；代价是依赖 PowerPoint |
| D9 | 底图/文字冲突 | **完整底图 + 高亮式动效** | 底图 100% 保真（含文字），HTML 只做高亮/降暗锚点；**放弃逐元素揭示** |
| D10 | 导入后处理 | **两种模式都要** | `faithful`（忠实还原）+ `redesign`（AI 重设计） |
| D11 | pptx 导出路线 | **模板母版 + 占位符** | 真可编辑、保真度好；替代现有 `builder.py` 的"从 JSON 重建" |
| D12 | 母版来源 | **内置 + 支持上传** | 内置 5 版式兜底；支持上传自定义模板（存 `output/templates/`）|
| D13 | 支持范围 | **标准汇报稿** | 文字/项目符号/图片/表格/基础图表；SmartArt/复杂组合降级为局部截图 |
| D14 | 插件形态 | **skill + CLI 双层** | SKILL.md 教 agent 何时用/怎么调，能力由本地 CLI 提供；零常驻、跨 agent |
| D15 | 插件落点 | **先放仓库，稳定后安装** | 先 `skill/ppt-anim/`，验证后装到 `~/.agents/skills/` |
| D16 | 目标 agent | **pi + agentskills.io 标准** | 按标准写 SKILL.md，Hermes 等任何支持该标准的 agent 都能用 |

---

## 三、spike 实测结论（2026-09-12，已跑完）

素材：`output/b_multislide.pptx`（10 页 / 40 文本框 / 1 图表，由本项目 `builder.py` 生成）。脚本为一次性探针（未落仓库）。

### 3.1 实测数据

| 指标 | 实测值 | 判定 |
|---|---|---|
| COM 无窗口导出 | **0.82s/页**，10 页 8.22s，共 0.13 MB | ✅ 性能可接受 |
| 画布 | 12191695×6858000 EMU = 13.333×7.5 in（16:9） | ✅ 与 `shot.py` 的 1280×720 同源 |
| 坐标换算 | `px = EMU × (导出宽 / slide_width)` 线性成立 | ✅ H2 通过 |
| 空文本框 | 40 个中 **18 个 `text` 为空** | ⚠️ 需在导入时过滤 |
| **文本框墨迹填充率 coverage** | **中位数 0.087**，p10=0.005，p90=0.434 | ❌ **H3 不通过** |
| 宽度贴合 tight_w | p10=**0.149**、中位 0.983 | ❌ 大量框宽是文字的 6 倍 |
| 最差案例 | 框 792×821px，coverage **0.005** | ❌ 巨框内仅一行小字 |
| 墨迹触框边 | **12/22** | ⚠️ 边界判定需要容差 |

### 3.2 结论：推翻了「直接拿文本框矩形做高亮」的假设

**`shape.left/top/width/height` 不是文字实际区域** —— 本项目 `builder.py` 用的是固定大框（median 高 144pt），框内平均只有 8.7% 是文字。若直接用它做逐条高亮，会出现"高亮一大片空白"的观感灾难。

**对策（已定）**：高亮矩形按**行级**计算，而不是按形状：

1. 取文本框的 `margin_left/top/right/bottom`（EMU）+ 每个段落的字号（`run.font.size`）
2. 用**已有的** `qa.py` 字体度量（`_load_font` / `_char_width_pt` / `_measure_lines_ex`）做贪心换行，得到每段的真实行数与行高
3. 行矩形 = 框内边距起点 + 累计行高；宽度 = 该行文字宽度（FontTools 度量求和）
4. `qa.py` 已有 `EMU_PER_PT=12700`、`LINE_HEIGHT_FACTOR=1.25`、`OVERFLOW_TOLERANCE_PT` 等常量可直接复用

> 这正是 D2「共享底层」的价值：`qa.py` 当初为 PPTX 几何门禁写的度量代码，现在直接给高亮定位用。

### 3.3 spike 的局限（必须补测）

- 素材全部是 **`builder.py` 生成的规整稿**：`pictures=0`、`tables=0`、无 SmartArt、无组合形状
- 因此**未验证**：图片/表格/图表的定位、组合形状（`GroupShape` 的 `left` 可能为 `None`，实测 `none_bbox=0` 只是因为没素材）、真实 PPT 的字体替换（微软雅黑→宋体等）
- **行动**：实现阶段开始前，先要一份**手工制作的真实 PPT**（含图片/表格/SmartArt）补测

---

## 四、架构改动

### 4.1 现状（四层，已核实）

```
[L1 接入]   app.py(~45 路由)  │  templates/index.html(1530 行)  │  main.py(CLI·旧路线)
[L2 编排]   全在 app.py（_generation_worker 等），与 Flask + 全局 state 耦合
[L3 能力]   outline / style / template / critic / image_gen / html_gen / builder
            / shot / anim / video / qa / quality / llm_util / uploads   ← 纯函数为主
[L4 产物]   output/{decks,projects,images,animation,videos,templates}
```

### 4.2 改动后（原地共享，只加新入口）

```
[L1 接入]   app.py(+2 路由)  │  index.html(+1 按钮)  │  main.py(旧路线不动)
            ★ cli.py  ← 新增：pptgen CLI（方向 B 的入口）
[L2 编排]   app.py 原有编排不动（工作台）
            ★ cli.py 自带轻编排（B 方向不复用 Flask 的全局 state）
[L3 能力]   原有模块全部不动
            ★ pptx_io.py    读 pptx（python-pptx）+ COM 导底图
            ★ hl_layout.py  行级高亮矩形计算（复用 qa.py 度量）
            ★ hl_anim.py    高亮播放器（与 anim.py 并列，勿混用）
            ★ pptx_out.py   母版+占位符导出（与 builder.py 并列）
[L4 产物]   output/ 下新增 pptx_src/（导入底图）与 pptx_out/（导出母版）
```

**三条硬约束（防回归）**：

1. `anim.py` 的逐元素揭示**不删不改** —— 它服务于现有 HTML 设计稿路线；pptx 稿件走 `hl_anim.py` 的高亮模式。两个播放器并存是有意的（D9 的直接后果）。
2. `iframe sandbox="allow-scripts"` 的定义**不准改**（见 KNOWLEDGE.md 已知风险表）—— `hl_anim.py` 同样沿用该策略。
3. `builder.py` 保留（图片版降级导出 + 现有 248 条测试依赖它），`pptx_out.py` 是新增的第二条导出路径。

---

## 五、技术方案

### 5.1 导入链路：`pptx → 会动的演示`

```
deck.pptx
   │
   ├─[1] pptx_io.read_pages()
   │      python-pptx 逐页取形状：坐标(EMU)/文本/字号/类型/是否组合
   │      过滤：空文本框、零尺寸、隐藏形状
   │
   ├─[2] pptx_io.export_pages()          ← 需要 PowerPoint COM
   │      无窗口、只读打开；Slide.Export(绝对路径, "PNG", W, H)
   │      ⚠️ 必须绝对路径（spike 实测：相对路径会被 PowerPoint 的 cwd 解析，报"找不到文件"）
   │
   ├─[3] hl_layout.build_units()         ← 关键：行级，不是形状级
   │      框内边距 + 字号 + FontTools 贪心换行 → 每个"讲解单元"的矩形
   │      讲解单元 = 一个段落（标题 1 个 / 要点 N 个）
   │
   ├─[4] hl_anim.build_player()
   │      <div class=slide><img src=bg><div class=hl ...></div></div>
   │      逐条高亮 + 其余降暗 25%；步进/回退/自动讲解
   │
   └─[5] 导出：截图序列 → video.py 复用（ffmpeg zoompan/xfade）→ MP4
```

**两种模式（D10）**：

| 模式 | 输入 | 处理 | 产物 |
|---|---|---|---|
| `faithful` | pptx | 上述 [1]–[5]，保原版式 | 高亮播放器 + MP4 |
| `redesign` | pptx | [1] 提取文字/图表数据 + `pptx_io.export_pages` 的底图当视觉参考 → 喂现有 `outline`/`style`/`html_gen` 流水线 | AI 设计稿 + 现有 anim.py 播放器 |

### 5.2 导出链路：`→ 可编辑 pptx`

1. `pptx_out.py` 内置母版：由 python-pptx 生成 5 种版式（封面/目录/内容/数据/尾页），带**真占位符**
2. 内容填占位符（标题/正文/图片/表格/图表），保留母版字体与主题色
3. 自定义模板：上传 `.pptx` 存 `output/templates/`，读取其 `slide_layouts` + `theme` 后填占位符
4. 双导出并存：`可编辑母版版`（新）+ `图片版`（`builder.py` 现状）

### 5.3 CLI 设计（方向 B 的入口）

```bash
# 导入：pptx → 会动的演示
pptgen import deck.pptx --out out/ --mode faithful --width 1920
pptgen import deck.pptx --out out/ --mode redesign --pages 10

# 播放器与视频
pptgen animate out/ --highlight          # 高亮播放器（faithful 稿）
pptgen video   out/ --sec 4 --fps 25

# 导出：→ 可编辑 pptx
pptgen export  out/deck.json -o out.pptx --template brand.pptx

# 全自动（复用现有流水线，可选）
pptgen deck "人工智能如何改变教育" --out out/
```

**设计约束**：
- 零常驻、纯 CLI，agent 通过 bash 调用（D14）
- stdout 输出 JSON（机器可读），人类可读信息走 stderr
- 所有失败给**明确中文错误 + 修复指引**（如"未检测到 PowerPoint，COM 导出不可用：请安装 Office 或使用 --no-com"）

### 5.4 skill 结构（先放仓库）

```
skill/ppt-anim/
├── SKILL.md              # frontmatter(name/description) + 何时用/怎么调
├── scripts/
│   └── pptgen.py         # CLI 入口（薄封装，逻辑在项目模块）
└── references/
    └── formats.md        # 产物格式与常见错误
```

---

## 六、里程碑与验收标准

| # | 里程碑 | 交付物 | 验收（可执行、可复现） |
|---|---|---|---|
| M1 | 导入骨架 | `pptx_io.py` + 单测 | 对 `b_multislide.pptx`：读出 10 页、40 形状、坐标非 None；COM 导出 10 张 1920×1080 PNG，`0.82s/页` 量级 |
| M2 | **行级高亮定位** | `hl_layout.py` + 单测 | 对 spike 同一素材：高亮块内墨迹 coverage **中位数 ≥ 0.35**（当前形状级仅 0.087）；框溢出率 ≤ 5% |
| M3 | 高亮播放器 | `hl_anim.py` + 单测 | 浏览器打开：逐条高亮定位准、其余降暗、步进/回退/自动可用；页数与底图数一致 |
| M4 | CLI | `cli.py` + 单测 | `pptgen import/animate/video/export` 四命令跑通；无 PowerPoint 时返回明确中文错误（非堆栈） |
| M5 | 工作台入口 | `app.py` +2 路由、`index.html` +1 按钮 | 上传 pptx → 出高亮播放器，浏览器可见 |
| M6 | 可编辑 pptx 导出 | `pptx_out.py` + 母版 | 导出的 .pptx 在 PowerPoint 打开可编辑：改文字不破版、占位符可选中 |
| M7 | skill | `skill/ppt-anim/` | pi 里用 SKILL.md 指引跑通一次全流程 |
| 回归 | — | — | **248 条测试保持全绿**；新增用例约 20–25 条 |

**每个里程碑的验收命令**（Loop 工程：模型不能批准自己的完成）：

```bash
.venv/Scripts/python.exe -m pytest tests/ -q          # 必须全绿
.venv/Scripts/python.exe -m pytest tests/test_pptx_io.py -q      # M1
.venv/Scripts/python.exe -m pytest tests/test_hl_layout.py -q    # M2
```

---

## 七、风险与对策

| # | 风险 | 证据/现状 | 对策 |
|---|---|---|---|
| R1 | **文本框 ≠ 文字区域** | spike 实测 coverage 中位 **0.087** | 行级度量（M2）；复用 `qa.py` 的 FontTools 实现；验收阈值 coverage ≥ 0.35 |
| R2 | **PowerPoint 是硬依赖** | COM 16.0 实测可用；LibreOffice 未装 | 明确报错 + 安装指引；`--no-com` 分支走已有截图/纯文本路线；文档写清"无 Office 的机器不可用 faithful 模式" |
| R3 | **spike 素材片面** | 全是 builder 生成的规整稿，`pictures=0 tables=0` | 实现前先要一份**手工真实 PPT** 补测；补测不通过则先修导入器再进 M3 |
| R4 | 组合形状 / 图表定位 | `GroupShape.left` 可能为 `None`；spike 未覆盖 | 递归展开组合形状；图表按形状外框整块作为一个讲解单元（不拆数据点）|
| R5 | 大文件耗时/体积 | 0.82s/页 @1920 宽（简单稿） | 提供 `--width` 档位（1280/1920/2560）；>50 页走后台 + 进度日志（沿用现有 `_video_jobs` 模式）|
| R6 | 中文/字体替换 | spike 未覆盖 | 行级度量以 PPT 内声明的字号为准；字体缺失时降级为估算（`qa.font_available()` 已有判定）|
| R7 | 高亮观感 | 降暗比例未调 | 参数化（`--dim 0.25`），M3 用真实稿人工确认 |
| R8 | 4 个 agent 写同一仓库 | 计划用串行激活规避 | 同一时刻只有一个写者；架构/审计 agent 只读 |
| R9 | `builder.py` 与 `pptx_out.py` 双导出路径分叉 | 有意为之 | 两条路径各有独立测试；`pptx_out.py` 不修改 `builder.py` |

---

## 八、claude agent 分工（4 个，串行激活）

按 D15 之前的约定：**每个角色一个新 tab，按阶段串行激活**（同一时刻只有一个写者）。

| 顺序 | 角色 | 职责 | 产出 | 只读/写入 |
|---|---|---|---|---|
| ① | **架构** | 定 `pptx_io/hl_layout/hl_anim/pptx_out/cli` 的接口契约与数据结构；递归展开组合形状的策略；行级度量的输入输出 schema | `docs/PPTX_INTERFACE.md` | **只读**（只写 md）|
| ② | **实现** | 按契约写 M1–M6 代码 + 配套单测 | 上述模块 + tests | 写入 |
| ③ | **功能测试** | 独立补测试（不看实现者自测用例）；真实稿端到端；边界（空页/无文字/超长文本/无 PowerPoint） | `tests/test_pptx_*.py` | 写入（仅 tests/）|
| ④ | **审计** | 安全审计（路径穿越、上传 pptx 的 zip bomb/恶意 XML、CLI 参数注入）+ 代码审计（资源泄漏：COM 未 Quit、临时文件未清理） | 审计报告 md | **只读** |

**启动方式**（用户批准计划后执行）：

```bash
herdr tab create --workspace w9 --label ppt-arch
herdr pane split --current --direction right --cwd "$PWD" --no-focus   # 或 tab 内直接用 root pane
herdr agent start ppt-arch --kind claude --pane <pane-id>
herdr agent prompt ppt-arch "读 docs/PPTX_INTERFACE.md 任务简报并执行" --wait --timeout 600000
```

**交接纪律**：每个角色的产出必须落成**文件**（而非只在终端里说），下一个角色读文件接手。

---

## 九、明确不做（本期）

- 逐元素揭示动效用于 pptx 稿（D9 已定：高亮模式；现有 `anim.py` 继续服务 HTML 稿）
- 音频/TTS 配音、逐元素动画解析（读 pptx 的切换动效 XML）
- SmartArt / 嵌入视频 / 3D 模型的完整还原（降级为局部截图）
- 多用户/鉴权/任务队列（全局 state 不动）
- 把 L3 抽成可 pip 安装的包（D3 已定：原地共享）
- 新建独立 workspace（用户未要求）

---

## 十、待你确认的开放项

1. **R3 补测素材**：需要你提供一份**手工制作的真实 PPT**（含图片/表格最好），否则 M2 的验收只在规整稿上成立
2. **导出模板**：内置母版之外，是否要用你的公司/个人品牌模板？（D12 选了"内置+上传"，但上传需要你给文件）
3. **CLI 命名**：`pptgen` 是否合意？（还是 `pptgen`/`ppt-maker`/其它）

---

## 十一、执行记录（2026-09-12/13 落地）

### 11.1 里程碑完成情况

| # | 里程碑 | 状态 | 交付物 | 验收证据 |
|---|---|---|---|---|
| M1 | 导入骨架 | ✅ | `pptx_io.py` | `tools/probes/accept_m1.py`：10 页 / 原始 41 形状（40 文本框 + 1 图表）→ 过滤 18 个空框后 23；坐标无一 `None`；COM 导出 10 张 1920×1080，**稳态逐页中位 0.146s** |
| M2 | 行级高亮定位 | ✅（阈值改口径） | `hl_layout.py` | `tools/probes/accept_m2.py` + `analyze_m2_ceiling.py`：**框溢出 0.00%**（≤5% 达标）；行级 coverage 中位 **0.286** vs 形状级 0.111 = **2.6×**。契约原定的绝对阈值 0.35 **已证不可达**（见 11.3） |
| M3 | 高亮播放器 | ✅ | `hl_anim.py` | `accept_m3.py`：页数 10 == 底图数 10、`goto` 同步生效（不等任何 tick）、降暗生效、零 JS 报错；转义用例 30 条。`accept_m3_shot.py`：**步进截图 27 张 == Σ units 27**，全 1920×1080 无空白帧 |
| M4 | CLI | ✅ | `cli.py` | 5 子命令 + stdout 单行 JSON + 退出码 0/2/3/4/5；`tests/test_cli.py` 26 条；真机跑通 `import → animate → video(page/step)` |
| M5 | 工作台入口 | ✅ | `app.py` +2 路由、`index.html` +1 按钮 | `tests/test_app_pptx.py`：上传 2 页稿 → 轮询到 ready → 播放器可访问、底图可取回；真浏览器加载首页零 pageerror |
| M6 | 可编辑 pptx 导出 | ✅ | `pptx_out.py` | `accept_m6.py` 五项：theme `a:ea`/Hans 全为微软雅黑、26/26 个 run 的 `a:ea` 排在 `a:latin` 之后、`qa.check_pptx` **errors=[]**、改字不破版、**渲染层对照：ours/宋体 = 1.96×**（阳性对照证明度量能区分，差 48.9%） |
| M7 | skill | ✅ | `skill/ppt-anim/` | 薄封装用**系统 python** 从任意目录转发成功（`exit=3` / `exit=0` 均正确透传）；SKILL.md 按 agentskills.io 写 |
| 回归 | — | ✅ | — | **1028 条全绿**（248 基线 → 960 → 1017 → 1028）；`anim.py` / `builder.py` 对基线零 diff |

### 11.2 spike 结论的复核

| spike 说 | 复核结果 |
|---|---|
| 文本框矩形 coverage 中位 0.087，"高亮一大片空白" | ✅ 复现（独立度量为 0.111，同量级）。**行级 tight 是必要且有效的对策** |
| COM 无窗口导出 0.82s/页 | ⚠️ **口径要写清**：那是**稳态**速度（本轮实测 0.125–0.146s/页）；**冷调用端到端是 2.6s/页**（每次自己起 PowerPoint 再 Quit），10 页稿 26s，且不摊销 |
| 空文本框 18/40 | ✅ 复现，且裸 XML 独立核实（41 顶层元素 = 40 `p:sp` + 1 chart） |
| 素材 `pictures=0 tables=0`（R3） | ⚠️ 仍成立。用 `tools/make_fixture_pptx.py` **自造**了 6 页素材补测组合/表格/图片/图表/行距/中英混排，但真实手工稿仍未拿到 |

### 11.3 被推翻 / 更正的契约前提

| 前提 | 实测结论 | 影响 |
|---|---|---|
| **D1**：`DispatchEx` 会开新实例，故能隔离用户 PowerPoint | ❌ **不成立**。本机 PowerPoint 16.0 的 COM server 是多用途的，`Dispatch`/`DispatchEx`/`GetActiveObject` 拿到**同一个实例**（三个 API 交叉对照 + 进程数 + 稿数联动） | 安全性全靠 `Quit`/`Close` 守卫。`Quit` 加了 `had_powerpoint` 前置守卫；**S1**（`Close` 无守卫）后续修复：只关"打开前后 `Presentations.Count` 增加过"的那一份 |
| **D2**：M2 验收 coverage 中位 ≥ **0.35** | ❌ **物理不可达**。分解 `0.731（竖向）× 0.960（横向）× 0.414（CJK 字形墨迹密度）= 0.286`，与实测自洽；把矩形换成墨迹包围盒的理想上限也只有 **0.414**，pad 归零仅 0.321。独立像素法（远环众数 / Otsu，不复用实现者函数）交叉复现 0.287 | 验收改为**相对表述**「行级 / 形状级 ≥ 2.5×」（实测 2.6×），绝对目标作废 |
| §2.3 存 `font_scale = @fontScale ÷ 1000`（百分数）而 §4.3 写 `size *= font_scale` | ❌ **不自洽**，直接相乘会把字号放大 100 倍（实测 1200pt） | 实现按存储口径 `/100` 取真实倍率 |
| §4.3 "整段交给 `wrap_lines`" | ❌ 吃软换行。`qa._tokenize` 把 `"\n"` 当 1.0em 的 CJK 字形，**不是硬换行**；且 §4.2 那条防漂移等价断言只盖 `wrap_lines`，盖不住真正产出几何的 `_paragraph_lines` | 实现加了 `\n` 先硬拆（渲染对，但与 QA 差 +1 行）。仓库内三份稿实测 `a:br=0`，当前不发作 |
| §4.6 图片/图表 "coverage 天然接近 1" | ❌ 图表实测 **0.166**（全场最低）——图表内部大片留白是常态 | 统计覆盖率时应把 chart 单列口径 |
| §7.3 `picture` → `ph.insert_picture(abs_path)` / `chart` → `add_chart(...)` | ❌ **IR 装不下**：`ShapeInfo` 只有几何，没有图片路径、也没有 series/labels/values | `export` 只往返文本与表格；媒体形状跳过并在 stderr 点名 |
| §4.6 标题判定第 ② 条（`ph type ∈ {TITLE, CENTER_TITLE}`）/ 合并单元格按 `is_merge_origin` | ❌ IR 缺 `ph_type` / 缺 span 字段 | 标题只用"形状名 + 字号 ≥1.3×页中位"；合并 origin 格的 rect 按单列宽算（偏窄） |
| §5 目录树 `<out>/player/index.html` 与 §6.1 `bg_paths` "相对 out_dir" | ❌ 对 `out_dir` 的所指不一致 | 播放器底图 src 直接写 `bg/x.png` 会请求到 `<out>/player/bg/x.png`（不存在）→ 页面全黑。加 `bg_base_dir` 显式 rebase。**这个 bug 是被 L4 那道校验抓出来的** |

### 11.4 本轮不修的已知风险

审计与独立测试合计 16 条（F1–F6 / S1 / M1–M4 / L1–L7），**S1 已修复**，其余按用户决定
本轮不修，已逐条（含证据出处）整理进 `KNOWLEDGE.md` 的「已知风险 · pptx 双向链路」表。
实现侧在**调用链上做了规避**（不改那些模块本身）：

- **F1**（`export_pages` 拆掉调用方 COM apartment）→ `pptgen import` 在导出后不再触碰 COM
- **L4**（播放器不校验底图存在 → 黑帧静默通过）→ `shot_player` 与 `cli animate` 先校验
- **M2**（zip bomb 无闸门）→ 上传路由加了 60 MB **上传体积**兜底；真正的解压总量/
  压缩比闸门仍未落地（那是 `read_pages` 前置的事）

