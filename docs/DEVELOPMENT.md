# 开发文档 · ppt-generator

> 面向**人类开发者**与接手维护的 agent。速查手册（符号锚点式）见 [`KNOWLEDGE.md`](../KNOWLEDGE.md)；
> pptx 链路的接口契约见 [`docs/PPTX_INTERFACE.md`](PPTX_INTERFACE.md)；使用说明见 [`docs/USER_GUIDE.md`](USER_GUIDE.md)。

---

## 一、项目概览

一个本地单机 Flask 应用，把「一个主题」或「一份现成 pptx」变成**可演示的幻灯片**。两条产线并存、互不干扰：

| 产线 | 输入 | 核心技术 | 产物 |
|---|---|---|---|
| **A · 从主题生成** | 一句话主题 / 文档 | LLM 自主设计单文件 HTML（非模板填充） | HTML 设计稿 → pptx / PDF / 动画播放器 / MP4 |
| **B · 从 pptx 导入** | 现成 `.pptx` | PowerPoint COM 导保真底图 + python-pptx 取几何 → **行级高亮**锚点 | 高亮讲解播放器 → MP4；反向可导回**可编辑** pptx |

**代码规模**（2026-09-13 核实）：

```
python 模块      7038 行 / 12 个
前端             1594 行（templates/index.html，单文件工作台）
测试             37 个文件 / 1066 条用例
HTTP 路由        45 条
```

**运行形态**：本地单用户。全局 `state` + `lock` + worker 线程，**单任务**（不面向多用户/公网）。

---

## 二、代码地图

### 2.1 分层

```
[L1 接入层]
  app.py (1631)          Flask + 45 路由 + 状态机 + 三套导出接线
  templates/index.html   工作台前端（单文件：步骤条/两栏/预览/日志）
  cli.py (433)           pptgen CLI —— agent 入口（方向 B）
  main.py                旧 CLI（仅 大纲→生图→builder pptx，保留未废弃）

[L2 编排层]  全在 app.py：_start_generation / _generation_worker / _gen_images
                            / _design_and_save / _pause_gate / _video_jobs

[L3 能力层]  与 Web 解耦，纯函数为主 —— 可被 CLI 与 app 同时复用
  内容生成   outline.py (233)   主题/文档 → 章节化大纲 JSON（自评迭代 + 布局多样化）
             style.py           6 套风格库 + AI 选题风格
             template.py (223)  内置模板 + 参考稿识别 + 自定义模板持久化
             critic.py          视觉校验 + 定点/整篇修改
  产物生成   html_gen.py        LLM 直出单文件 HTML 设计稿（注入 print CSS 兜底）
             builder.py (477)   pptx 图片版导出（旧路线，保留）
             pptx_out.py (430)  pptx 可编辑版导出（母版 + 占位符）
  渲染导出   shot.py            设计稿截图（Playwright + 三重确定化）
             anim.py            逐元素揭示播放器（服务产线 A）
             hl_anim.py (477)   高亮讲解播放器（服务产线 B）
             video.py (120)     ffmpeg 配方层（zoompan + xfade → MP4）
  pptx 专用  pptx_io.py (813)   读形状（python-pptx）+ COM 导保真底图
             hl_layout.py (547) **行级**高亮矩形（复用 qa.py 字体度量）
  质检       qa.py (301)        pptx 几何门禁（FontTools 量文本溢出/越界）
             quality.py         deck 级重复页 / 过瘦页检测
  基建       llm_util.py (212)  多渠道 / 运行时模型切换 / 配图总开关
             uploads.py         PDF 导入（pdfplumber）

[L4 产物层]  output/{decks,projects,images,animation,videos,templates,pptx_src}
```

### 2.2 模块职责速查

| 模块 | 一句话职责 | 关键符号 |
|---|---|---|
| `app.py` | 状态机 + 路由 + 导出接线 | `_generation_worker` `_design_and_save` `_video_jobs` |
| `outline.py` | 主题 → 大纲 JSON | `generate_outline` `_normalize` `_normalize_chart` `_diversify_layouts` |
| `html_gen.py` | 大纲 → 单文件 HTML 设计稿 | `generate_html_deck` `_ensure_print_css` `_extract_html` |
| `builder.py` | 大纲 → 图片版 pptx | `build_ppt` `_set_ea` `_fit_title_size` |
| `pptx_out.py` | deck.json → **可编辑** pptx（母版+占位符） | `build_deck_pptx` `read_template` `fit_text` |
| `pptx_io.py` | pptx 读取 + COM 导底图 | `read_pages` `export_pages` `powerpoint_available` `PptxError` |
| `hl_layout.py` | **行级**高亮矩形 + 覆盖率度量 | `wrap_lines` `build_units` `measure_coverage` `page_shapes` |
| `hl_anim.py` | 高亮播放器 + 步进截图 | `build_player` `shot_player` |
| `anim.py` | 逐元素揭示播放器（产线 A） | `build_player` |
| `shot.py` | 设计稿截图 | `shot_deck` `_launch_browser` `_screenshot_settled` |
| `video.py` | ffmpeg 配方（构造/执行分离） | `build_page_args` `build_final_args` `synthesize` |
| `qa.py` | pptx 几何门禁 + **字体度量底座** | `check_pptx` `_measure_lines_ex` `_char_width_pt` `_is_hard_break` |
| `llm_util.py` | 多渠道/模型运行时 | `llm_client` `get_model` `add_channel` `images_enabled` |
| `cli.py` | agent 入口 | `main`（5 子命令） |

---

## 三、两条产线的数据流

### 产线 A · 从主题生成（`app.py` 状态机）

```
POST /api/generate
   └─ _start_generation()  ← 锁内初始化并置 phase（防前端早于 worker 停止轮询）
        └─ worker: _generation_worker()
             ├─ outline.generate_outline()        phase=outline
             ├─ style.decide_style() / 模板优先    （tpl_style 优先于 AI 自选）
             ├─ _gen_images()                      phase=images
             │     └─ 逐页 image_gen + critic.review_image 闭环（不契合→改词重生，封顶 2 次）
             ├─ _pause_gate("images")              ← 分步确认时在此阻塞（await_step）
             ├─ _design_and_save()                 phase=designing
             │     └─ html_gen.generate_html_deck()
             └─ _save_project_snapshot()           phase=ready
```

前端每秒轮询 `GET /api/status` 取 `phase / await_step / slides / html_path / log`。

### 产线 B · 从 pptx 导入（独立于状态机，单开线程）

```
POST /api/pptx/import   （或 cli.py import）
   └─ pptx_io.read_pages()        每页形状：EMU 坐标 / 每 run 字号 / 段落 / 类型（递归展开组合）
   └─ pptx_io.export_pages()      PowerPoint COM 无窗口导出保真底图（**视觉 100%**）
   └─ hl_layout.build_units()     **行级**高亮矩形（框内边距 + 字号 + FontTools 贪心换行）
   └─ hl_anim.build_player()      <div class=slide><img src=bg><div class=hl>…</div></div>
   └─（可选）hl_anim.shot_player() → video.synthesize() → MP4
```

**为什么必须"底图 + 文字层"而非直接重建 HTML**：`shape.left/top/w/h` 是**文本框外框**，与文字实际落墨位置无关 —— 实测框内墨迹填充率中位数仅 **0.087**。视觉保真只能靠 PowerPoint 自己渲染。

---

## 四、核心约定

### 4.1 状态机与并发

```python
state = {...}          # 全局字典
lock  = threading.Lock()
phase: idle → outline → images → designing → ready      # 分步确认时插入 review
await_step: "images" | "outline" | None
```

- **只允许一个 worker**：`_start_generation` 在锁内判断并置 `phase`，避免任务重叠
- **`review` 态不放行 `/api/refine`**（会另起 worker 与阻塞的闸门竞争）；暂停期改文字用 `/api/slide/<i>/text`
- 视频合成另有 `_video_jobs: set[str]` 做**同稿并发守卫**

### 4.2 单位制（pptx 链路）

| 单位 | 定义 | 用途 |
|---|---|---|
| `*_emu` | OOXML 原生 EMU，**唯一权威坐标** | 所有几何计算 |
| `*_pt` | `emu / 12700` | 字号、字体度量 |
| `*_px` | `emu × 导出宽 / 画布宽EMU`，**固定在导出底图像素空间** | 播放器 CSS 定位 |

换算常量在 `qa.EMU_PER_PT = 12700`；1920 宽 @13.333in 画布时 `1pt = 2px`。
**播放器不重算坐标**，靠舞台 `transform: scale()` 整体缩放。

### 4.3 产物目录

```
output/
├── decks/          HTML 设计稿（产线 A）
├── projects/       项目快照（ready 时自动存，上限 30）
├── images/         配图 + cache/（key = md5(prompt+title+points)）
├── animation/      截图序列 + 逐元素播放器
├── videos/         MP4
├── templates/      JSON 风格模板（template.py）；pptx 母版在 templates/pptx/
└── pptx_src/       pptx 导入产物：deck.json + bg/ + units.json + player/
```

> `output/` 整体在 `.gitignore` 内 —— 新克隆的仓库里没有验收素材，`tests/test_pptx_io.py` 等的真机用例会**静默 skip**。要跑全量真机验收，先跑 `tools/probes/accept_m1.py` 生成素材。

### 4.4 错误语义

CLI 与 `PptxError` 共用错误码表（见 `docs/PPTX_INTERFACE.md` §9）。约定：

- **失败一律给明确中文 + 修复指引**，不抛裸堆栈
- 错误码要能区分「用户稿有问题」与「我们的 bug」：例如缺 `p:sldSz` 应报 `PPTX_UNREADABLE(3)` 而非 `INTERNAL(5)`
- CLI `stdout` **只有一行 JSON**（`{"ok":true,"cmd":...,"data":{...}}`），人读信息走 `stderr` —— 便于 agent 解析

---

## 五、如何扩展

### 5.1 加一套视觉风格

1. `style.py::STYLE_LIBRARY` 加一条（色板 + 气质描述 + 设计指引）
2. `style.py::decide_style` 的选择提示词里会带上风格库，无需改动
3. 如果要同时作为**内置模板**（可选，供用户手选）：`template.py::builtin_templates()`

### 5.2 加一个导出格式

**照现有形状做，不要改既有导出**：

- 纯 Python 生成 → 新模块（如 `pptx_out.py` 之于 `builder.py`），并在 `cli.py` 加子命令
- 需要浏览器/外部进程 → 复用 `shot.py`（截图）或 `video.py`（ffmpeg 配方），**命令构造与执行分离**（前者是纯函数、可单测；后者才真跑）
- 接线：`app.py` 加路由 + `templates/index.html` 加按钮 + 加进 `/api/artifacts` 产物列表

### 5.3 加一条 API 路由

```python
@app.route("/api/xxx", methods=["POST"])
def api_xxx():
    ...
```

约定：读 `state` 加锁；长任务开 worker 线程（别阻塞请求）；返回 `jsonify({...})`；静态资源走白名单路由（见 `app.py` 的 `/files` `/videos` `/animation` `/pptx`）。

---

## 六、测试策略

```
tests/                37 个文件 / 1066 条
tests/conftest.py     autouse 夹具把 runtime_config.json 隔离到临时目录
                      → 用例不受本地渠道/配图开关影响（**别删这个夹具**）
```

**四类测试**

| 类型 | 例子 | 要点 |
|---|---|---|
| 纯逻辑单测 | `test_outline.py` `test_qa.py` | mock 掉 LLM，快 |
| 接线测试 | `test_wiring.py` `test_cli.py` | 薄用例覆盖路由/CLI 参数与错误码 |
| **真机 COM** | `test_pptx_io.py` 的部分用例 | 会启动 PowerPoint，**必须能 skip** |
| **真浏览器** | `test_shot_settle.py` | Playwright，无 Chrome 则 skip |

**写 COM/浏览器用例的范式**（照 `tests/test_s1_close_guard.py`）：

```python
# 优先打桩：把 win32com / pythoncom 换成假模块，不启动 PowerPoint
monkeypatch.setitem(sys.modules, "pythoncom", fake_pythoncom)
monkeypatch.setitem(sys.modules, "win32com.client", fake_client)
```

真机用例务必加 skip 条件；**同一时刻只允许一个 pytest 在跑**（并发会抢 PowerPoint COM，产生假失败）。

---

## 七、调试与验证手法

### 7.1 探针范式（真机行为验证）

放 `tools/probes/`，参考 `tools/probes/s1_probe_v2.py`：

1. **独立观察者**：要验证「某个调用是否破坏了调用方的状态」，观察者必须在**另一个进程/线程且状态有效**处 —— 同进程自证会出现假阳性（S1 的 v1 探针就是这么翻车的）
2. **阳性/阴性对照**：同一探针在「修复前 / 修复后」各跑一次（`git stash` 前后），两次输出都留档
3. **安全三重闸门**（碰用户软件时强制）：
   - 启动前若已有该软件进程（如 `POWERPNT.EXE`）→ **整条中止**
   - 只用 `tempfile` 副本，绝不触碰真实文件
   - 只清理**自己启动**的实例；跑完 `tasklist` 复核无孤儿进程

### 7.2 深坑速查

| 现象 | 真因 | 处置 |
|---|---|---|
| 无头截图里图表空白 | `IntersectionObserver` 在滚动后才填充 | `shot.py` 的三重确定化（IO shim + `animations="disabled"` + 稳定性轮询）—— **别退回固定 sleep** |
| 导出 PPT/PDF 报 500 | 上游回 **WebP**，python-pptx 图片白名单不含 WEBP | 落盘前按**魔数**归一为 PNG（`image_gen._write_png` / `copy_as_png`，缓存复用路径也要转） |
| 调用方后续用 COM 报 `CO_E_NOTINITIALIZED` | `export_pages` 内的 `CoUninitialize` 拆了调用方 apartment | 已修：用 `CoInitializeEx` 的 HRESULT 判定，只在自初始化时收尾 |
| 明明没在用 PowerPoint，进程却是 1 个 | `Dispatch`/`DispatchEx` 在本机是**同一个实例** | 安全只能靠守卫，见 §8 |
| pytest 跑很久没反应 | 子进程没起来（如 venv trampoline 坏了） | 看 **CPU 时间**而非墙钟；`sys.prefix` 是否等于 venv 目录是个好指标 |
| 高亮块圈住一大片空白 | 用了**文本框矩形**而非行级矩形 | 用 `hl_layout.build_units`（行级 tight），别直接用 `shape.left/top/w/h` |

---

## 八、红线与已知风险

### 8.1 绝不许动的（改动前必须读 `KNOWLEDGE.md` 对应条目）

| 红线 | 原因 |
|---|---|
| 设计稿 `iframe` 的 `sandbox="allow-scripts"`（**禁** `allow-same-origin`） | 有意设计：LLM 产物可翻页但**不能**同源 fetch 调 API / 读 cookie |
| `pptx_io.export_pages` 的 **`we_opened` 守卫**（只 `Close` 自己新增的稿） | PowerPoint 对**已打开的同一路径**返回既有实例 → 无条件 Close 会关掉**用户未保存的稿**（S1，真机复现过） |
| `export_pages` 的 **`had_powerpoint` + 进程数双守卫**（不 Quit 用户实例） | 我们实际附着在用户的 PowerPoint 上（D1 实测） |
| **绝不写 `app.Visible`** | 会连带把用户正在看的窗口藏起来 |
| `templates/index.html` 的路径白名单（`/files` `/videos` `/animation` `/pptx`） | 产物可见性与目录穿越防护 |
| `tests/conftest.py` 的 autouse 配置隔离夹具 | 用例不受本机渠道/开关影响 |

### 8.2 已知未修风险（证据见 `docs/AUDIT_REPORT.md`）

| 编号 | 问题 | 现状 |
|---|---|---|
| F3 | 冷调用 ~2.6s/页（每次新建+Quit PowerPoint） | **放弃修**（复用实例在 Flask 每请求新线程下触发 `RPC_E_WRONG_THREAD`，需队列化重构；实测数据见 `docs/FIX_REPORT.md`） |
| L2 | 表格列宽/行高数组短于单元格时静默截断 | 未修（低危） |
| L3 | 被改坏的 `deck.json` 裸崩而非 `IR_MISMATCH` | 未修（低危） |
| L5 | `had_powerpoint` 采样与 `DispatchEx` 之间存在 TOCTOU | 未修（低危） |
| L6 | 共享 PowerPoint 实例上没有互斥 | 未修（低危） |
| L7 | `measure_coverage` 用 `0.0` 重载三种语义 | 未修（低危） |

---

## 九、本地开发环境

```bash
# 环境：Windows + .venv（uv 建的 venv）
.venv/Scripts/python.exe -m pytest tests/ -q          # 全量回归（约 3 分钟）
.venv/Scripts/python.exe -m pytest tests/test_cli.py -q  # 单文件
```

- **本机有多个 python / 多个 uv**：venv 的 trampoline 依赖 uv 的 python 目录命名，**uv 升级可能导致 venv 整体失效**（症状：`uv trampoline failed to spawn Python child`）。正确修法是用当前版 uv `venv --python 3.11 --allow-existing .venv` 重建 trampoline（实测：改 `pyvenv.cfg` 无效、建 junction 会让 `sys.prefix` 错位）
- **本机 bash 是 WSL**：路径用 `/mnt/d/...`；调 Windows 侧 CLI 走 `powershell.exe`
- **WSL 连不上 `github.com:443`**：推送/拉取必须走 Windows 侧 git
- LLM 配置在 `runtime_config.json`（多渠道，**key 不落仓库**）或 `.env`；`images:false` 一键关配图

---

## 十、相关文档

| 文档 | 内容 |
|---|---|
| [`README.md`](../README.md) | 项目简介 + 快速开始 |
| [`docs/USER_GUIDE.md`](USER_GUIDE.md) | **使用说明**（工作台操作 / CLI 参考 / 故障排查） |
| [`KNOWLEDGE.md`](../KNOWLEDGE.md) | agent 执行手册（符号锚点 / 状态机 / 路由 / 已知风险表） |
| [`docs/PPTX_INTERFACE.md`](PPTX_INTERFACE.md) | pptx 链路接口契约 v2（含被实测推翻的 16 条前提） |
| [`docs/FIX_REPORT.md`](FIX_REPORT.md) | 9 项遗留风险的修复记录（四段式：根因/修法/前后对照/用例） |
| [`docs/AUDIT_REPORT.md`](AUDIT_REPORT.md) | 只读审计报告（1 严重 + 4 中 + 7 低） |
| [`docs/TEST_REPORT.md`](TEST_REPORT.md) | 独立测试报告（用不同方法交叉验证） |
| [`PLAN_PPTX_ANIM.md`](../PLAN_PPTX_ANIM.md) / [`PLAN_RISK_FIX.md`](../PLAN_RISK_FIX.md) | 计划与执行记录 |
