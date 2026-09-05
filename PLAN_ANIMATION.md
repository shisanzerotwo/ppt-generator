# 动画与视频导出 · 开发计划文档

> 版本 v1（2026-09-06）｜ 前置基线：commit `c8ba9b3`（模型切换，202 tests passed）
> 目标：把 PPT（HTML 设计稿）输出成**动画**（阶段一）与**视频 MP4**（阶段二），画幅 **16:9**。

## 一、Skill 调研结论（按要求先找 skill）

| 范围 | 结果 |
|---|---|
| 本机 8 个已装 skill（ppt-maker / superpowers / web-design / unity-pipeline 等） | **无**动画/视频导出能力；ppt-maker 只覆盖 HTML/pptx 生成与截图验证 |
| 公共生态（agentskills.io / GitHub / SkillsLLM） | 找到 3 类参考：**n0an/ffmpeg-skill**（ffmpeg 命令配方库）、**html-video**（HTML→MP4 可插拔渲染引擎）、**Remotion 系**（编程化视频，需 Node 工程） |

**结论**：没有开箱即用的「PPT→动画→视频」skill。**采纳 ffmpeg-skill 的思路**——把 ffmpeg 命令做成经过验证的配方函数（本项目中固化成 `video.py` 配方层），而非引入 Node/Remotion 重依赖。实施完成后可选把配方沉淀为本仓库自建 skill（`~/.agents/skills/ppt-animation/`），供其它 agent 复用（后续可选项）。

## 二、环境审计（已完成）

| 依赖 | 现状 | 动作 |
|---|---|---|
| Chrome | ✅ 已装（`C:/Program Files/Google/Chrome`） | Playwright 走 `channel="chrome"` 复用，**免下浏览器内核** |
| playwright（pip 包） | ❌ venv 未装 | `uv pip install playwright`（纯 Python 包，很小） |
| ffmpeg | ❌ 未装（阶段二必需） | `winget install ffmpeg`（v1.29 可用），约 100MB |
| Pillow | ✅ 12.3.0 | 图片尺寸校验用 |
| 画幅基础 | ✅ 设计稿 print CSS 已是 `@page 1280px×720px`（16:9） | 截图 viewport 直接对齐 |

## 三、技术方案

### 阶段一：图片切分成动画（M1 + M2）

**M1 截图管线（`shot.py`）**
- 输入：设计稿 HTML 路径 → 输出：`output/animation/<稿名>/slide_1..N.png`（1280×720）
- 实现：Playwright 同步 API，`chromium.launch(channel="chrome", headless=True)`；逐页 `evaluate` 滚动到第 N 个 `.slide` section（scroll-snap 定位）→ `viewport 1280×720` 截图；等待字体/图片 settle（`networkidle` + 300ms 缓冲）
- 兜底：Chrome 未装/启动失败 → 明确报错并给安装指引；截图页数与 `.slide` 数量校验一致

**M2 动画播放器（`anim.py` + 模板）**
- 产物：`output/animation/<稿名>/index.html`——**零依赖单文件**（延续项目哲学）
- 动画形态：每页 **Ken Burns 缓推**（随机方向 scale 1.0→1.08，时长可调，默认 4s/页）+ **0.8s 交叉淡入**转场；封面页淡入开场、尾页定格
- 控制条：播放/暂停、上一页/下一页、进度点、循环开关；16:9 自适应舞台（letterbox 居中）
- 参数注入：页图列表/每页时长/转场时长由后端以 JSON 注入 `<script>`
- 路由：`GET /animation/<path>` 提供访问；`POST /api/export_animation {seconds_per_slide?, fps?}` → ready 时截图+生成播放器，返回 `{path, pages}`；生成中 409

### 阶段二：视频导出（M3，依赖 ffmpeg）

**M3 视频合成（`video.py` 配方层，借鉴 ffmpeg-skill 配式）**
- 流程：复用 M1 截图 → 每页 `zoompan`（Ken Burns，25fps，d=4s×25 帧）→ 逐页片段 → `xfade` 链式转场（fade 0.8s）→ `libx264 + yuv420p + scale=1280:720,setsar=1` 兜底 → `output/videos/<稿名>_<时间戳>.mp4`
- 配方固化为纯函数 `build_ffmpeg_args(pages, seconds, out)`：**命令构造与执行分离**（可单测命令正确性，不必真跑视频）
- API：`POST /api/export_video` → 先确保截图存在（无则先跑 M1）→ 合成 → `GET /videos/<file>` 下载路由（白名单 .mp4）
- ffmpeg 缺失：`shutil.which("ffmpeg")` 预检，缺失时返回明确安装指引（`winget install ffmpeg`）；**长任务防超时**：合成走后台线程 + 轮询（沿用现有轮询机制），完成/失败写日志

### 前端（两个入口）
- 导出区新增「🎞 导出动画」与「🎬 导出 MP4」（阶段二按钮在 ffmpeg 就绪前隐藏/置灰，由 `/api/models` 式探测接口或导出时报错提示驱动——首版：报错提示装 ffmpeg）
- 动画播放器在新标签/内嵌查看器打开（复用现有 iframe 查看器）

## 四、里程碑与验证标准

| 里程碑 | 交付物 | 验证 |
|---|---|---|
| M1 截图管线 | shot.py + `POST /api/export_animation` 的截图部分 | 对现有 13 页设计稿截图 13 张，全部 1280×720（pytest 抽验 + 文件尺寸断言） |
| M2 动画播放器 | anim.py + 播放器模板 | 浏览器打开自动轮播、Ken Burns 生效、控制条可用（真机验收）；HTML 含页数 JSON 注入（单测） |
| M3 视频导出 | video.py + `POST /api/export_video` | ffprobe 验证：分辨率 1280×720、时长 ≈ 页数×4s−转场重叠、编码 h264；命令构造纯函数单测 |
| 全量回归 | — | 现有 202+ 测试保持全绿；新增测试约 8~10 条（ffmpeg 实跑类用例标记 skipif 无 ffmpeg） |

## 五、风险与对策

| 风险 | 对策 |
|---|---|
| LLM 生成 CSS 千奇百怪，headless 截图时动画/懒加载未 settle | `networkidle` + 固定缓冲 + 字体 ready 等待；必要时注入 CSS 禁用过渡 |
| scroll-snap 页定位漂移 | 用 `element.scrollIntoView()` 而非像素偏移；截图前读 `boundingRect` 校验 |
| ffmpeg 下载/安装失败 | winget 失败则手动下 gyan.dev 静态包并加 PATH（文档写清）；M3 前 `which` 预检 |
| 晚高峰 Agnes 无关——本管线**零 LLM 调用** | —（纯本地渲染，成本为零） |
| 长稿视频合成慢（CPU 编码）| preset=veryfast；后台线程 + 日志进度；页数上限提示 |

## 六、明确不做（首版）

- 音频/配乐/TTS 配音、逐元素动画（需解析设计稿 DOM，复杂度高）、云端渲染、竖屏/其它画幅
- Remotion/Node 渲染链（重依赖，与零依赖哲学冲突）

## 七、实施顺序

**阶段一**：装 playwright → M1 截图管线+测试 → M2 播放器+测试 → 前端按钮 → commit
**阶段二**：`winget install ffmpeg` → M3 配方层+测试 → 后台合成+下载路由 → 前端按钮 → commit
每里程碑向你汇报验证证据后再进下一个（Loop 三段式）。
