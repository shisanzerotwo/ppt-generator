# PPT 生成器（ppt-generator）

AI 驱动的 PPT 生成器：输入主题，自动完成「大纲 → 生图 → 视觉校验 → **AI 自主设计 HTML 幻灯片** → 智能选风格」全流程。核心亮点是**让 LLM 根据内容自主设计每页布局**，而非固定模板——数据页画图表、概念页出卡片、封面做视觉冲击、时间线页做时间轴，真正“按内容适配”。

## ✨ 功能特性

- **LLM 自主设计 HTML**：直接把大纲 + 图片交给 LLM 生成完整单文件 HTML 幻灯片，每页布局因内容而异，纯内联 CSS + SVG，零外部依赖，可离线打开、可打印 PDF
- **AI 智能选风格 + 模板导入**：默认 AI 根据主题从 6 套风格库自选（深空科技 / 极简商务 / 清新浅色 / 暖调人文 / 活力渐变 / 自然墨绿）；也可**手动选内置模板**（色板 + 版式骨架）或**导入参考稿**（上传设计图/HTML，AI 识别其配色与风格气质仿制），用户指定后跨生成保留、可清除
- **左右分栏工作区**：出结果后左编辑（对话/卡片/排序）、右放映（设计稿常驻预览），改一处右侧实时看，不必上下滚动；窄屏自动回退单栏
- **完整的生成流水线**：大纲（章节化、8~10 页）→ 逐页生图 → 视觉校验（提议者-审核者闭环，不契合自动改词重生）→ LLM 设计 HTML
- **对话式修改（含定点）**：ready 后可整篇改（“整篇换商务风”），也可在右侧「AI 协作」面板圈选某页/某条要点定点改——只动被圈定的部分，其余页连配图一并保留，不打回重做
- **内嵌查看器 / 放映 / 灯箱**：历史设计稿在工作台内的居中弹窗查看（ESC 或点遮罩关闭），一键全屏放映（方向键翻页）；卡片配图点击放大看清生图质量
- **分步确认（human-in-the-loop）**：可选开启，生成在“大纲与风格”“配图”两个节点暂停，审阅横幅上「继续」放行或「先改改」先调整
- **多格式导出 + 产物区**：主产物 HTML（AI 设计稿，可预览/打印 PDF），另有 pptx（可编辑版）、PDF、大纲 txt；导出后侧栏「导出产物」可直接点开/下载
- **打印分页兜底**：LLM 漏写 `@media print` 时，`html_gen` 自动注入分页规则（1280×720 每页一张、`page-break-after`），保证浏览器打印/导出 PDF 排版正确
- **历史设计稿 / 历史项目库**：侧栏可折叠分组列出最近记录；历史项目含大纲快照（主题/风格/品牌/页面/设计稿），可载入继续编辑（设计稿缺失时自动置空可重设计）
- **品牌模板**：设定品牌名与主色（`#RRGGBB`），用品牌色覆盖 AI 风格的强调色并注入设计指引；用户级偏好跨生成保留，传空即清除
- **页面排序与增删**：每张卡片可上移/下移/删除，也可插入新页；调整后点「重新设计」让设计稿同步
- **专业工作台 UI**：左侧栏 + 主工作区，中性石墨配色 + 青瓷绿强调、内联 SVG 图标、暗色模式（`prefers-color-scheme`），遵循 web-design skill 反“AI 味”规范
- **📊 pptx 双向链路（新）**：导入一份**现成的 `.pptx`** → PowerPoint 导出保真底图 + 行级高亮锚点 → 可步进讲解的**高亮播放器**与 MP4；反向还能把产物导回**可编辑 pptx**（母版 + 占位符，中文渲染为微软雅黑而非宋体）。同一套能力另有 `pptgen` CLI 与 `skill/ppt-anim/`，供 agent 直接调用
- **开发友好**：Web 界面实时进度条与阶段、可折叠日志、风格徽章；1028 条 pytest 单测

## 🚀 快速开始

```bash
# 1. 安装依赖（建议虚拟环境）
python -m venv .venv
# Windows:
.venv\Scripts\activate
pip install -r requirements.txt

# 2. 配置 API Key
cp .env.example .env   # 然后编辑 .env 填入 ZHIPUAI_API_KEY 与 ZHIPUAI_BASE_URL

# 3. 启动
python app.py
# 浏览器打开 http://127.0.0.1:5000，输入主题 → 生成
```

`.env` 示例：

```ini
ZHIPUAI_API_KEY=sk-你的-key
ZHIPUAI_BASE_URL=https://apihub.agnes-ai.com/v1
# 可选：设计任务单独指定更强的模型（默认取 ZHIPUAI_CHAT_MODEL 或 agnes-2.0-flash）
# ZHIPUAI_DESIGN_MODEL=agnes-2.0-flash
```

## 🧠 工作流程

```
输入主题 / 文档
      │
      ▼
┌─────────────┐   outline.py     章节化大纲（8~10 页，含 type/title/points/chart/layout）
│   大纲生成   │──────────────►  自评迭代（结构完整性校验）
└─────────────┘
      │
      ▼
┌─────────────┐   style.py       AI 按主题/内容从风格库自主判断视觉基调
│  智能选风格  │──────────────►  深空科技 / 极简商务 / 清新浅色 / 暖调人文 / 活力渐变 / 自然墨绿
└─────────────┘
      │
      ▼
┌─────────────┐   image_gen.py   逐页生图（按 image_prompt）
│   逐页生图   │──────────────►  critic.py 视觉校验（不契合 → 改词重生，封顶 2 次）
└─────────────┘
      │
      ▼
┌─────────────┐   html_gen.py    LLM 直接生成完整 HTML（AI 自主设计每页布局）
│  AI 设计 HTML│──────────────►  注入风格基调 + 图片相对路径；截断自动重试
└─────────────┘
      │
      ▼
   ready（Web 顶部 iframe 预览 / 新标签打开 / 重新设计 / 侧栏历史设计稿）
```

## 🔄 pptx 双向链路（导入现成稿 → 高亮讲解 → 导回可编辑）

上面那条流水线是"**从主题生成**"；这条是"**已有 pptx**"的另一条路，两条并存、互不干扰。

```
现成的 deck.pptx
   │
   ├─[1] pptx_io.read_pages()      python-pptx 逐页取形状：坐标(EMU)/文本/字号/类型/组合
   │                               过滤空文本框、零尺寸、隐藏形状
   ├─[2] pptx_io.export_pages()    PowerPoint COM 无窗口导出每页 PNG（**视觉 100% 保真**）
   │                               ⚠️ 需要装 Office；且必须用绝对路径
   ├─[3] hl_layout.build_units()   行级 tight 矩形：框内边距 + 字号 + FontTools 贪心换行
   │                               （直接拿文本框矩形做高亮会"高亮一大片空白"）
   ├─[4] hl_anim.build_player()    底图 + 高亮锚点 + 其余降暗；步进/回退/自动讲解
   └─[5] video.synthesize()        截图序列 → ffmpeg → MP4（整页翻页 / 按单元逐步）
                                   反向：pptx_out.build_deck_pptx() → 可编辑 pptx
```

**工作台入口**：侧栏「导入 PPTX 高亮讲解」→ 选一份 `.pptx` → 十几秒后侧栏出现
「▶ 打开高亮播放器」。

**CLI**（`pptgen`，agent 直接调；stdout 恒为一行 JSON，进度走 stderr）：

```bash
pptgen import deck.pptx --out out/ --width 1920   # → out/{deck.json,bg/,units.json}
pptgen animate out/ --dim 0.25                    # → out/player/index.html
pptgen video   out/ --sec 4                       # → out/out.mp4（整页翻页）
pptgen video   out/ --mode step --sec 2           # → 按讲解单元逐步展开
pptgen export  out/deck.json -o out/deck.pptx     # → 可编辑 pptx（可加 --template 品牌母版）
pptgen deck    "人工智能如何改变教育" --out out/    # 可选：复用 AI 流水线出设计稿
```

- 退出码：`0` 成功｜`2` 参数错｜`3` 输入问题｜`4` 缺外部依赖｜`5` 内部错误
- 前置：**Windows + PowerPoint**（`faithful` 模式的保真底图只能由 PowerPoint 渲染）、
  `ffmpeg`（导视频）、Chrome/Edge（`--mode step` 截图）
- agent 侧封装：`skill/ppt-anim/`（`scripts/pptgen.py` 是薄转发，逻辑不复制）



| 模块 | 职责 |
|---|---|
| `outline.py` | 主题/文档 → 章节化大纲 JSON（含自评迭代、布局多样化） |
| `style.py` | AI 智能选风格：从 6 种预设风格库判断，决定色板/气质/设计指引 |
| `template.py` | 内置模板库（色板 + 版式骨架）+ 参考稿识别（vision/文本提炼风格） |
| `image_gen.py` | 按提示词生成每页配图 |
| `critic.py` | 视觉校验 + 对话式修改大纲（提议者-审核者闭环） |
| `html_gen.py` | **核心**：LLM 自主设计每页 HTML 布局，注入风格，容错重试 |
| `builder.py` | pptx 降级导出（可编辑版） |
| `pptx_io.py` | **pptx 读入**（python-pptx）+ COM 无窗口导出保真底图 + `PptxError` 错误码 |
| `hl_layout.py` | pptx 稿的**行级**高亮矩形（复用 `qa.py` 的字体度量，不另立回退值） |
| `hl_anim.py` | **高亮**播放器（与 `anim.py` 的逐元素揭示并存，勿混用）+ 步进序列截图 |
| `pptx_out.py` | deck.json → **可编辑** pptx（母版 + 占位符；与 `builder.py` 双导出并存） |
| `cli.py` | `pptgen` CLI：`import / animate / video / export / deck` |
| `skill/ppt-anim/` | 给 agent 用的 skill（SKILL.md + 薄封装），逻辑全在 `cli.py` |
| `app.py` | Flask 后端 + Web 工作台（异步生成、状态机、导出接口、pptx 高亮入口） |
| `templates/index.html` | 前端：步骤条、AI 设计稿 iframe 预览、重新设计、对话修改 |

## 🔌 API 一览

| 接口 | 说明 |
|---|---|
| `POST /api/generate` | 按主题生成 |
| `POST /api/import` / `api/import_file` | 从文档文本/文件生成 |
| `POST /api/refine` | 对话式修改：不带 `target` 改整篇；带 `target: {slide, quote?}` 只改指定页/要点（定点修改，不重跑整篇） |
| `POST /api/redesign` | 重新触发 AI 设计（重生成 HTML） |
| `GET /api/status` | 轮询状态（phase / await_step / stepwise / slides / html_path / style_name / tpl_style_name / brand / log） |
| `POST /api/slide/<i>/text` | 保存单页标题与要点 |
| `POST /api/slide/<i>/image` | 单页重新生图（可带自定义画面描述） |
| `POST /api/slide/reorder` | 页面重排序（body 传 `order`：0..n-1 的完整排列；前端提供上移/下移按钮） |
| `POST /api/slide/add` | 在 `after` 下标后插入一页空白 content 页 |
| `DELETE /api/slide/<i>` | 删除某页（至少保留一页，删至最后一页时返回 400） |
| `POST /api/brand` | 设置/清除品牌模板（名称 + 主色 `#RRGGBB`；用户级偏好跨生成保留，传空清除） |
| `GET /api/templates` | 列出内置模板（6 套色板 + 版式骨架） |
| `POST /api/template/select` | 选内置模板（`{key}`，传空清除回 AI 自动选） |
| `POST /api/template/analyze` | 上传参考稿（图片 multipart 或 `{html}`）→ AI 识别风格存为模板 |
| `POST /api/stepwise` | 开关分步确认（`{enabled}`） |
| `POST /api/continue` | 分步确认放行（仅 `review` 态有效，否则 409） |
| `GET /api/projects` | 历史项目库列表（ready 时自动快照，上限 30） |
| `POST /api/projects/load` | 载入历史项目快照继续编辑/重设计 |
| `GET /api/decks` | 历史设计稿列表（按修改时间倒序，上限 30，含标题/链接/时间） |
| `GET /api/artifacts` | 导出产物列表（pptx/pdf/大纲 txt + 设计稿 html，倒序上限 30） |
| `POST /api/export` `/export_pdf` `/export_html` `/export_txt` | 多格式导出 |
| `POST /api/pptx/import` | **上传 .pptx → 后台导入 → 高亮播放器**（进度见 `/api/status` 的 `pptx` 字段） |
| `GET /pptx/<file>` | 访问 pptx 导入产物（播放器 + `bg/*.png`） |
| `GET /decks/<file>` | 访问 AI 设计稿 HTML |
| `GET /files/<file>` | 下载导出产物（仅 pptx/pdf/txt） |
| `GET /images/<file>` | 访问生成的配图 |

## 🧪 测试

```bash
pytest -q          # 1028 条单测
```

测试通过 mock 截获 LLM 调用，覆盖：大纲解析、风格判定与回退、HTML 提取/重试、路径编码（防 XSS）、builder 版式等。

**pptx 链路的验收要先有素材**：`output/b_multislide.pptx` 与 COM 底图在 `.gitignore` 内，
新克隆上标了 `needs_pptx_src` / `needs_material` / `needs_pptx` 的用例会**静默 skip**。
先跑对应探针生成素材即可恢复：

```bash
python tools/probes/accept_m1.py       # 底图（M1/M2 验收的前置）
python tools/probes/accept_m2.py       # 覆盖率 + overlay 图（人工复核用）
python tools/probes/accept_m3.py       # 播放器（真浏览器）
python tools/probes/accept_m3_shot.py  # 步进序列截图
python tools/probes/accept_m6.py       # 导出 pptx 的字体（静态 + COM 渲染对照）
```

## ⚠️ 说明

- 默认模型走 `ZHIPUAI_CHAT_MODEL`（agnes 兼容通道），`ZHIPUAI_DESIGN_MODEL` 可单独指定更强的设计模型（注意余额充足，否则会 403）
- HTML 为 AI 生成产物，已做路径编码、iframe sandbox、零外链校验；本地单用户工具，请勿面向公网暴露
- **pptx 高亮链路要装 Office**：`faithful` 模式的保真底图只能由 PowerPoint 自己渲染，
  没有 Office 的机器只能用 `--mode redesign`。底图导出是**每次冷启动 PowerPoint**
  （10 页稿约 20–30 秒），不是稳态速度
- **导出的可编辑 pptx 不含图片与图表**：IR 里只有几何，没有图片路径/图表数据，
  这两类形状会跳过并在 stderr 点名
- `POST /api/pptx/import` 是新增的上传面；本轮只按"上传体积"兜底（60 MB），
  **解压总量/压缩比闸门尚未落地**（见 `KNOWLEDGE.md` 已知风险 M2）
- 图片循环、视觉校验、对话修改 API 刻意不改，聚焦“AI 自主设计”这一核心升级

## 📜 License

MIT
