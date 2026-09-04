# PPT 生成器（ppt-generator）

AI 驱动的 PPT 生成器：输入主题，自动完成「大纲 → 生图 → 视觉校验 → **AI 自主设计 HTML 幻灯片** → 智能选风格」全流程。核心亮点是**让 LLM 根据内容自主设计每页布局**，而非固定模板——数据页画图表、概念页出卡片、封面做视觉冲击、时间线页做时间轴，真正“按内容适配”。

## ✨ 功能特性

- **LLM 自主设计 HTML**：直接把大纲 + 图片交给 LLM 生成完整单文件 HTML 幻灯片，每页布局因内容而异，纯内联 CSS + SVG，零外部依赖，可离线打开、可打印 PDF
- **AI 智能选风格**：去掉固定主题下拉框，AI 根据主题/内容从风格库自主判断视觉基调（深空科技 / 极简商务 / 清新浅色 / 暖调人文 / 活力渐变 / 自然墨绿）
- **完整的生成流水线**：大纲（章节化、8~10 页）→ 逐页生图 → 视觉校验（提议者-审核者闭环，不契合自动改词重生）→ LLM 设计 HTML
- **对话式修改**：ready 后可输入指令（如“第 3 页更简洁 / 整篇换商务风”）重新生成，自动重设计 HTML
- **多格式导出**：主产物 HTML（AI 设计稿，可预览/打印 PDF），另有 pptx（可编辑版）、PDF、大纲 txt 降级方案
- **打印分页兑底**：LLM 漏写 `@media print` 时，`html_gen` 自动注入分页规则（1280×720 每页一张、`page-break-after`），保证浏览器打印/导出 PDF 排版正确
- **历史设计稿**：侧栏「历史设计稿」列出最近 30 份已生成的设计稿（按时间倒序，含标题与时间），点击即可在新标签打开
- **开发友好**：Web 界面实时显示阶段、日志、风格徽章；65 条 pytest 单测

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

## 📂 核心模块

| 模块 | 职责 |
|---|---|
| `outline.py` | 主题/文档 → 章节化大纲 JSON（含自评迭代、布局多样化） |
| `style.py` | AI 智能选风格：从 6 种预设风格库判断，决定色板/气质/设计指引 |
| `image_gen.py` | 按提示词生成每页配图 |
| `critic.py` | 视觉校验 + 对话式修改大纲（提议者-审核者闭环） |
| `html_gen.py` | **核心**：LLM 自主设计每页 HTML 布局，注入风格，容错重试 |
| `builder.py` | pptx 降级导出（可编辑版） |
| `app.py` | Flask 后端 + Web 工作台（异步生成、状态机、导出接口） |
| `templates/index.html` | 前端：步骤条、AI 设计稿 iframe 预览、重新设计、对话修改 |

## 🔌 API 一览

| 接口 | 说明 |
|---|---|
| `POST /api/generate` | 按主题生成 |
| `POST /api/import` / `api/import_file` | 从文档文本/文件生成 |
| `POST /api/refine` | 对话式修改并重设计 |
| `POST /api/redesign` | 重新触发 AI 设计（重生成 HTML） |
| `GET /api/status` | 轮询状态（phase / slides / html_path / style_name / log） |
| `GET /api/decks` | 历史设计稿列表（按修改时间倒序，上限 30，含标题/链接/时间） |
| `POST /api/export` `/export_pdf` `/export_html` `/export_txt` | 多格式导出 |
| `GET /decks/<file>` | 访问 AI 设计稿 HTML |

## 🧪 测试

```bash
pytest -q          # 65 条单测
```

测试通过 mock 截获 LLM 调用，覆盖：大纲解析、风格判定与回退、HTML 提取/重试、路径编码（防 XSS）、builder 版式等。

## ⚠️ 说明

- 默认模型走 `ZHIPUAI_CHAT_MODEL`（agnes 兼容通道），`ZHIPUAI_DESIGN_MODEL` 可单独指定更强的设计模型（注意余额充足，否则会 403）
- HTML 为 AI 生成产物，已做路径编码、iframe sandbox、零外链校验；本地单用户工具，请勿面向公网暴露
- 图片循环、视觉校验、对话修改 API 刻意不改，聚焦“AI 自主设计”这一核心升级

## 📜 License

MIT
