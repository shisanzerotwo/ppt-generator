---
name: ppt-anim
description: 把现成的 .pptx 变成"会动的讲解"——用 PowerPoint 导出保真底图，叠加行级高亮锚点，生成可步进讲解的高亮播放器与 MP4；反向还能把产物导回**可编辑**的 pptx（母版+占位符）。当用户给出一个 .pptx 要求做讲解视频、逐条高亮演示、或要把已生成的 deck.json 变回可编辑 pptx 时使用。不负责"从主题生成新 PPT"（那是 ppt-maker 的活）。
---

# ppt-anim · pptx 双向链路

把 `.pptx` 变成会动的高亮讲解，再把产物导回可编辑的 `.pptx`。全部能力由本地
`pptgen` CLI 提供，**零常驻、无网络**（除 `--mode redesign` 会调 LLM）。

## 何时用

- 用户给了一份现成的 `.pptx`（汇报稿、讲义），想要**逐条高亮 + 步进讲解**的演示。
- 要把它导出成**讲解视频 MP4**（整页翻页，或按讲解单元逐步展开）。
- 要把先前生成的 `deck.json` 变回**可编辑的 pptx**（文字可选中可改，不是图片版）。
- 不适用：从一句主题从零生成新 PPT → 用 `ppt-maker`；
  已有的 HTML 设计稿 → 用本仓库的 `anim.py` 路线（那套是逐元素揭示，与高亮模式并存）。

## 怎么调

先定位这个封装脚本（它只是个转发器，逻辑在项目根的 `cli.py`）：

```bash
export PPTGEN_HOME="/d/GitHub/xiangmu/ppt-generator"   # 按实际路径改
PY="$PPTGEN_HOME/skill/ppt-anim/scripts/pptgen.py"
python "$PY" <子命令> [...]
```

### import —— pptx → 底图 + deck.json + 讲解单元

```bash
python "$PY" import deck.pptx --out out/ --width 1920
python "$PY" import deck.pptx --out out/ --pages 10      # 只导前 10 页
python "$PY" import deck.pptx --out out/ --no-com --mode redesign   # 不要底图，走 AI 重设计
```

- `--width` 只能是 `1280 / 1920 / 2560`，会**固化进 deck.json**（换宽度必须重导）。
- `--no-com` 与 `--mode faithful` **不能同时用**（faithful 的视觉保真完全依赖 COM 底图）。
- 默认 `--out` 是 `<项目根>/output/pptx_src/<稿名>/`。

### animate —— 生成高亮播放器

```bash
python "$PY" animate out/ --dim 0.25 --auto-ms 2000
```

产物 `<out>/player/index.html`。`--dim` 是"其余区域降暗比例"，须在 `(0, 0.6)`；
刚进页（还没揭示任何单元）不降暗，避免第一眼看到一片黑。

### video —— 导出 MP4

```bash
python "$PY" video out/ --sec 4 --fps 25 --fade 0.8            # 整页翻页
python "$PY" video out/ --mode step --sec 2 -o out/step.mp4    # 按讲解单元逐步展开
```

`--mode step` 会先驱动播放器逐单元截图（帧数 == Σ 讲解单元），再合成。
默认输出 `<out>/out.mp4`。`--fade` 必须 `0 < fade < --sec`。

### export —— deck.json → 可编辑 pptx

```bash
python "$PY" export out/deck.json -o out/deck.pptx
python "$PY" export out/deck.json --template 品牌母版.pptx
```

导出的文字是**真文本**（母版 + 占位符），在 PowerPoint 里可选中、可改、不破版；
中文显式写 `a:ea` 并改写主题字体，渲染为**微软雅黑而非宋体**。
导出后自动跑一遍 `qa.check_pptx` 自检，结果放在 JSON 的 `warnings` 里。

### deck —— 主题 → AI 设计稿（可选，需 LLM 渠道）

```bash
python "$PY" deck "人工智能如何改变教育" --out out/
```

## 产物在哪

```
<out>/
├── deck.json          # 唯一真源：形状 + 页序 + 画布
├── units.json         # build_units 的快照（QA/排查用，程序不读它）
├── bg/slide_1.png …   # COM 导出的保真底图（faithful）
└── player/index.html  # 高亮播放器（浏览器直接打开，零外部依赖）
```

字段级格式见 `references/formats.md`。

## 输出约定（很重要）

- **stdout 恒为一行 JSON**：`{"ok":true,"cmd":"import","data":{…}}` 或
  `{"ok":false,"cmd":"import","error":{"code","message","hint"}}`。
  解析 stdout 是安全的；进度与警告全在 **stderr**。
- **退出码**：`0` 成功｜`2` 参数错｜`3` 输入问题｜`4` 缺外部依赖｜`5` 内部错误。

```bash
json=$(python "$PY" import deck.pptx --out out/) || echo "失败：$(echo "$json" | jq -r .error.message)"
```

## 常见错误码

| code | 退出码 | 含义与对策 |
|---|---|---|
| `BAD_ARGS` | 2 | 参数组合非法（如 `--no-com` + `faithful`）。看 `hint` |
| `PPTX_NOT_FOUND` | 3 | 路径不存在；建议用**绝对路径** |
| `PPTX_UNREADABLE` | 3 | 非 zip / 损坏 / 非 pptx 包 / 包体异常 / 缺幻灯片尺寸 |
| `PPTX_ENCRYPTED` | 3 | 稿子加密了；先用 PowerPoint 去掉打开密码 |
| `PPTX_EMPTY` | 3 | 没有任何幻灯片 |
| `IR_MISMATCH` | 3 | deck.json 与底图不一致 / 底图缺失 → 重新 `import` |
| `TEMPLATE_INVALID` | 3 | 自定义模板不可用（需标准 .pptx 且至少一个版式） |
| `NO_POWERPOINT` | 4 | 没装 PowerPoint → 装 Office，或改用 `--mode redesign` |
| `NO_FFMPEG` | 4 | `winget install ffmpeg` |
| `NO_BROWSER` | 4 | 装 Chrome，或 `playwright install chromium` |
| `COM_EXPORT_FAILED` | 5 | PowerPoint 被占用或弹了对话框 → 关掉再试 |
| `INTERNAL` | 5 | 我们的 bug；附命令与 stderr 里的堆栈反馈 |

## 前置条件

- **Windows + Microsoft Office（含 PowerPoint）**：`faithful` 模式的保真底图只能由
  PowerPoint 自己渲染（python-pptx 拿到的文本框坐标与文字实际落墨位置无关）。
  没有 Office 的机器只能用 `--mode redesign`。
- `ffmpeg`（导视频）、Chrome/Edge（`--mode step` 截图）。
- 项目 venv 里已装 `python-pptx / pywin32 / fonttools / playwright / Pillow`。

## 已知限制

- 底图导出是**每次冷启动 PowerPoint**：10 页稿约 20–30 秒（不是 1.5s/页那种稳态速度）。
- 导出的可编辑 pptx **不含图片与图表**：IR 里只有几何，没有图片路径/图表数据，
  这两类形状会跳过并在 stderr 点名（见 `references/formats.md`）。
- 表格的合并单元格高亮会偏窄（IR 不带 span 信息）。
- 播放器截图前会校验底图是否真的加载出来，缺图直接报 `IR_MISMATCH` 而不是出黑帧。
