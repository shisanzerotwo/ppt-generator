# 产物格式与常见错误

配套 `SKILL.md`。这里只写"产物长什么样"和"报错怎么读"。

## 目录布局

```
<out>/                       # 即 pptgen import --out 指定的目录
├── deck.json                # 唯一真源（形状 + 页序 + 画布）
├── units.json               # build_units 快照，仅供 QA/人工排查；程序不读它
├── bg/slide_1.png …         # COM 底图，1920 宽（或 --width 指定），页序从 1 起
├── player/index.html        # animate 产出（高亮播放器）
└── frames/step_0001.png …   # video --mode step 的中间帧（可由其再生）
```

## `deck.json`

```jsonc
{
  "schema": 1,
  "source_pptx": "D:/.../deck.pptx",   // 绝对路径
  "mode": "faithful",                  // faithful | redesign
  "export_width_px": 1920,             // 底图像素宽；px 坐标都基于它
  "width_emu": 12191695,               // 画布，EMU（唯一权威坐标）
  "height_emu": 6858000,
  "pages": [
    {
      "index": 0,
      "bg": "bg/slide_1.png",          // 相对 <out> 的 POSIX 风格相对路径；faithful 才有
      "shapes": [ /* ShapeInfo… */ ]
    }
  ],
  "skipped": [                          // 导入期被过滤的形状，供排查
    {"page_index": 0, "shape_name": "TextBox 3", "reason": "empty_text"}
  ]
}
```

**单位约定**（只读实现不要反向由 px 推 EMU）：

| 后缀 | 含义 |
|---|---|
| `*_emu` | OOXML 原生 EMU，唯一权威坐标 |
| `*_pt` | `emu / 12700` |
| `*_px` | `emu × export_width_px / width_emu`，固定在"导出底图像素空间" |

`skipped[].reason` 取值：`no_xfrm` / `zero_size` / `hidden` / `empty_text` /
`unsupported`（`kind=="other"`，保号供 redesign 参考）/ `group_no_xfrm` / `degenerate_group`。

## `units.json`

二维快照：`[{page_index, units: [Unit…]}]`。`Unit` 里：

- `rect` = 整个段落的 union 包围盒（滚动定位/命中测试用）；
- `lines` = **每行一个 tight 矩形**（渲染实体，播放器按它画高亮）；
- `kind` ∈ `title / body / bullet / cell / chart / picture / table`；
- `warnings` ∈ `vertical_anchor_inherited` / `autofit_scaled` / `justify_approximated`；
- `is_estimated=true` 表示字体度量走了降级估算（本机没读到微软雅黑）。

## 播放器 `player/index.html`

- **零 iframe、零外部依赖**：一张张底图 `<img>` 叠放 + 绝对定位的高亮 div。
- 底图 `src` 是**相对播放器文件**解析的：CLI 写出来的是 `../bg/slide_1.png`
  （播放器在 `<out>/player/` 下）。自己搬动文件时别把相对关系弄断。
- 程序驱动 API（`video --mode step` 依赖它）：

```js
window.hl.goto(page, step)   // 切页并揭示前 step 个单元；**同步改 DOM**，无过渡/无 rAF
window.hl.next() / back()
window.hl.state()            // -> {page, step, totalPages, totalUnits}
window.hl.ready              // Promise<void>，底图加载完成后 resolve
```

- 键盘：`空格`/`→` 下一步、`←` 回退；点画面也可以步进。
- **不自动校验底图是否存在**：缺图时页面全黑而 `ready` 照样 resolve。
  用 `shot_player`（`video --mode step` 内部走它）会先把关；自己写脚本要记得查
  `img.naturalWidth > 0`。

## 可编辑 pptx 的边界

`export` 只往返**文本与表格**：

| IR 里的形状 | 导出行为 |
|---|---|
| `text` | 写进母版占位符（真文本，可改） |
| `table` | `add_table` 落在正文区（可编辑） |
| `picture` | **跳过**——IR 只带几何，没有图片路径 |
| `chart` | **跳过**——IR 只带几何，没有 series/labels/values |

跳过的会在 stderr 点名，并在 CLI JSON 的 `data`（如启用）里可见。
中文一律显式写 `a:ea` 并改写 `theme1.xml`，保证渲染为**微软雅黑而非宋体**。

## 报错怎么读

`stdout` 永远是**一行** JSON：

```json
{"ok": false, "cmd": "import",
 "error": {"code": "NO_POWERPOINT", "message": "未检测到 PowerPoint…", "hint": "安装…"}}
```

- `message` 说**发生了什么**；`hint` 说**下一步做什么**。
- 完整堆栈只写 `stderr`（`INTERNAL` 时用它排查）。
- 退出码：`0` 成功｜`2` 参数错｜`3` 输入问题｜`4` 缺外部依赖｜`5` 内部错误。

```bash
out=$(python "$PY" animate out/ 2>/dev/null) || true
echo "$out" | python -c "import json,sys; d=json.load(sys.stdin); \
  print(d['error']['message'] if not d['ok'] else d['data']['player'])"
```
