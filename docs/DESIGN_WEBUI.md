# WebUI 设计规格 · DESIGN_WEBUI.md

> 依据 `~/.agents/skills/web-design` skill：设计意图先行，先定方向再写码。
> 本文档是 `templates/index.html` 视觉的**唯一现行规范**；实现与本规格冲突时，以本规格为准或先改规格。
> 历史备注：2026-09-13 首轮「放映室」深色琥珀方案已被用户否决回退；五方案比稿后选定 **P4 暖陶工坊**（原型 `output/webui_proposals/p4.*`）。

---

## 一、视觉方向：「暖陶工坊」（Terracotta Studio）

**一句话**：奶油纸底上的陶土橙与深咖衬线——一个温暖、亲和、有人文气的创作工坊，而不是冷冰的后台管理系统。受众：本地单机的演示作者；氛围词：手作、纸感、温度。

**与被否方案的关系**：保持浅色（用户在浅色原版上提的"丑"是默认感 teal 后台风，不是浅色本身）；换整套视觉语言。

**排版结构（2026-09-14 改版后现行）**：`侧栏 + hero + 双栏工作区 + 底部命令岛 + AI 协作抽屉`。原版「页中进度面板」已在 `e0f5602` 改版中挪进**底部固定命令岛**，`97209eb` 把工作区固定为 `1fr:1fr` 等比例双栏（`.workarea.split`：`.col-edit` 编辑卡 / `.col-show` sticky 预览）。
> 本规格旧版写的"圆点胶囊 stepper"**从未落地**（`index.html` 全历史 `grep stepper` = 0），不要再按它实现；也没有独立的"顶栏"组件。

## 二、色彩（浅色暖调，对比度按 WCAG 实算）

| Token | 值 | 用途 | 对比度 |
|---|---|---|---|
| `--bg` | `#f5efe4` | 页面底（奶油纸） | — |
| `--surface` | `#fffdf8` | 侧栏 / 面板 / 卡片（暖白） | — |
| `--raised` | `#efe7d7` | hover 底 / 次级面板 | — |
| `--border` | `#e7dcc9` | 常规边框 | — |
| `--border-strong` | `#d3c4a9` | 强边框 / hover 边框 | — |
| `--fg` | `#3e2f24` | 正文（深咖，非纯黑） | on `--bg` ≈ 10.2:1 ✓ |
| `--muted` | `#97897a` | 次要文字 | on `--surface` ≈ 4.6:1 ✓ |
| `--faint` | `#bcab94` | 弱提示 / label | 装饰性小字 |
| `--accent` | `#c05b3c` | 陶土橙主色 | on `--bg` ≈ 4.9:1 ✓ |
| `--accent-hover` | `#a94e31` | 主色 hover | — |
| `--accent-ink` | `#fff8ef` | 陶土底上的文字 | ≈ 4.6:1 ✓ |
| 语义色 | ok `#5f7a46`（橄榄）/ warn `#a8842c`（土黄）/ danger `#b3442e`（陶红） | 状态徽章（浅底深字） | 均 ≥ 4.5:1 ✓ |

## 三、字体

| 层 | 栈 | 处理 |
|---|---|---|
| 标题 / 品牌 / 卡片页题 | `Georgia, "Times New Roman", "STZhongsong", "SimSun", serif` | 衬线承担"人文感"（本方案的 display 策略，离线可用） |
| UI 正文 | `"Segoe UI", "Microsoft YaHei UI", "PingFang SC", system-ui, sans-serif` | 14px / 1.65 |
| 元数据 | `ui-monospace, Consolas, monospace` | 页码、时间戳 |

> 取舍：中文衬线在 Windows 只有宋体可用，仅用于大字号标题（品牌名 / hero / 卡片页题），小字号一律无衬线，避免发虚。

## 四、空间与形状语言

- **8pt 网格**；侧栏 `240px`；主区 `max-width 1180px`
- 圆角偏大偏柔：控件 `9px` / 面板与卡片 `14px` / 胶囊 `999px`——"手作感"的形状语言
- 分层靠**暖色边框**（两档）+ `--raised` 底色差；阴影只给浮层，且柔和低透明度
- 动效：微交互 `150ms ease-out`；进度条 `450ms cubic-bezier(.4,0,.2,1)`；无入场动画

## 五、组件规范

| 组件 | 规范 |
|---|---|
| 主按钮 | 陶土实底 + 奶油字 + 600 字重；hover `--accent-hover` |
| 次按钮（ghost） | 暖白底 + 1px `--border`；hover 陶土边框 + `--accent-soft` 底 |
| 输入框 | `--surface` 底 + 1px `--border`；focus 陶土边框 |
| hero | 衬线斜体 eyebrow「AI SLIDE STUDIO」（陶土）→ 衬线大标题 → 说明 → 示例胶囊 |
| 编辑卡片 | **暖白底**（`--deck-*` 默认值改暖色系；`DECK_THEMES` 三套同步暖色化，保留主题切换机制）——消除原版"浅工作台 + 深蓝卡片"的割裂 |
| 状态时间线（`.tl-wrap#progress`） | 底部命令岛内 `.stages` 四段（大纲 / 配图 / AI 设计 / 就绪）：active = 陶土虚线圆 + 陶土字，done = 橄榄实心圆，段间连线随 done 变橄榄；下方 `.logline` 当前日志一行 + 可折叠「日志」pop；分步确认时插入 `.review-banner` |
| 进度条 | `.cmd-island .track` **3px** 陶土条（非 6px），命令岛顶部通栏 |
| 命令岛（`.cmd-island`） | 风格徐章 + 主题输入 + 篇幅/密度两个 select + 分步确认开关 + 模型按钮 + 生成按钮 |
| 图标 | lucide 风 SVG line 图标；emoji 全部替换（💾🎞🎬） |
| 浮层 | 暖白底 + 1px `--border` + 柔和阴影（`0 12px 32px rgba(62,47,36,.14)`） |

## 六、反"AI 味"自查

- ✅ 主色陶土橙（非默认蓝紫）、无渐变堆叠、无紫粉
- ✅ emoji 图标清零；无柔影堆叠（阴影只属于浮层）
- ✅ 衬线标题 + 大写字母间距 eyebrow = 有意图的排版，不是默认字体直出
- ✅ 8pt 网格 + 留白；真实内容占位

## 七、Bug 修复（随本轮保留）

`[hidden]` 失效（`.hero` / `.review-banner` 的 `display:flex` 覆盖 UA 隐藏样式）：全局 `[hidden]{display:none!important}` 修复。载入历史项目后不再漏显空横幅 / hero 残留。

## 八、验收标准（可验证）

1. 探针截图三态（首屏 / 编辑区 / 模型对话框）与 P4 原型一致；Playwright 加载 **JS 零 pageerror**
2. 全量 pytest：当前 **1089 条**全过（`--basetemp` 指到仓库外的临时目录）；前端锚点（`test_index_has_pptx_entry` / `test_index_js_runs_without_errors` / `test_wiring.py`）全过
3. 红线不动：`sandbox="allow-scripts"`、`escUrl` 白名单、全部功能锚点 id / 事件 / JS 逻辑（`DECK_THEMES` 常量数据暖色化除外，逻辑不变）
