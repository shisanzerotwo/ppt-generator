# 修复报告 · KNOWLEDGE.md 风险表遗留 9 项（F1/F2/F3/M1-M4/L1/L4）

> 角色：修复 agent（ppt-fix）｜日期 2026-09-13｜基线 commit `e52504f`（1028 tests collected）
> 上游：`docs/TASK_FIX.md`（任务卡）、`KNOWLEDGE.md` 风险表、`docs/TEST_REPORT.md` §6、
> `docs/AUDIT_REPORT.md` §2、`tools/probes/s1_probe_v2.py`（探针范式）
> **结果：8 项修复 + 1 项（F3）附数据放弃；全量 1066 passed。**

## 汇总

| # | 问题 | 状态 | commit | 新增用例 |
|---|---|---|---|---|
| F1 | `export_pages` 拆掉调用方 COM apartment | ✅ 修复 | `0fe912e` | 5 |
| M2 | zip bomb 无闸门 | ✅ 修复 | `544e7b1` | 5 |
| M4 | 缺 `p:sldSz` → 错误码退化成 INTERNAL(5) | ✅ 修复 | `2e72493` | 1 |
| M3 | 形状被静默丢弃 | ✅ 修复 | `692d83d` | 7 |
| M1 | `a:br` 几何与 `qa` 分叉 | ✅ 修复 | `9b4e3c0` | 9 |
| F2 | `wrap="none"` 导致精度退化为段级 | ✅ 修复 | `01cb207` | 4 |
| L4 | 播放器不校验底图 → 黑帧静默通过 | ✅ 修复 | `8bd539c` | 3 |
| L1 | 老 `.ppt` 被误判成「已加密」 | ✅ 修复 | `da574f8` | 4 |
| F3 | 冷调用 2.6s/页 → 复用 PowerPoint 实例 | ⛔ **放弃**（附实测） | `238eff4` | 0 |

回归：`1028 → 1066`（+38 条），每一步之后都跑全量、全绿。
硬约束遵守：**S1 的 `we_opened` 守卫与 D1 的 `had_powerpoint`/进程数守卫原样保留**（F1 的改动
只影响"我们是否收尾自己那一份 apartment"，不碰这两道）；项目代码里 `Visible =` **零命中**
（`grep -rn "Visible\s*=[^=]" --include=*.py .` 的命中全在 `.venv/` 的第三方包里）；
`anim.py` / `builder.py` 对基线 `e52504f` 零 diff。

---

## 1 · F1（中高）`export_pages` 拆掉调用方的 COM apartment

### ① 根因（比原报告更精确）

原报告归因为"引用计数不对称"。实测找到了确切的加/减配对：

```
干净状态 → CoInitializeEx 返回 0x0 (S_OK)
import win32com.client
之后     → CoInitializeEx 返回 0x1 (S_FALSE)   ← **导入这个模块本身就把 apartment 初始化了一次**
```

`export_pages` 原先无条件 `pythoncom.CoInitialize()` + `finally: CoUninitialize()`。
那句 `CoUninitialize` 减掉的正是 `import win32com.client` 加的那一份，计数归零 →
**整个线程的 apartment 被拆**，调用方之后任何 COM 调用报 `CO_E_NOTINITIALIZED`
（连 `Scripting.FileSystemObject` 都建不出来）。这也解释了它为什么是**偶发 1/3**：
取决于此前有没有别的东西也加过计数。

### ② 修法

新增 `pptx_io._co_initialize() -> bool`：用 **ctypes 直接问 ole32 要 HRESULT**
（pywin32 的 `pythoncom.CoInitializeEx` 恒返回 `None`，读不到返回值）：

| HRESULT | 含义 | 我们怎么做 |
|---|---|---|
| `S_OK(0)` | 这一份是我们加的 | 结束时由我们 `CoUninitialize` |
| `S_FALSE(1)` | 调用方已初始化 | **一次都不减** |
| `RPC_E_CHANGED_MODE(0x80010106)` | 线程是别的 apartment 模型 | **更不该动** |

规则一句话：**不是自己加的计数，就不要减。**

**被否决的替代方案**：
- *只在 CLI 侧重新 `CoInitialize()` 兜底*（测试 agent 的建议）—— 治标：调用方的
  Flask worker、后台线程照样被拆，而且"被拆之后要记得重开"是很容易漏的约定。
- *把 `pythoncom.CoInitialize()` 换成 `pythoncom.CoInitializeEx()`*（任务卡字面建议）——
  **行不通**：pywin32 这个函数不返回 HRESULT，拿不到 S_OK/S_FALSE 的区分。任务卡说
  "以实测为准"，这就是实测结果。

### ③ 修复前/后对照实测

```
PYTHONIOENCODING=utf-8 ./.venv/Scripts/python.exe tools/probes/fix_f1_apartment.py
```

探针给 `pythoncom.CoUninitialize` 装间谍，用**确定性不变量**
`调用次数 == (we_initialized ? 1 : 0)` 判定；两组各跑在独立子进程里，
保证 apartment 初态干净（同进程里 pywin32 的 Dispatch 会悄悄初始化 apartment
且不配对释放，把实验污染掉 —— 本探针第一版就因此误报过 FIXED）。

**修复前**

```
GROUP no-init     export_pages 成功，10 页 → COUNINIT_CALLS=1
GROUP caller-init 调用方 CoInitializeEx → 0x00000001（已初始化）
                  export_pages 成功，10 页 → COUNINIT_CALLS=1   ← 减的是别人加的计数
判定：**REPRODUCED** —— 调用方已初始化时它仍会 CoUninitialize（F1 成立）
```

**修复后**

```
  caller-init  we_initialized=False  CoUninitialize 调用=0  应为 0  OK
  no-init      we_initialized=False  CoUninitialize 调用=0  应为 0  OK
  两个子进程事后都仍能用 COM：可用
判定：**FIXED** —— 每个分支都只收自己开的计数
```

⚠️ **必须说明的边界**：真机环境下 `we_initialized` **恒为 False**（因为
`import win32com.client` 早就把 apartment 初始化过了），所以 `S_OK` 那条分支
在真实运行里走不到。它由**单测打桩**覆盖（3 条参数化 + 2 条让导出走到 finally
分别验证两个分支）。

### ④ 新增用例数：5

`test_co_initialize_reports_ownership`（3 参数：S_OK / S_FALSE / RPC_E_CHANGED_MODE）、
`test_export_pages_uninitializes_only_when_it_owns_it`、
`test_export_pages_leaves_callers_apartment_alone`。

---

## 2 · M2（中）zip bomb 无闸门

### ① 根因

`read_pages` 直接 `Presentation(path)`，而 python-pptx 一解析就把条目读成 Python
字节串 —— **"读"是不可逆的**：把 64 MiB 全零塞进 `[Content_Types].xml`（包只有
64.3 KB，压缩比 1029:1），字节先展开进内存，之后才因 XML 解析失败被拒。

### ② 修法

`read_pages` 在 `Presentation()` **之前**调 `_check_package_safety(path)`：读 zip
中央目录（`infolist()`，**零解压**）累计声明体积与单条压缩比：

| 阈值 | 取值 | 依据 |
|---|---|---|
| 解压总量 | 512 MB | 真实汇报稿含底图也就几十 MB；512 MB 已极宽松 |
| 单条压缩比 | 200:1 **且**该条 ≥ 8 MiB | 纯 XML 文本压缩比通常 3~20；加绝对体积下限是为了**不误伤**"某个几百字节的小 XML 恰好压得很好" |

不是 zip / 读不出目录时**放行**，归类仍归 `_classify_bad_package`（各司其职）。

**错误码取舍（与任务卡的冲突，已按契约处理）**：任务卡写 `PPTX_UNSAFE`，但契约 v2
§3.1 已把"包体超限"归入 `PPTX_UNREADABLE`，且 §9 明令"实现期不得自行扩写"错误码
（`cli.py` 的退出码表也只认既有码）→ 用 `PPTX_UNREADABLE`（退出码 3）。
文案按 §9 的示例写成 `包体异常：<条目> 单条压缩比 1029:1（64 MB），超过上限 200:1`。

### ③ 修复前/后对照实测

```
PYTHONIOENCODING=utf-8 ./.venv/Scripts/python.exe tools/probes/fix_m2_zipbomb.py
```

| | 报错 | **Python 侧内存峰值** |
|---|---|---|
| 修复前 | `PPTX_UNREADABLE`「不是有效的 .pptx 包」（解析失败后才拒） | **141.5 MB** |
| 修复后 | `PPTX_UNREADABLE`「包体异常：[Content_Types].xml 单条压缩比 1029:1…」 | **0.0 MB** |

（样本文件 65,874 B = 64.3 KB。）正常稿不受影响：`b_multislide.pptx` 仍读出 10 页。

### ④ 新增用例数：5

高压缩比被拒 / 总量超限被拒 / **小条目高压缩比不误伤**（阳性对照的另一面）/
非 zip 放行给分类器 / 正常稿放行。

---

## 3 · M4（中）缺 `p:sldSz` → 错误码从 3 退化成 5

### ① 根因

`prs.slide_width` 可以是 `None`（包里没有 `<p:sldSz>`），而 `int(prs.slide_width)`
那行在 `Presentation()` 的 `try` **之外** → 裸 `TypeError` 冒到 CLI，被归成
`INTERNAL(5)`。契约 §9 的文案纪律是"message 说发生了什么、hint 说下一步做什么"，
而 `INTERNAL` 的语义是"这属于 bug，请附命令与 deck.json" —— 把**用户稿子畸形**
说成"我们的 bug"，把排查方向引偏。

### ② 修法

读尺寸前显式判空，抛 `PptxError("PPTX_UNREADABLE", "…（缺少幻灯片尺寸定义 p:sldSz）",
"用 PowerPoint 打开后另存一次，或换一份稿子")`。

### ③ 修复前/后对照实测

```
PYTHONIOENCODING=utf-8 ./.venv/Scripts/python.exe tools/probes/fix_m4_sldsz.py
```

样本 = 真 pptx + 正则删掉 `presentation.xml` 里的 `<p:sldSz/>`（1 处），
自检 `prs.slide_width = None` ✓。

| | `read_pages` | CLI 退出码 / code |
|---|---|---|
| 修复前 | `TypeError: int() argument must be … not 'NoneType'` | `5` / `INTERNAL`（"这属于 bug，请附命令与 deck.json"） |
| 修复后 | `PptxError code=PPTX_UNREADABLE` | `3` / `PPTX_UNREADABLE` |

### ④ 新增用例数：1

`test_missing_sldsz_is_unreadable_not_internal`（先断言 `slide_width is None` 作为前提）。

---

## 4 · M3（中）形状被**静默丢弃**

### ① 根因

契约 §4.3 明写 `inner_w_pt <= 0` → "跳过该形状（**记 warning**，不产出 Unit）"。
实现只做到"不产出 Unit"——`return` 一句就走，`build_units` 连个收集器参数都没有。
`read_pages` 有 `skipped` 出口，布局期丢弃却**一条都进不去** CLI 的 JSON：
用户在播放器里只看到"有些文字没高亮"，排查时没有任何线索。

### ② 修法

`build_units(page, …, skipped=None)` 增加可选收集器，记
`{page_index, shape_id, shape_name, reason}`；reason ∈
`no_inner_width` / `no_inner_height` / `zero_cell` / `col_array_short` / `row_array_short`。
`cli import` 把它并入 `data["skipped"]` 与 `data["warnings"]`，并逐条打到 stderr。

**顺带覆盖 L2**：同一循环里"表格列宽/行高数组短于单元格矩阵"也是同一种静默丢弃
（审计 L2，虽不在本次 9 项内）。既然这次修复的主题就是"别再静默丢弃"，
在同一个循环里留一半没说不过去，故一并记上。

### ③ 修复前/后对照实测

```
PYTHONIOENCODING=utf-8 ./.venv/Scripts/python.exe tools/probes/fix_m3_skipped.py
```

样本：宽 100000 EMU、左右边距各 91440 EMU → `inner_w_pt = -6.53`（≤ 0）。

| | build_units | CLI |
|---|---|---|
| 修复前 | `TypeError: build_units() got an unexpected keyword argument 'skipped'` | — |
| 修复后 | `skipped = [{'page_index': 0, 'shape_id': 3, 'shape_name': 'TooNarrowBox', 'reason': 'no_inner_width'}]` | 退出码 0，`skipped = 1`，`warnings = ['no_inner_width×1']` |

### ④ 新增用例数：7

6 条 `hl_layout`（收集器记录 / 不传收集器不抛 / 高度分支 / 正常形状不上报 /
零尺寸单元格 / 列数组截短）+ 1 条 CLI（JSON 与 stderr 都能看到形状名）。

---

## 5 · M1（中）`a:br` 段落几何与 `qa` **分叉**

### ① 根因

契约 §2.3 让 `a:br` 读成 `"\n"` 存进 `ParaInfo.text`，但 `qa._tokenize` 把 `"\n"`
归进 **cjk 分支**（既不是空格也不是 alnum）→ 被当成一个 **1.0em 宽的字形**，
而不是 PowerPoint 会断行的地方。于是：

- 真正产出几何的 `_paragraph_lines` 私自先按 `"\n"` 硬拆（渲染对，但与 qa 差 +1 行）；
- 契约 §4.2 那条"防漂移"等价断言只盖 `wrap_lines` —— **盖不住真正被调用的那个函数**。

`qa` 自己的 docstring 写着它的立身之本是"度量与渲染同源、算出来的行数逼近
PowerPoint 实际排出的行数" —— 把硬换行当字形，这条就不成立了。

### ② 修法（任务卡选项 a：抽公共判据，两处同源）

- 新增 `qa._is_hard_break(tok)`，**两边共用**；
- `qa._measure_lines_ex` 遇硬换行 → 收行、另起一行（行尾挂起空格随行丢弃）；
- `hl_layout.wrap_lines` 用同一个判据做同样的事；
- `_paragraph_lines` 因此**退回契约 §4.3 的字面写法**（删掉那次私自硬拆）。

**为什么选 (a) 而不是 (b)"声明两者语义本就不同"**：(b) 只是把审计指出的问题写下来，
断言依旧盖不住 `_paragraph_lines`；而两处本来就是**同一件事**（都在建模 PowerPoint
怎么排字），分歧的根源是 `qa` 的模型在这里错了。

### ③ 修复前/后对照实测

```
PYTHONIOENCODING=utf-8 ./.venv/Scripts/python.exe tools/probes/fix_m1_hardbreak.py
```

| 用例 | 修复前 `_paragraph_lines` / `wrap_lines` / `qa` | 修复后 |
|---|---|---|
| 两行 `"第一行\n第二行"` | 2 / 1 / 1 ❌ | 2 / 2 / 2 ✅ |
| 段尾换了行 `"只有一行\n"` | 2 / 1 / 1 ❌ | 2 / 2 / 2 ✅ |
| 连续两个硬换行 `"甲\n\n乙"` | 3 / 1 / 1 ❌ | 3 / 3 / 3 ✅ |
| 硬换行 + 软折行 | 5 / 4 / 4 ❌ | 5 / 5 / 5 ✅ |
| 行首带空格 `"甲\n 乙"` | 2 / 1 / 1 ❌ | 2 / 2 / 2 ✅ |

**对既有基线的影响**：`qa` 的行为变了（含 `"\n"` 的段落会多算行），但审计实测
仓库内三份稿的 `a:br` 均为 **0**，真实素材的溢出告警不受影响（全量 1066 全绿佐证）。

**同步修正了独立测试 agent 的两条不变量**（它们写在"`\n` 是字符"的旧语义下）：
`test_wrap_lines_loses_no_non_space_char` 现在把 `\n` 与空格一样剥掉再比；
`test_wrap_lines_no_leading_space_and_no_empty_line` 的"无空行"只对不含 `\n` 的语料
断言（硬换行**合法地**产生空行）。

### ④ 新增用例数：9

6 条参数化硬换行等价（`_paragraph_lines` ↔ `wrap_lines` ↔ `qa` 三方一致）
+ `qa` 认硬换行 + 硬换行把文本拆到不同行 + 段尾硬换行留空行。

---

## 6 · F2（中）`wrap="none"` 导致"行级高亮"退化为**段级**

### ① 根因

python-pptx 的 `shapes.add_textbox()` **默认写 `wrap="none"`**，读出来就是
`word_wrap=False`；这类框按契约 §4.3 不折行 → 一个段落恒出 1 行，"行级 tight 高亮"
在那里**就是段级**。`b_multislide.pptx` 有 13/22 个文本框如此，所以 M2 的 coverage
结论在这份素材上等价于段级 —— `wrap_lines` 那层只被单元测试覆盖过。
其中最危险的是"文本还超出框宽"的那些：要么真的溢出、要么作者的折行本意被关掉了，
而我们**不猜** PowerPoint 怎么折（契约的选择），但至少该如实声明精度降级。

### ② 修法

`build_units` 在 `word_wrap is False` 且任一行估算宽度 > 框内宽时，给该 Unit 加
`nowrap_overflow_degraded` 警告；`cli import` 把 **Unit 级**警告汇总成 `名字×条数`
并入 `data["warnings"]`，并打一条 stderr。

**⚠️ 任务卡的验收在这一稿上无法触发（已实测确认，未硬做）**：它要求"对
`b_multislide.pptx` 跑 import 应在 warnings 里看到降级提示"，但该稿 13 个
`wrap="none"` 框的文本**全都比框窄**：

```
页 形状        wrap  inner_w( pt)  首段估宽(pt)  超宽?
 1 TextBox 2  False    801.6         528.0      False
 1 TextBox 3  False    801.6         108.0      False
 …（13 行全部 False）…
10 TextBox 2  False    830.4         432.0      False
```

→ "估宽 > 框宽"一处都不成立，警告**不该**出现。我保留任务卡给的条件（它才是
"几何可能算错"的有效信号），改用**合成阳性对照**证明判据本身是通的。若对全部
`wrap=none` 一律告警，这一稿会平白多出 13 条无害警告，反而稀释信号。

### ③ 修复前/后对照实测

```
PYTHONIOENCODING=utf-8 ./.venv/Scripts/python.exe tools/probes/fix_f2_nowrap.py
```

| | b_multislide（13 个 wrap=none 框） | 合成阳性对照（2 英寸框 + 28pt 长中文） |
|---|---|---|
| 修复前 | `Unit.warnings` 无标记，CLI `warnings = ['empty_text×18']` | — |
| 修复后 | 仍无标记（上面已说明：条件确不成立） | **1 个单元带 `nowrap_overflow_degraded`** |

### ④ 新增用例数：4

超宽→告警 / 放得下→不告警 / `wrap=True`→不告警且不影响其它警告 /
CLI JSON 与 stderr 都能看到。

---

## 7 · L4（低）播放器不校验底图 → 黑帧静默通过

### ① 根因

底图 `<img>` 没有 `onerror` 处理。缺图时浏览器只是不显示，而 `window.hl.ready`
**仍然 resolve** → 截图拿到黑帧 → 安静地喂进 ffmpeg。调用侧（`shot_player` /
`cli animate`）确实会校验，但**播放器对自己"能不能正确显示"完全不设防**：
任何自己写脚本驱动 `goto` 的调用方都会踩。

### ② 修法

- 每张底图**先挂回调再设 `src`**（避免竞态）：`onerror` 记进 `bgFailed`；
- `hl.ready` 收齐所有底图后，**若有失败则 throw**（页面照常渲染，人看得见）；
- 新增 `hl.failed()` 返回缺图清单；
- 模板尾部 `hl.ready.then(fit).catch(…)` 接住 reject，并给人一个可见的红色横幅
  （`.bgwarn`），而不是一片没有解释的黑屏；
- `shot_player` 里吞掉浏览器异常，仍统一给带错误码的中文 `PptxError`。

### ③ 修复前/后对照实测

```
PYTHONIOENCODING=utf-8 ./.venv/Scripts/python.exe tools/probes/fix_l4_bgcheck.py
```

样本：2 页播放器，第 2 页底图**故意缺失**。

| | `window.hl.ready` | `hl.failed()` | `shot_player` |
|---|---|---|---|
| 修复前 | **resolve（静默通过）** | — | 靠自己那份 check 才拦下 |
| 修复后 | **reject** → `底图加载失败 1/2：bg/slide_2_missing.png` | `['bg/slide_2_missing.png']` | 仍抛 `PptxError(IR_MISMATCH)`，文案不变 |

**同步修正了独立测试 agent 的 `test_adversarial_player_in_real_browser`**：它的底图
列表里有一个**故意缺的** `bg/missing.png`（用来试路径处理），修 L4 后会让 `ready`
reject —— 那正是底图校验该做的事，与该案要验的转义/注入无关，故改成在 JS 侧接住 reject。

### ④ 新增用例数：3

模板含 `onerror`/`failed` 出口 / 缺图 reject 且 failed 清单正确 /
阳性对照（底图齐全时正常 resolve，证明不是"永远 reject"）。

---

## 8 · L1（低）老 `.ppt`（OLE2）被误判成「PPTX 已加密」

### ① 根因

`_classify_bad_package` 只看文件头 `D0CF11E0`，而那是 **OLE2 复合文档的通用头** ——
老 `.ppt` / `.doc` / `.xls` 全是它。于是用户拿一份 `.ppt` 过来被告知"请先去掉打开密码"，
而那份文件根本没有密码；契约 §9 里本该给的正确提示（"确认是 .pptx（非 .ppt/.pdf）"）
也被这条分支抢走。真正的加密 OOXML 同样是 OLE2，区别在于它里面有一个特征流
`EncryptedPackage`（CFB 目录以 UTF-16LE 存名）。

### ② 修法（两层判据）

1. 扩展名 ∈ `{.ppt,.doc,.xls,.pps,.pot}` → `PPTX_UNREADABLE` +「另存为 .pptx」；
2. 否则探 `EncryptedPackage` 特征流（读前 1 MiB + 末 8 MiB，**零 CFB 解析**）：
   确认为"没有" → `PPTX_UNREADABLE`（多半是老 `.ppt` 改了扩展名，同样给「另存为」）；
   读到 / 读不出 → `PPTX_ENCRYPTED`（保守当加密，与旧行为一致，不因修 L1 而放松）。

### ③ 修复前/后对照实测

```
PYTHONIOENCODING=utf-8 ./.venv/Scripts/python.exe tools/probes/fix_l1_ole2.py
```

| 样本 | 修复前 | 修复后 |
|---|---|---|
| a) `legacy.ppt`（OLE2 + `PowerPoint Document`） | `PPTX_ENCRYPTED`「请去掉打开密码」❌ | `PPTX_UNREADABLE`「请在 PowerPoint 里打开它，另存为 .pptx」✅ |
| b) `encrypted.pptx`（OLE2 + `EncryptedPackage`） | `PPTX_ENCRYPTED` ✅ | `PPTX_ENCRYPTED` ✅（不变） |
| c) `renamed.pptx`（老 `.ppt` 改名） | `PPTX_ENCRYPTED` ❌ | `PPTX_UNREADABLE`「很可能是把旧版 .ppt 改了扩展名」✅ |

**顺带修正两处「把误判锁成期望行为」的假测试**（任务卡点名的实例）：
`tests/test_pptx_io.py::test_encrypted_code` 原来只写一个 OLE2 头就断言
`PPTX_ENCRYPTED`；`tests/test_pptx_independent.py::test_pathological_inputs_...`
的 `ole` 样本同一毛病。两者都改成带 `EncryptedPackage` 特征流的**真**样本，
用例意图不变但样本不再说谎。

### ④ 新增用例数：4

老 `.ppt` → UNREADABLE / 改名 `.ppt` → UNREADABLE / 读不出流名时仍保守 ENCRYPTED /
`_ZIP_MAGIC` 大小写回归（见「发现的新问题」#1）。

---

## 9 · F3（中）冷调用 2.6s/页 → 复用 PowerPoint 实例 —— **放弃**

### ① 根因

`export_pages` 每次调用都 `DispatchEx` 起一个 PowerPoint、结束时 `Quit`。
成本构成：启动 ~5.8s + `CoUninitialize` ~2.7s + Quit 延迟退出。

### ② 修法：**不做**（任务卡允许），理由与数据如下

### ③ 实测数据

```
PYTHONIOENCODING=utf-8 ./.venv/Scripts/python.exe tools/probes/fix_f3_reuse.py
```

```
  A 冷调用（现状）         25.44s   2.54s/页
  B 常驻·首次              6.96s   （含一次性启动 5.76s）
  C 常驻·复用同一个实例     1.30s   0.13s/页
  → 复用相对冷调用提速 19.5×，每次省 24.1s —— **收益是真的**
```

**但方案落不了地，两条实测理由**：

1. **CLI 收益为 0**：`pptgen import` 每次都是新进程、只导一次，常驻实例省不到钱，
   只留下一个常驻 `POWERPNT.EXE` 的成本与新的失败模式。而 CLI 正是
   D14 定义的主要面向（agent 通过 bash 调用）。
2. **唯一的受益者（工作台）用不了**：`app.py` 的 pptx worker 是
   `threading.Thread(target=...)` —— **每次请求新起一个线程**。实测：

```
== D 线程亲和性：在**另一个线程**里用同一个 app ==
  另一个线程（已自建 apartment）使用同一个 app → **失败**：
    com_error: (-2147417842, '应用程序调用一个已为另一线程整理的接口。', ...)
  → 结论：模块级单例对 Flask 的「每请求新线程」模型直接不可用
```

`-2147417842` = `RPC_E_WRONG_THREAD`。**任务卡给的「模块级单例 + Lock」对真正的
受益者直接不成立**；要让它生效，必须把 COM 调用收敛到**一个常驻线程**（队列化）——
那是比"加个单例"大得多的重构，且会把**长期存活**的 COM 代理压进那个管着
S1/D1 用户数据守卫的模块。用"给唯一受益者做一次队列化重构 + 在关键模块里长期持有
COM 代理"去换"工作台第二次导入从 25s 到 1.3s"，本轮不做。

**探针第一版测错了地方，一并记录**：我最初在另一个线程里直接读
`app.Presentations.Count`，得到的是 `CO_E_NOTINITIALIZED`（新线程根本没 apartment），
而不是线程亲和错误 —— 差点据此得出"单例可行"的反向结论。必须**先在新线程里
`CoInitialize`** 才能隔离出真正的亲和性问题。

### ④ 新增用例数：0（本提交只加探针，不改产品代码）

---

# 未修的

| 项 | 理由与数据 |
|---|---|
| **F3**（复用实例） | 见上：对 CLI 收益 0；对唯一受益者因 `RPC_E_WRONG_THREAD` 不可用。已探明可行路径见下 |
| **L6**（共享实例无互斥） | 不在本次 9 项内。`export_pages` 全程无锁，两次并发导出会在同一个实例上交错。守卫偏保守（少 Quit）故不破坏数据。F1 的改动没有让它变差 |
| **L5 / L7 / F4 / F5 / N1-N4** | 不在本次 9 项内，维持 `KNOWLEDGE.md` 风险表的状态 |

**F3 的替代路径（留给后续）**：在 `app.py` 内起一个**专用 COM 线程**（队列化）承载
pptx 导入，实例生命周期与那个线程绑定；`pptx_io.export_pages` 保持"一次调用一份实例"
的干净语义不动。这样既不把长期代理压进关键模块，也绕开了线程亲和问题。
代价：导入变成异步排队（工作台本来就已经是后台 worker + 轮询，所以代价可控）。

---

# 我不确定的

1. **F1 我无法按需复现"症状"**（测试 agent 实测它是 1/3 偶发）。我改用了
   **确定性不变量**做前后对照：`CoUninitialize` 的调用次数必须等于"我们是否自己
   初始化过"。这条是我**实测**的；而"修复前的症状在真机上一定发生"是**推断**
   （由 1/3 的频次推得）。
2. **`we_initialized` 在真机恒为 False**（因为 `import win32com.client` 已经初始化过）。
   也就是说修复后的生产路径其实从不调用 `CoUninitialize`。这是**实测**；
   对"这算不算轻微的资源泄漏"我的判断是**不算**（我们没加计数，就不该减；
   那一次初始化是 pywin32 导入时有意留的），但这是我的**推断**。
3. **M2 的三个阈值**（512 MB / 200:1 / 8 MiB）是按推理定的，**没有**用真实世界的
   恶意样本校准过。可能有合法的超大稿被误拦（例如整本杂志级 PPT）——若发生，
   错误码是 `PPTX_UNREADABLE` 且 message 里带具体体积，可据此调阈值。
4. **M1 的语义选择**：段尾的硬换行会**多出一个空行**（`"甲\n"` → 2 行）。这是按
   `text.split("\n")` 的契约字面语义选的（**实测**与 qa 一致），但 PowerPoint 是否
   真的为段尾的 `a:br` 渲染一个空行，我**没有**用 COM 底图验证过 —— 标为**推断**。
   影响有限：多算一行只会让行数偏保守。
5. **L1 的扩展名判据可以被骗**：把加密的 pptx 改名成 `.ppt`，我们会对它说
   "旧格式，请另存为 .pptx"（而不是"已加密"）。两者都不可执行，属**推断**的边界；
   严格做法是解析 CFB 目录，但那是另一个量级的实现。
6. **F2 的条件是否该放宽**：我按任务卡保留了"估宽 > 框宽"。若希望"任何
   `wrap=none` 都提示精度为段级"，改一行即可，但会让 b_multislide 平白多 13 条警告。

---

# 发现的新问题

1. **`_ZIP_MAGIC` 曾是小写 `pk\x03\x04`，导致 zip 分支永远不成立**（**实测**）。
   zip 本地文件头签名是 `PK\x03\x04`（大写），Python 的 `bytes.startswith` 大小写敏感。
   后果：任何"是合法 zip 但不是 pptx"的包都会落到"不是有效的 .pptx 包"那条分支，
   文案与实情不符（包结构其实没问题，只是内容不对）。**已在 M2 的提交里一并修正**
   （`b"PK\x03\x04"`），并加了回归用例 `test_zip_magic_is_case_sensitive`。
   这条是本次修复过程中**顺手撞出来的**，原审计与测试报告都没提。

2. **环境事故：uv 托管的 CPython 3.11.15 运行时被掏空**（**实测**）。
   会话中途 `.venv/Scripts/python.exe` 突然无法启动（`uv trampoline failed to spawn
   Python child process` / `os error 2`）。排查发现
   `%APPDATA%\uv\python\cpython-3.11.15-windows-x86_64-none\` 下只剩下
   `DLLs/` 与几个 DLL，**`Lib` 整块消失**（另一个 3.11 安装同样如此）。
   venv 自身的 `Lib/site-packages`（111 项、依赖齐全）完好。
   用 `uv python install 3.11.15 --reinstall` 修复（40.6s），恢复后 1066 全绿。
   **原因未查明** —— 不在本仓库范围内，也没有任何本仓库的代码/探针会碰那个目录
   （本次全部探针只写 `tempfile` 与 `output/`）。已如实记录，供排查。

3. **任务卡 F2 的验收条件与其修复条件自相矛盾**（**实测**）。见 §6：给的判据
   （`wrap=none` 且估宽 > 框宽）在指定素材 `b_multislide.pptx` 上**一处都不成立**，
   所以"对该稿跑 import 能看到降级提示"这条验收不可能通过。已按"修法与任务卡冲突
   → 写进报告说明，不要硬做"处理。

4. **两处「假测试」**：`test_encrypted_code`（任务卡已点名）与
   `test_pptx_independent.py` 的 `ole` 样本，都用**只有 OLE2 头**的合成文件断言
   `PPTX_ENCRYPTED` —— 等于把"任何 OLE2 都算加密"的误判锁成期望行为。均已改成
   带 `EncryptedPackage` 特征流的真样本。

5. **`import win32com.client` 会初始化当前线程的 apartment**（**实测**）。
   这条不只解释了 F1 的根因，也是一个通用陷阱：任何"导入即初始化"的模块都会
   让后续 `CoUninitialize` 的配对变得难以手工推理。

6. **任务卡 F1 字面建议的 API 拿不到所需信息**（**实测**）：`pythoncom.CoInitializeEx`
   恒返回 `None`，读不到 S_OK/S_FALSE。必须走 ctypes 问 ole32，或另找判据。
