# 独立测试报告 · pptx 双向链路（步 1–4）

> 角色：**独立验证 agent（ppt-test）**｜日期 2026-09-12｜基线 commit `b4f40cc`
> 被测：`pptx_io.py` / `hl_layout.py` / `hl_anim.py`（实现者步 1–4，自报 545 passed）
> 上级文档：`docs/TASK_TEST.md`（任务卡）、`docs/PPTX_INTERFACE.md`（契约）、`docs/IMPL_REPORT.md`（实现报告）
> 本次产出：`tools/probes/indep_*.py`（探针，一条命令一项）+ `tests/test_pptx_independent.py`（415 条用例）+ 本报告
> **未改任何实现代码、未改任何既有测试**（见 §8 变更清单）

---

## 0. 结论摘要

| # | 验证项 | 判定 | 一句话 |
|---|---|---|---|
| ① | M2 coverage 0.286 / 上限 0.414 独立复算 | **复现成功** | 用完全不同的像素法得 0.287 / 0.417 / pad0 0.323（实现者 0.286 / 0.414 / 0.321）；契约 0.35 确实不可达 |
| ② | D1 COM 实例隔离（真机） | **前提被证伪 + 加固有效** | 三个创建 API 拿到同一实例；`pre_count==0` 缺口被 `had_powerpoint` 守卫真堵住；`Visible` 红线未破 |
| ③ | M3 转义对抗（含真实 Chrome） | **通过** | 恶意 title/文本/路径全部惰性化，无 pageerror、无对话框、编码路径能取回真实文件 |
| ④ | 边界与异常路径 | **通过（附 1 条低危缺陷）** | 7 类病态输入均给中文 PptxError；`build_player` 对结构类型错误抛裸异常 |
| ⑤ | wrap_lines ↔ qa 等价性 | **通过** | 自选 18 类语料 × 24 组字号/行宽 = **864/864 一致**，另加 3 条独立不变量全绿 |
| ⑥ | M1 验收数字口径 | **一致 + 1 条新发现** | 裸 XML 独立核实 41 = 40 文本框 + 1 图表、过滤后 23；但"≤1.5s/页"只在附着/单次运行内成立 |

**新发现（实现报告未提）3 条**：F1 `export_pages` 拆掉调用方 COM apartment（中高）｜F2 真实素材上行级高亮退化为段级（中）｜F3 冷调用端到端 2.6s/页（中）。详见 §5。

**全量回归**：`.venv/Scripts/python.exe -m pytest tests/ -q` → **960 passed**（545 基线 + 415 新增）。
（输出里 `Windows fatal exception: code 0x80010108` 的 faulthandler 噪音来自实现者的 `tests/test_pptx_io.py`，非新增用例；见 §5 F1 附注。）

---

## 1. 方法与环境

- 环境：Windows 11 / Python 3.11.15 / `python-pptx 1.0.2` / PowerPoint 16.0 / Chrome（系统安装）
- 纪律：**每条断言都有命令输出支撑**；探针一律落盘成 `tools/probes/indep_*.py`，一条命令跑完一项（避免长内联命令被中断）。
- **COM 触碰的用户数据安全**：所有真机 COM 探针开头都有一道**闸门**——启动前若已有 `POWERPNT.EXE` 进程则**整条中止、一次 COM 都不调**。本次 12 次 COM 运行前闸门均为"无 PowerPoint 进程"；每次运行后均用 `tasklist` 复核无孤儿进程（本次自建的实例全部清理干净）。
- 本报告中：**"实测"= 有命令输出**；凡无输出支撑的一律标注 **【推断】**。

---

## 2. 验证 ① · M2 的 coverage 结论用独立像素法交叉验证

**测什么**：实现者报"行级 coverage 中位 0.286、任何矩形法的理论上限 0.414、pad 归零 0.321 → 契约阈值 0.35 不可达"。**不复用 `hl_layout.measure_coverage`**，自建两套独立度量：

- **法 A**：底色取 rect **外** 14px 远环的**众数色**（非契约的"内侧 1px 环中位"，也非实现者参照的"外侧 3px 环带中位"）；墨迹判据用 **RGB 欧氏距离 > 40**（非"任一通道差 > 30"）。
- **法 B**：**完全不做底色估计**——对 rect 区域做 Otsu 亮度双簇，取少数类占比。
- 另有 5 张 **ground-truth 对照图**（已知覆盖率）校验两法本身是否测准。

**命令**

```
PYTHONIOENCODING=utf-8 ./.venv/Scripts/python.exe tools/probes/indep_m2_coverage.py
```

**真实输出（关键行）**

```
== 0. 对照图：ground truth 校验 ==
  ctrl_30pct.png      期望 0.30   A=0.300  B=0.300  契约默认=0.300
  ctrl_solid.png      期望 0.00   A=0.000  B=0.000  契约默认=0.000
  ctrl_dark_1p0.png   期望 1.00   A=1.000  B=0.000  契约默认=0.000   ← 契约口径盲点
  ctrl_border_ink.png 期望 1.00   A=1.000  B=0.000  契约默认=0.000   ← rect==墨迹块
  ctrl_half_block.png 期望 0.50   A=0.500  B=0.500  契约默认=0.500

== 1. 真实素材（b_multislide.pptx，27 行 rect）==
  【A 独立法】coverage 中位 = 0.287      ← 实现者报 0.286
  【B Otsu】  coverage 中位 = 0.235
  【契约默认】coverage 中位 = 0.286
  【上限·墨迹 bbox 密度】中位 = 0.417    ← 实现者报 0.414
  【pad 归零敏感性】中位 = 0.323         ← 实现者报 0.321

== 2. 分解自洽性 ==
  竖向填充中位 = 0.731  横向填充中位 = 0.950  密度中位 = 0.417
  中位之积 = 0.289   逐行乘积的中位 = 0.287   差 = 0.002

== 3. 契约默认底色估计的偏差 ==
  契约默认 vs 独立口径：中位偏差 0.003，最大 0.008
  环上墨迹占比（ring_ink）：中位 0.000，最大 0.037 → ring_ink ≥ 0.5 的行数 = 0

== 7. 契约阈值判定 ==
  契约公式（pad=2.0/1.0pt）coverage 中位 0.287 ≥ 0.35 → FAIL
  契约公式把 pad 归零 → 0.323 ≥ 0.35 → FAIL
```

**判定：复现成功。** 差异 ≤0.004，远小于我设的 0.04 容差。

**由此得到的独立结论**

1. **D2（0.35 不可达）成立，不是口径 artifact。** 换底色估计法（远环众数 vs 内侧环中位）、换墨迹判据（欧氏 vs 单通道）、换统计方式（直接 rect 内统计 vs bbox 密度）都指向同一量级。连**完全不估底色**的 Otsu 法（0.235）也远低于 0.35。
2. **口径自洽性核查**：实现者称"0.731×0.960×0.414=0.286 完全自洽"。我实测中位之积 0.289 / 逐行乘积中位 0.287，差 0.002 —— 结论可用，但严格说这是**数值巧合级自洽**（中位数的乘积 ≠ 乘积的中位数），不是恒等式。
3. **契约默认口径在本素材上是安全的**（中位偏差 0.003），但**机制上有空洞**：rect 恰好等于墨迹块时，环被墨迹铺满 → 中位色倒向墨迹色 → 测出 0.0（对照图 `ctrl_border_ink`：真值 1.0，契约口径 0.0）。本素材 `ring_ink ≤ 0.037` 所以不发作；实现者报告称"12/22 行墨迹触框边"，离这个失效模式不远 —— 验收脚本必须显式传局部底色（实现者已这么做），此点已由我的测试锁住。
4. **"上限 0.414"的措辞**：该值是"墨迹 bbox 的密度"，而墨迹 bbox 必须**先有底图**才能算出（循环依赖），字体度量法拿不到。所以准确表述是"**基于字体度量的任何矩形法上限 0.414**"。这条只影响措辞，不影响"0.35 不可达"的结论。

**对应新增用例**：`test_independent_method_matches_ground_truth`、`test_contract_default_bg_is_blind_when_ink_fills_rect_border`、`test_fill_colored_pixels_inflate_coverage_with_wrong_bg`、`test_ring_bg_estimator_tilts_when_ink_covers_rect_edges`、`test_real_material_median_coverage_reproduced_independently`。

---

## 3. 验证 ② · D1（高危）COM 实例隔离（真机）

**测什么**：① 独立（换 API 组合）复现"`DispatchEx` 与 `Dispatch` 同一实例"；② D1 指出的残余缺口——**用户开着 PowerPoint 但没打开任何稿**（`pre_count == 0`）时，加固后的守卫能否保住用户的 PowerPoint；③ "绝不触碰 `app.Visible`"红线。

**命令**

```
PYTHONIOENCODING=utf-8 ./.venv/Scripts/python.exe tools/probes/indep_d1_com.py
```

**真实输出（关键行）**

```
[闸门] 启动前 POWERPNT.EXE PIDs = []（无）
A. Dispatch / DispatchEx / GetActiveObject 是否同一个实例
  Dispatch 后：PIDs = ['5340']（1 个进程）  Count = 0
  DispatchEx 后：PIDs = ['5340']（1 个进程）  Count = 0
  GetActiveObject 后：Count = 0
  app1 打开 1 份稿后：app1.Count=1  app2.Count=1  app3.Count=1
  打开后 PIDs = ['5340']（1 个进程）
  → 三个 API 是否同一个实例：是（DispatchEx 未开新进程/新实例）

B. D1 残余缺口：用户开着 PowerPoint 但没打开任何稿（pre_count==0）
  B 开始时 PIDs = ['5340']，Count = 0
  export_pages 完成：10 张，17.9s（width=640）
  导出后 PIDs = ['5340']   ← 用户的 PowerPoint 被保住
```

**判定：D1 的前提确实不成立（与实现者一致），且加固**有效**。**

- 我用**三个不同的创建 API** 交叉对照 + 进程数 + 一次"往 app1 加稿、看 app2/app3 是否同步可见"的独立观测，三重指向同一实例。
- **加固的必要性由反事实说明**：该场景下 `pre_count == 0`，若只按契约 §3.4-5 的守卫实现，`Quit()` 会执行 → 用户的 PowerPoint（开着但没稿）被关掉。实现者新增的 `had_powerpoint`（启动前 `tasklist` 探 POWERPNT.EXE）把这条路堵死。实测用户实例存活且随后仍可用（能再打开稿）。
- 该守卫在**附着场景**也成立：`indep_export_cost.py` 场景乙显示用户实例被保住（`PIDs = ['29444']`）。
- **`app.Visible` 红线**：静态检查 `pptx_io.py` 中"Visible"只出现在注释里，无赋值；真机运行前后读 `app1.Visible` 不变（见 `indep_d1_com.py` C 节）。

**附带发现（已升级为 F1）**：调用后**调用方自己的 COM 代理失效**（`RPC_E_DISCONNECTED`），且该线程后续 COM 调用报 `CO_E_NOTINITIALIZED`。见 §5 F1。

---

## 4. 验证 ③ · M3 转义对抗（自己构造恶意输入）

**测什么**：`title` = `</title><script>alert(1)</script>__CONFIG__<!--x-->`；单元文本含 `</script><script>`、`<!--`、`</SCRIPT>`、U+2028/U+2029、空字节、`"}; alert(4); var x={"`、`<img src=x onerror=alert(5)>`；底图路径含 `#`/`%`/空格/中文/引号/尖括号。**静态断言 + 真实 Chrome 打开产物**（挂 `pageerror` 与 `dialog` 监听）。

**命令**

```
PYTHONIOENCODING=utf-8 ./.venv/Scripts/python.exe tools/probes/indep_m3_escape.py
```

**真实输出（关键行）**

```
  <script 出现 3 次、</script> 1 次（模板各 1）
[browser] title='</title><script>alert(1)</script>__CONFIG__<!--x--> · 高亮讲解'
[browser] h1='</title><script>alert(1)</script>__CONFIG__<!--x--> · 高亮讲解'
[browser] state={'page': 0, 'step': 1, 'totalPages': 3, 'totalUnits': 11}  可见高亮=1  页点=3
[browser] 底图加载情况（naturalWidth, src）：
             64  bg/slide_%231%25x%20%E4%B8%AD%E6%96%87.png
              0  bg/a%22b%3Cscript%3Ec.png
              0  bg/%3Csvg%20onload%3Dalert%286%29%3E.png
  → 成功加载 1/3 张底图（含 #/%/空格/中文 那张）

== 用例 3：底图路径逃逸（含反斜杠/编码双写）==
  '..\\..\\x.png'              → 拦截 IR_MISMATCH
  'bg\\..\\..\\x.png'          → 拦截 IR_MISMATCH
  'C:foo.png'                  → 拦截 IR_MISMATCH
  '\\\\server\\share\\x.png'   → 拦截 IR_MISMATCH
  '//server/share/x.png'       → 拦截 IR_MISMATCH
  'bg/..%2f..%2fetc.png'       → 未拦截（放行）
判定：全部对抗性检查通过（静态 + 真实 Chrome）
```

**判定：通过。** 关键点：

- `document.title` 与 `h1.textContent` 都**等于原始恶意串**（带模板后缀）——这正好证明它被转义成字面文本渲染、没有被解析成标签。
- 全文 `</script` 恰好 1 次（模板自身），说明不存在脚本块提前闭合。
- **`<script` 出现在 script 体内的 JSON 字符串里是惰性的**（HTML 解析器只认 `</script` 终止）；真正的危险组合是 `<!--` + `<script` 把解析器拖进 double-escaped 态，而 `<!--` 被转义成 `<\!--` → 该状态不可达。**我的探针一开始把"体内出现 `<script`"误报为问题，核对机制后撤回**（已在探针里降级为 note）。
- 磁盘上真实存在 `bg/slide_#1%x 中文.png`，浏览器 `naturalWidth=64` → **百分号编码能取回真实文件**（这是端到端证明，不是断言编码字符串好看）。
- 路径逃逸 5/5 拦截；`bg/..%2f..%2fetc.png` 放行（见 F5，低危）。
- 单测版本：`tests/test_pptx_independent.py::test_adversarial_*`（含真实浏览器版本，无浏览器自动 skip）。

---

## 5. 验证 ④ 边界与异常路径 ／ ⑤ 断行等价性 ／ ⑥ M1 口径

### 验证 ④ · 边界与异常路径

```
PYTHONIOENCODING=utf-8 ./.venv/Scripts/python.exe tools/probes/indep_edges.py
```

```
== A. 输入文件病态 ==
  OK  不存在的路径            code=PPTX_NOT_FOUND  message='找不到文件：…'
  OK  目录当文件              code=PPTX_NOT_FOUND
  OK  txt 改名成 .pptx        code=PPTX_UNREADABLE message='…（不是有效的 .pptx 包）'
  OK  OLE2 头（伪装加密稿）    code=PPTX_ENCRYPTED  message='PPTX 已加密，无法读取'
  OK  截断的 zip              code=PPTX_UNREADABLE
  OK  合法 zip 但不是 pptx     code=PPTX_UNREADABLE
  OK  0 页 pptx               code=PPTX_EMPTY      message='PPTX 中没有任何幻灯片'

== B. 内容极端 ==
  读入形状 3 个：['HugeChar', 'LongText', 'NoWrap']；skipped = {empty_text:1, zero_size:1, hidden:1}
  OK  纯空白文本框被过滤（empty_text）／零宽形状（zero_size）／隐藏形状（hidden）
  OK  单字（200pt）宽于行宽仍只出 1 行   行数=1 rect宽=544px 形状宽=192px
  OK  1200 字段落行数与 qa 一致         行数=53（qa=53）
  OK  word_wrap=False 只出一行
  OK  有页但无形状 → 0 单元不崩          页数=1 units=0

== 表格 / 合并单元格 ==
  合并 origin 格：按单列宽(201.6pt)断行 = 2 行；按合并宽(417.6pt)应为 1 行
      → 证实 D5

== D. build_player 的畸形参数 ==
  OK  dim=nan / dim=True / auto_step_ms=0 / auto_step_ms='abc' / canvas_width_px=0 / canvas_height_px=-5
  ✗   bg_paths 非列表(None)  **裸 TypeError**：object of type 'NoneType' has no len()
  ✗   pages_units 非列表(None)  **裸 TypeError**
  ✗   units 里含 None  **裸 AttributeError**
```

**判定：通过（附低危缺陷 F4）。** 7 类病态输入全部落到正确的中文 `PptxError` 且带 `hint`，没有裸堆栈；极端内容（超宽单字、1200 字段落、零尺寸、隐藏、空页、无 xfrm 组合）都不崩且按契约过滤。

**D5 得到数值证实**：合并单元格的 origin 格按**单列宽**断行 → 本例切出 2 行而实际渲染 1 行 → 会多出高亮行；若该合并格是居中/右对齐，水平定位也会错（按起始列算）。

### 验证 ⑤ · `wrap_lines` ↔ `qa.measure_text_lines` 等价性

```
PYTHONIOENCODING=utf-8 ./.venv/Scripts/python.exe tools/probes/indep_wrap_equiv.py
```

自选 **24 类语料**（连续空格、全角标点行首、emoji、制表符、CRLF、纯数字长串、超长英文词+中文、U+3000、零宽字符、组合音标、长破折号、ASCII 标点串、URL、换行前后缀、只有空格、空串、日韩、全角数字标点…）× 6 字号 × 6 行宽：

```
  等价对：864/864
== 2. 独立不变量（不依赖 qa）==  累计问题 0 项
```

三条独立不变量（能抓"掉字/串行"类真 bug，qa 等价性抓不到）：
- **I1** 除空格外的字符一个都不丢（去掉空格后逐行拼接 == 原文）；
- **I2** 没有一行以空格开头（契约的行首空格丢弃）；
- **I3** 行宽超上限只允许出现在"单字符本身超宽"的强拆行。

**判定：通过。** 另锁住一条**有意分歧**：`word_wrap=False` 时 `_paragraph_lines` 只出 1 行而 qa（永远按换行算）给多行——这是契约 §4.3 规定的，不是漂移。**但这条分歧在真实素材上被大面积触发，见 F2。**

### 验证 ⑥ · M1 验收数字口径

**(a) 裸 XML 数形状**（不经过 python-pptx API）：

```
PYTHONIOENCODING=utf-8 ./.venv/Scripts/python.exe tools/probes/indep_m1_counts.py
```

```
  合计：顶层元素 41 个 → {'sp': 40, 'chart': 1}
  文本框 = 40（实现者报 40）一致 ｜ 图表 = 1（报 1）一致 ｜ 总数 = 41（报 41）一致
  read_pages：页数 10 一致；保留 23 → {'text': 22, 'chart': 1}；skipped 18 → {'empty_text': 18}
  对账：保留 23 + 过滤 18 = 41 vs 裸 XML 顶层 41 一致
判定：口径与实现者完全一致（独立核实通过）
```

**(b) COM 导出：页数/尺寸/耗时**

```
PYTHONIOENCODING=utf-8 ./.venv/Scripts/python.exe tools/probes/indep_m1_com.py
PYTHONIOENCODING=utf-8 ./.venv/Scripts/python.exe tools/probes/indep_m1_com_page.py
PYTHONIOENCODING=utf-8 ./.venv/Scripts/python.exe tools/probes/indep_export_cost.py
```

```
[cold] width=1920：10 张，27.96s（2.796s/页）尺寸集合={(1920, 1080)} 全绝对路径=True
[warm] width=1920：10 张，26.45s（2.645s/页）     ← 每次调用自己起/自己 Quit，"第二次"仍是冷启动
两次导出后 PIDs = []（无）  → 我们自己起的 PowerPoint 被 Quit：是
文件名 = slide_1.png … slide_10.png

== 成本拆分（同一实例内）==
  PowerPoint 启动（DispatchEx→可用） = 5.77s
  Presentations.Open                  = 0.18s
  逐页 Export：['0.343','0.119','0.119','0.129','0.127','0.170','0.102','0.137','0.123','0.098']
  单页中位 = 0.125s   ← 实现者报 0.146s

== 账目对照（indep_export_cost.py）==
  tasklist 子进程            = 0.47s
  冷调用端到端（自起实例）    = 26.42s → 2.64s/页
  附着调用端到端（用户实例）  =  2.26s → 0.23s/页
```

**判定：口径数字对得上，但"≤1.5s/页"的含义需要写清 —— 见 F3。**

---

## 6. 发现的问题（按严重度排序）

### F1（中高 · 新发现）`export_pages()` 会拆掉调用线程的 COM apartment，并让调用方持有的代理失效

**现象（实测）**：`export_pages()` 成功返回后，同一线程里

- 调用方在调用**之前**持有的 PowerPoint 代理 → `RPC_E_DISCONNECTED（0x80010108，对象没有连接到服务器）`；
- 该线程随后**任何** COM 调用都失败：`CO_E_NOTINITIALIZED（0x80010108 之外的 -2147221008，"尚未调用 CoInitialize"）`——**连 `Scripting.FileSystemObject` 都建不出来**，说明是整个 apartment 被卸掉，不是某个对象的问题；
- 重新 `pythoncom.CoInitialize()` 后一切恢复。

**归因（实测）**：逐步复刻 `export_pages` 的内部序列，在每步之后探 apartment：

```
  1) 我方 CoInitialize     → 存活    5) DispatchEx        → 存活    9) Slides(1).Export → 存活
  2) Dispatch 建实例        → 存活    6) Presentations.Count→ 存活   10) pres.Close()   → 存活
  3) subprocess.run(tasklist)→存活    7) Presentations.Open → 存活   11) 释放代理       → 存活
  4) export_pages 的 CoInitialize → 存活  8) PageSetup/Count → 存活  12) CoUninitialize → 已死(-2147221008)
```

**但**：单独做"一对 `CoInitialize`/`CoUninitialize`"（不碰 PowerPoint 业务）**不复现**（5 组对照 A–E 全部存活）。所以触发条件是**那一对 init/uninit 与 PowerPoint 打开/导出路径的交互**，机制未定 **【推断：条件概率性的服务端断开 + pywin32 引用计数不对称】**。

**稳定性**：**偶发**。3 线程频次探针 1/3 死亡；`indep_d1_impact.py` 里同线程连续两次 `export_pages` 都成功、两个调用方代理也都存活；`indep_export_cost.py` 的冷路径又复现了一次。

**影响面（实测 + 推断）**

- **CLI（步 6）**：`export_pages` 自身在入口 `CoInitialize`、出口 `CoUninitialize`，**自愈**——实测同线程连续两次调用都成功。所以单命令工具大概率不受影响。
- **Flask（步 8）/ 计划 R5 的 >50 页后台线程** **【推断】**：若同一 worker 线程在 `export_pages` 之前/之后还要用 COM（例如 `/api/export_pdf` 走 Word COM），会拿到 `CO_E_NOTINITIALIZED` 或死代理，报错信息与真实原因无关，很难排查。**我没有在 Flask 里实测这条路径**（见 §7）。
- 实现者在代码注释里把它记为"噪音，不用管"——**建议改写为"调用方需自行重新 CoInitialize"，并在 CLI 侧加一道防御**。修法（**我没有动手改**）：只在**本次调用真正完成了初始化**时才 `CoUninitialize`（例如判断 `CoInitializeEx` 的 S_OK/S_FALSE，或自己记一个 `we_initialized` 标志）。

**最小复现**

```
PYTHONIOENCODING=utf-8 ./.venv/Scripts/python.exe tools/probes/indep_com_apartment2.py
# 第 4 步会打印：旧 PPT 代理 = RPC_E_DISCONNECTED；FSO / 新 PPT = -2147221008
```

（同一脚本的"重新 CoInitialize 之后"一节显示两者立刻恢复。）

**附注**：`tests/test_pptx_io.py` 全量跑时 stderr 会出现 `Windows fatal exception: code 0x80010108` —— 与 F1 同源（faulthandler 抓到的 RPC 断开），实现者已在代码注释里记录。我的新增用例不产生该噪音（单跑 415 条无任何 fatal 输出）。

---

### F2（中 · 新发现）真实素材上"行级 tight 高亮"退化为段级 —— M2 的验收素材没有考查断行层

**命令 / 输出**

```
PYTHONIOENCODING=utf-8 ./.venv/Scripts/python.exe tools/probes/indep_wordwrap.py
  b_multislide.pptx：文本形状 word_wrap 取值分布 = {False: 13, True: 5, None: 4}
  多段文本框数量：3
  讲解单元 27 个 → 行 rect 共 27 条（多行单元 0 个）
  把 word_wrap 强置 True 后 → 行 rect 共 27 条
  qa 口径（无视 word_wrap）→ 共 26 行
```

**两件事**

1. 该素材上**每个段落都只出 1 行 rect**（27 单元 → 27 行，多行单元 0 个）。把 word_wrap 全部强置 True 后**仍然是 27**（说明没有段落真的需要折行）。因此 M2 报的 "coverage 中位 0.286（行级 tight）" 在这份素材上**等价于"段级 tight"**；`wrap_lines` 那层只被**单元测试**覆盖，没被验收素材碰过。
2. python-pptx 的 `shapes.add_textbox()` **默认写 `wrap="none"`**（即 `word_wrap=False`）——本素材 13/22 个文本框如此。凡经这类代码产出的稿，`hl_layout` 都会按"不换行"处理（契约 §4.3 规定如此），每段恒 1 行。

**影响**：M2 的结论**不能外推到会折行的真实稿**。折行一旦发生，行高×1.25 撑出的竖向填充、pad 占比、以及 coverage 的分解都会变，"0.35 不可达"的论证需要在新素材上复核。**建议**：要么补一份"长文本折行的真稿"重测 M2，要么在 M2 结论里显式声明"仅覆盖单行段落"。

**最小复现**：`indep_wordwrap.py`（另见 `indep_edges.py` B 节：1200 字段落在 `word_wrap=True` 下算出 53 行 == qa）。

---

### F3（中 · 新发现）冷调用端到端 ~2.6s/页，"≤1.5s/页"只在"附着/单次运行内"成立

**实测**

| 口径 | 数字 | 说明 |
|---|---|---|
| 单页稳态（同实例内逐页） | **0.125s/页**（中位），最大 0.343s | 达标，与实现者报的 0.146s 一致 |
| 附着调用（用户已开 PowerPoint，不启动不 Quit） | **0.23s/页**（10 页 2.26s） | 达标 |
| **冷调用（CLI 每次调用的真实形态）** | **2.64s/页**（10 页 26.42s） | **不达标** |
| 第二次连续冷调用 | 26.45s（2.645s/页） | **不摊销**：每次调用自己起 PowerPoint、自己 Quit |

**成本构成（实测）**：`DispatchEx` 启动 5.77s ＋ `Open` 0.18s ＋ 10 页 `Export` 1.47s ＋ `Close` 0.07s ＋ `Quit()` **立即返回**（但进程延迟退出，实测一次 >30s 未退出）＋ `CoUninitialize` 2.67s；`tasklist` 子进程 0.47s。冷/附着两场景差 24.16s = 启动＋Quit＋释放。

**影响**：**实现报告的"我不确定的地方 #7"被证实为真**——CLI 每次 `import` 会是 ~26s 的体感（10 页稿），且不会因为"第二次"变快。契约的"≤1.5s/页"若指的是**用户体感**，冷路径不达标；若指**稳态导出速度**，达标。**建议**：契约把口径写明（"稳态单页导出 ≤1.5s"），并在步 6 决定是否复用常驻实例（代价是与 D1 的实例隔离策略纠缠）。

**最小复现**：`indep_export_cost.py`（冷/附着一屏对照）。

---

### F4（低 · 新发现）`build_player` 对"结构类型错误"的参数抛裸异常而非 `PptxError`

```
  bg_paths=None      → TypeError: object of type 'NoneType' has no len()
  pages_units=None   → TypeError: object of type 'NoneType' has no len()
  units 含 None      → AttributeError: 'NoneType' object has no attribute 'kind'
```

契约 §6.1 的校验只点名了"页数不一致 → IR_MISMATCH"与"dim 越界 → BAD_ARGS"（两者实现正确），未覆盖**类型错误**。**缓解**：CLI 路径上 `load_deck` 已把结构校验前置（`except (KeyError, TypeError, ValueError) → IR_MISMATCH`），所以实际风险低。**最小复现**：`indep_edges.py` D 节。

---

### F5（低 · 新发现）`_check_bg_path` 不做百分号解码校验，`bg/..%2f..%2fetc.png` 放行

```
  'bg/..%2f..%2fetc.png'  → 未拦截（放行）
```

该路径在 `build_player` 里**只被引用、从不被本程序读取**（写进 `const BGS` 供浏览器加载），所以危害取决于下游：`file://` 下 Chrome 规范化掉；Flask 的静态路由一般会拒绝解码后的斜杠。**影响面为推断，未实测**。建议不修也可，但若要严谨，先 `unquote` 再跑同一套 normpath 检查。

---

### F6（低 · 信息）契约 §4.5 默认底色估计的失效条件已量化（本素材不发作）

- 真素材：契约默认 vs 独立口径 **中位偏差 0.003、最大 0.008**；`ring_ink` 最大 0.037（远未到 0.5 的失明阈值）。
- 构造极端（rect 上下边被文字铺满，环上 200/280 像素是墨迹）：默认口径测 **0.400**，真值 **0.600**——低估 20 个百分点。
- **结论**：本素材安全，但"验收脚本必须显式传局部底色"这条纪律是必要的（实现者已遵守，我用测试锁住）。

---

### 对任务卡 D1–D9 的独立复核结论

| 缺陷 | 我的独立结论 |
|---|---|
| D1 §3.4-2 前提不成立 | **确认**（三个 API 同实例、DispatchEx 不开新进程）。加固 `had_powerpoint` **必要且有效**。**新增 F1**（apartment 被拆） |
| D2 阈值 0.35 不可达 | **确认**（独立像素法 0.287 / 0.417 / pad0 0.323） |
| D3 `font_scale` 口径不自洽 | **未独立验证**（见 §7）—— 代码里确按 `/100` 取真实倍率，但我没有造出触发 `normAutofit` 的稿实测 |
| D4 标题判定第 ② 条不可实现 | **确认为真（读代码）**：`ShapeInfo` 无 ph type 字段。**影响未经实测**（本素材 `is_placeholder` 全 False） |
| D5 合并单元格 | **确认并量化**：合并 origin 格按单列宽断行 → 本例 2 行 vs 实际 1 行；居中/右对齐会错位 |
| D6 `DeckIR.pages` 与 `PageShapes` 缺桥 | **确认为真**：`hl_layout.page_shapes(page, deck)` 是必需补丁；缺 deck 时给中文报错 |
| D7 `Length` 行距重复计入 1.25 | **未独立验证**（见 §7） |
| D8 图表 coverage 不接近 1 | **确认**：独立法测得该 chart 单元 coverage **0.166**（全场最低） |
| D9 "40 形状"口径 | **确认**：裸 XML 41 = 40 `p:sp` + 1 chart，过滤后 23（22 text + 1 chart），18 条 `empty_text` |

---

## 7. 我无法验证的（环境限制 / 缺素材 / 超范围）

1. **真实手工 PPT（契约 V12）**：用户未提供。M2 的覆盖率结论仍只在**规范稿**（`builder.py` 直写 XML、从未被 PowerPoint 保存过）上成立；含图片/表格/SmartArt/被 PowerPoint 重排过的真稿未测。→ **F2 让这条限制更严重**（连"折行"都没覆盖到）。
2. **夹具稿的覆盖率**：`output/fixtures/fixtures_cover.pptx`（组合/表格/图片/图表）我只做了**几何级**检查（合并单元格断行），**没有**对它跑 coverage —— 那需要先 COM 导出底图（~26s/次）且夹具从未被 PowerPoint 保存过，代表性有限。
3. **F1 的真实事故场景**：Flask worker 线程里"先 COM 后 export_pages"（或反之）会不会真炸，我没有在 Flask 里实测（只证明了 apartment 会被拆、重新 init 可恢复）。
4. **并发**：3 线程同时驱动同一 PowerPoint 实例时，2/3 线程报 `Slide.Export : Object does not exist` / `Presentation.Close : Object does not exist` —— 说明**多线程并发访问同一 PowerPoint 实例不可靠**。这是 COM/Office 的一般特性，我只记录了现象，没有把它做成可复现结论（它也可能是我的探针线程编排问题）。
5. **D3 / D7 的实测**：需要专门构造带 `normAutofit` 与"绝对值行距"的稿并 COM 对照渲染，本次未做（只在建夹具时构造过，未量底图）。
6. **超大规模**：>50 页 / >100MB 的稿、zip bomb、恶意 XML —— 未测（属审计 agent 的范围）。
7. **真实字体替换（W10）**、旋转形状（V11）、`vertical_anchor is None` 的继承语义（V4）—— 未独立验证（实现者已用夹具做过，我未复核底图）。
8. **步 5–8**（`shot_player` / `cli.py` / `pptx_out.py` / 工作台）尚未实现，不在本次范围。

---

## 8. 变更清单（遵守硬约束）

**只新增，未修改任何既有文件**：

| 文件 | 内容 |
|---|---|
| `tools/probes/indep_00_env.py` | 环境勘察 + COM 闸门 |
| `tools/probes/indep_m2_coverage.py` | ① 独立像素法复算 coverage / 上限 / pad 敏感性 + 5 张对照图 |
| `tools/probes/indep_d1_com.py` | ② D1 真机：三 API 同实例、残余缺口守卫、Visible 红线 |
| `tools/probes/indep_d1_disconnect.py` | ②附：代理失效现象刻画 |
| `tools/probes/indep_com_apartment.py` | ②附：apartment 归因对照 T1/T2/T3 |
| `tools/probes/indep_com_apartment2.py` | ②附：**F1 最小复现**（插桩定位到哪一步） |
| `tools/probes/indep_apartment_bisect.py` | ②附：逐步复刻 export_pages，定位到第 12 步 |
| `tools/probes/indep_com_release.py` | ②附：A–E 五组对照，排除"单纯 init/uninit 对" |
| `tools/probes/indep_apartment_freq.py` | ②附：频次（1/3，偶发） |
| `tools/probes/indep_d1_impact.py` | ②附：影响面（连续两次调用 / 调用方代理） |
| `tools/probes/indep_m3_escape.py` | ③ 对抗性转义 + 真实 Chrome |
| `tools/probes/indep_edges.py` | ④ 病态输入 / 极端内容 / 打桩错误路径 / 畸形参数 |
| `tools/probes/indep_wordwrap.py` | ④附：**F2 最小复现**（真实稿 word_wrap 分布） |
| `tools/probes/indep_wrap_equiv.py` | ⑤ 864 组等价 + 3 条不变量 |
| `tools/probes/indep_m1_counts.py` | ⑥ 裸 XML 数形状 |
| `tools/probes/indep_m1_com.py` | ⑥ COM 导出页数/尺寸/耗时（冷+热） |
| `tools/probes/indep_m1_com_page.py` | ⑥附：逐页计时拆分 |
| `tools/probes/indep_com_quit_cost.py` | ⑥附：Close/Quit/CoUninitialize 成本 |
| `tools/probes/indep_export_cost.py` | ⑥附：**F3 最小复现**（冷 vs 附着） |
| `tests/test_pptx_independent.py` | 415 条独立用例（全绿） |

所有 `output/spike/indep_*/` 下的图与中间产物在 `.gitignore` 内，重跑探针即可再生。
