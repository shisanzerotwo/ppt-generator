# 任务卡 · 修复 agent（ppt-fix，复用 ppt-impl 会话）

> 本轮修复 `KNOWLEDGE.md` 风险表里遗留的 9 个问题（F1/F2/F3/M1-M4/L1/L4）。
> **基线：commit `e52504f`，`1028 tests collected`。**
> 每项修复都要有**「修复前复现 → 修复后消失」的对照证据**（照 `tools/probes/s1_probe_v2.py` 的范式：独立观察者 + 阳性/阴性对照）。

## 必读

1. `KNOWLEDGE.md` 的「已知风险 · pptx 双向链路」表 —— 每条问题的根因与证据出处都在这
2. `docs/TEST_REPORT.md` §6（F 系列实测细节）
3. `docs/AUDIT_REPORT.md` §2（S1/M1-M4/L1-L7 的分析）
4. `docs/PPTX_INTERFACE.md` **v2**（契约；注意 §12 已修正的 16 条前提）
5. `tools/probes/s1_probe_v2.py` —— **探针写法范式**（独立进程观察 + 安全三重闸门）

## 硬约束（违反即失败）

1. **S1 的守卫必须原样保留**：`we_opened`（只 Close 自己新增的那份稿）—— 这是防止关掉用户未保存稿的唯一防线
2. **D1 的守卫必须原样保留**：`had_powerpoint` + 进程数双守卫（不 Quit 用户的实例）
3. **绝不写 `app.Visible`**（会把用户正看的窗口藏起来）
4. `anim.py` / `builder.py` **一字不改**
5. 每项修复**独立 commit**；每步跑全量 `pytest tests/ -q` 保持**全绿**（当前 1028 条）
6. 涉及 COM / 浏览器的用例要能在缺环境的机器上 **skip**（照 `tests/test_shot_settle.py`）

---

## 修复清单

### F1（中高）· `export_pages` 拆掉调用方的 COM apartment

**根因**：`pptx_io.export_pages` 无条件 `pythoncom.CoInitialize()` + `finally: pythoncom.CoUninitialize()`。若调用方线程**已**初始化 apartment（如 Flask worker 先前做过 COM 操作），我们的 `CoUninitialize` 会连它的一起拆掉 → 调用方后续 COM 调用报 `0x800401F0 CO_E_NOTINITIALIZED`。
（编排者的 S1 v1 探针就是被它绊倒的 —— 见 `docs/AUDIT_REPORT.md` 的 S1 节。）

**修法**：改用 `pythoncom.CoInitializeEx(flags)` 并读 **HRESULT**：
- `S_OK(0)` = 我们自己新初始化 → finally 才 `CoUninitialize`
- `S_FALSE(1)` = 调用方已初始化 → **实测确认**该不该 `CoUninitialize`（引用计数语义可能与直觉不同，**以实测为准**）

**验收（必须真机实测）**：新增 `tools/probes/fix_f1_apartment.py`
- 主线程先 `CoInitializeEx` → 建一个与 PowerPoint 无关的 COM 代理（如 `Scripting.FileSystemObject`）
- 调 `export_pages`
- 再访问**旧代理** → 修复前应报 `CO_E_NOTINITIALIZED`，修复后应正常
- **阳性/阴性对照**：`git stash` 前后各跑一次，把两次输出都写进报告
- 安全闸门：若已有 `POWERPNT.EXE` 则整条中止（同 s1_probe_v2）

---

### M2（中）· zip bomb 无闸门

**修法**：`read_pages` 前置 `_check_package_safety(path)`：`zipfile.ZipFile(path).infolist()` 累计 `file_size`，设**解压总量上限**与**单条压缩比上限**；超限抛 `PptxError(..., "PPTX_UNSAFE", 中文指引)`。
**阈值建议**：解压总量 ≤ 512 MB、单条压缩比 ≤ 200:1（AUDIT_REPORT M2 实测样例是 **601:1 / 64 MiB**）。请在报告里写出取值依据。
**验收**：照 `tools/probes/audit_pptx_io.py` 构造样例 → 断言被拒且错误码正确；**正常稿不受影响**（1028 全绿）。

---

### M4（中）· 缺 `p:sldSz` → 裸 `TypeError`，错误码从 3 退化成 5

**修法**：`prs.slide_width` / `slide_height` 为 `None` 时抛 `PptxError("...", "PPTX_UNREADABLE", 中文指引)`，不要裸 `int(None)`。
**验收**：构造缺 `p:sldSz` 的包 → 断言错误码是 `PPTX_UNREADABLE`（3）而**不是** `INTERNAL`（5）。

---

### M3（中）· `inner_w_pt <= 0` 的形状被**静默丢弃**

**修法**：`hl_layout.build_units` 增加可选收集器（如 `skipped: list | None = None`），记录被丢弃的形状标识 + 原因；`cli import` 把它并入 JSON 输出的 `warnings`。
**验收**：断言 `skipped` 有内容，且 `pptgen import` 的 JSON 里能看到。

---

### M1（中）· `a:br` 段落几何与 `qa` **分叉**

**根因**：`qa._tokenize` 把 `"\n"` 当 1.0em 的 CJK 字形（不是硬换行），而 `hl_layout._paragraph_lines` 先硬拆 → 两者差 **+1 行**；且契约 §4.2 那条"防漂移"等价断言只盖 `wrap_lines`，**盖不住真正产出几何的 `_paragraph_lines`**。
**修法**（二选一并说明理由）：
- (a) 抽公共函数，让两处用**同一套** `a:br` 规则
- (b) 明确声明两者语义本就不同（"渲染行数" vs "溢出行数"），并把等价断言**指向正确对象**
**验收**：新增用例覆盖 `a:br`；若选 (a) 则断言两者对含 `\n` 文本给出相同行数。

---

### F2（中）· 真实素材上图级高亮**退化为段级**

**根因**：`wrap="none"` 的文本框（python-pptx `add_textbox` 默认）读不到实际换行 → 按单行算，真实换行被忽略（该稿 **13/22** 个框如此）。
**修法**：检测"`wrap="none"` 且文本估算宽度 > 框宽" → 记 **warning 声明精度降级为段级**。**不要**硬猜 PowerPoint 的换行规则（不可靠）。
**验收**：对 `output/b_multislide.pptx` 跑 `pptgen import`，warnings 里能看到降级提示。

---

### L4（低）· 播放器不校验底图存在 → 黑屏静默通过

**修法**：`hl_anim` 给 `<img>` 加 `onerror` → 让 `hl.ready` **reject**（或置错误标志），避免黑帧静默进 MP4。
**验收**：构造缺图场景 → 断言不会静默 resolve。

---

### L1（低）· 老 `.ppt`（OLE2）被误判成「PPTX 已加密」

**修法**：`_classify_bad_package` 区分 OLE2 复合文档（老 `.ppt`/`.doc`）与真正的加密 OOXML → 给对的提示（"这是旧版 .ppt 格式，请先另存为 .pptx"）。
**注意**：`tests/test_pptx_io.py::test_encrypted_code` 用合成 OLE2 头**把误判锁成了期望行为** → 该用例需要一并修正（这是"实现者的假测试"实例）。

---

### F3（中）· 冷调用 2.6s/页 → **复用 PowerPoint 实例**

**修法**：模块级单例 app + `threading.Lock`：
- 首次 `DispatchEx` 后**缓存** app；后续调用先探活（`Presentations.Count`）再复用
- **复用路径不 Quit** —— 省时间就省在这里
- 提供显式释放（`release_app()` + `atexit`），CLI 支持 `--no-reuse` 回退
- **必须保持** S1（`we_opened`）+ D1（`had_powerpoint`）守卫语义：**不复用时行为与现在完全一致**
- 加锁保证并发安全（Flask 多线程 + CLI）
**验收（必须实测对比）**：新增 `tools/probes/fix_f3_reuse.py` —— 同一进程内连续两次 `export_pages`（各 10 页），给出**实测秒数**；修复后第二次应显著快于第一次。
**风险**：会留一个**常驻 `POWERPNT.EXE`**。报告里必须写清：常驻何时释放、如何确认不留孤儿进程（`tasklist` 复核）。
**允许放弃**：若实测收益不明显或引入数据风险（S1/D1），**可以放弃并给出数据与理由** —— 用性能换数据安全不划算时就不换。

---

## 报告

写 `docs/FIX_REPORT.md`，每项固定四段：
① **根因** ② **修法**（含被否决的替代方案与理由） ③ **修复前/后的对照实测输出**（贴真实命令输出） ④ **新增用例数**

末尾固定小节：
- **未修的**（含理由与数据）
- **我不确定的**（明确标注推断 vs 实测）
- **发现的新问题**（修复过程中撞到的）

## 纪律

- **不许只改代码不验证**：每项都要有可复现命令 + 真实输出
- 修法与任务卡冲突（发现不可行）→ **写进报告说明**，不要硬做
- 真机探针一律带**安全三重闸门**（有 `POWERPNT.EXE` 即中止 / 只用 tempfile 副本 / 只清理自己启动的）
- 分批小步做（每项一个 commit），避免长命令被外部按键打断
