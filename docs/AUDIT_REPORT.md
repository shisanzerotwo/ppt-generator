# 审计报告 · pptx 双向链路 步 1–4（`pptx_io.py` / `hl_layout.py` / `hl_anim.py`）

> 角色：审计 agent（ppt-audit）｜日期 2026-09-12｜基线 commit `b4f40cc`（工作区含未提交改动）
> 上游：`docs/TASK_AUDIT.md`（任务卡）、`docs/PPTX_INTERFACE.md`（契约）、`docs/IMPL_REPORT.md`（实现报告 D1–D9）、
> `PLAN_PPTX_ANIM.md`（R1–R9）、`KNOWLEDGE.md`（已知风险红线）、`docs/TEST_REPORT.md`（前一轮独立测试 F1–F6）
> **本次未修改任何被测代码**（`pptx_io.py` / `hl_layout.py` / `hl_anim.py` 一字未动，`git status` 不含这三个文件的新增改动）。
> 证据分级：**【已实测复现】= 有本次执行的命令与输出**；**【静态推断】= 只读代码/文档得出的结论，未执行**。

---

## 0. 结论摘要

| # | 严重度 | 一句话 | 证据级别 |
|---|---|---|---|
| **S1** | **严重（条件性）** | `export_pages` 的 `Quit()` 有两道守卫，`pres.Close()` **一道也没有**。既然 D1 已证实"我们附着在用户的 PowerPoint 实例上"，当**用户自己开着同一份稿**时，被关掉的可能是用户的稿（有未保存内容即数据损失）。安全的另一半（Quit）守住了，危险的一半（Close）裸露。 | 静态推断（按任务卡要求：需启动 PowerPoint 才能实测，**未执行**） |
| **M1** | 中 | 几何用的 `_paragraph_lines` 在 `a:br`（段内软换行）上与溢出 QA 的模型**分叉**（同一段差 +1 行），而契约 §4.2 强制的那条"防漂移"等价断言只覆盖 `wrap_lines`，**盖不住真正被调用的那个函数**。且这是报告未申报的 §4.3 偏离。 | 已实测复现 |
| **M2** | 中 | **zip bomb 无任何大小/比率闸门**：把 64 MiB 全零塞进 `[Content_Types].xml`（包 109 KB），`Presentation()` 照单全收、67,108,864 字节在内存里展开。契约 §8.3 点名要审这一项。 | 已实测复现 |
| **M3** | 中 | `inner_w_pt <= 0` 的形状被**静默丢弃**：契约 §4.3 明写"记 warning"，但实现里连一条诊断都没有，`build_units` 也没有 `skipped` 出口——"高亮为什么少了这一块"无从得知。 | 已实测复现 |
| **M4** | 中 | 缺 `p:sldSz` 的包 → `prs.slide_width is None` → `int(None)` 抛裸 `TypeError`。该行在 `read_pages` 的 `try` **之外**，于是错误码从 `PPTX_UNREADABLE(3)` 退化成 `INTERNAL(5)`。 | 已实测复现 |
| **L1** | 低 | 老 `.ppt`（同为 OLE2 复合文档）被 `_classify_bad_package` 一律判为**「PPTX 已加密」**，给的指引（"去掉打开密码"）不可执行；`tests/test_pptx_io.py::test_encrypted_code` 用合成 OLE2 头把这个误判**锁定成了期望行为**。 | 已实测复现 + 静态推断 |
| **L2** | 低 | 表格 `table_col_widths_emu` / `table_row_heights_emu` 短于单元格矩阵时**静默截断**（实测 6 单元→2 单元），无异常、无 warning。 | 已实测复现 |
| **L3** | 低 | 被改过的 `deck.json`（几何为 `null`、画布 `0`）能让 `build_units` 裸崩 `TypeError` / `ZeroDivisionError`，而 `load_deck` 不做几何校验 → CLI 只能归 `INTERNAL`，不是 `IR_MISMATCH`。 | 已实测复现 |
| **L4** | 低 | 播放器**不校验底图是否存在/是否新鲜**：缺图时页面全黑，`window.hl.ready` 仍会 resolve —— 步 5 的 `shot_player` 会把它安静地截进 MP4。 | 静态推断 |
| **L5** | 低 | `had_powerpoint` 采样与 `DispatchEx` 之间存在 TOCTOU 窗口：用户恰好在这个窗口里**启动** PowerPoint（还没开稿）→ 结束时 `pre_count==0` 且 `had_powerpoint==False` → 用户的 PowerPoint 被 `Quit`。D1 的加固堵住了"事先开着"，没堵住"这期间打开"。 | 静态推断 |
| **L6** | 低 | `export_pages` 对**共享的** PowerPoint 实例没有任何互斥：两次并发导出（Flask 工作台 + CLI，或 >50 页后台路径）会交错。守卫偏保守（少 Quit）故不至于破坏数据，但导出可能失败。 | 静态推断 |
| **L7** | 低 | `measure_coverage` 用 `0.0` 重载了三种语义（矩形退化 / 完全在图外 / 真的没有墨迹），验收脚本无法区分"高亮算错"与"这块本来就没墨迹"。 | 已实测复现 |
| **N1–N4** | 提示 | 契约 §2.3 自相矛盾（"已展开组合" vs `children` 字段）；`CFG.title` / 单元 `t` 字段从未被 JS 使用；测试数字三处不一致且素材被 gitignore；`app.py` 的 `_export_pdf_via_com` 仍是"Dispatch + 无条件 Close/Quit"的老写法。 | 见正文 |

**对前一轮 TEST_REPORT F1–F6 的复核**：F1（拆掉调用方 COM apartment）我独立确认代码根因成立；F5（`..%2f` 放行）**予以修正/降级**——`_web_path` 会把 `%` 再编码成 `%25`，越界不可达；F2（真实素材上退化为段级）我确认其结论，并补上一条它没覆盖的机制（见 M1）。

---

## 1. 范围与方法

**审什么**：`pptx_io.py`（含 `export_pages` 的 COM 全路径）、`hl_layout.py`（断行层 + 定位层 + 覆盖率度量）、`hl_anim.py`（播放器与转义）；`tests/test_pptx_io.py`、`tests/test_hl_layout.py`、`tests/test_hl_anim.py` 的测试质量；D1–D9 的契约一致性。

**方法与纪律**

- 每条断言必须有命令输出；只读代码得出的结论一律标【静态推断】。
- 需要**启动 PowerPoint** 才能验证的（真机 COM、D1 隔离、`Quit` 守卫实测）**一律跳过**并在 §5 声明——前一轮 `docs/TEST_REPORT.md` §3 已用真机做过，本次不重复冒险。
- 路径/注入验证一律**在内存里做**：`build_player` 会真的写 `index.html`，故本次用 `builtins.open` + `os.makedirs` 拦截，把产物留在 `io.StringIO` 里（跑的是真代码路径，磁盘零产物）。注入判定另起真 Chromium（`page.set_content()`，不落文件）。
- 本次新增的复现脚本（**这是本次唯一新增的文件，除本报告外**）：

| 脚本 | 覆盖 |
|---|---|
| `tools/probes/audit_inject.py` | title / 单元文本 / 底图路径的注入与穿越反例（内存内） |
| `tools/probes/audit_xss_browser.py` | 真 Chromium 解析 12 例恶意 payload，判据 `window.__pwned` + `<script>` 元素数 + dialog（含阳性对照） |
| `tools/probes/audit_layout.py` | 静默丢弃 / 表格截断 / `a:br` 分叉 / 覆盖率盲点 / 被改 deck.json 的裸崩 |
| `tools/probes/audit_pptx_io.py` | zip bomb / 实体展开 / 联网面 / OLE2 误判 / `sldSz` 缺失 / 真实稿 a:br 分布 |
| `tools/probes/audit_paths.py` | `_check_bg_path` 的百分号解码链（F5 复核） |

---

## 2. 发现

### S1（严重·条件性）· `pres.Close()` 无守卫：`Quit` 守住了，`Close` 裸露

**位置**：`pptx_io.py::export_pages`，`pptx_io.py:645-656`（`finally`）。

```python
finally:
    try:
        if pres is not None:
            pres.Close()                       # ← 无条件，无任何守卫
    except Exception:
        pass
    try:
        # 两道守卫都通过才 Quit
        if app is not None and not had_powerpoint and app.Presentations.Count == 0:
            app.Quit()
    except Exception:
        pass
```

**为什么这是本次最危险的路径**：模块自己的头注释（`pptx_io.py:14-21`）已经承认——`DispatchEx` 在本机**附着到用户实例**上（D1）。既然 `app` 是用户的 PowerPoint，那么 `pres` 是否"我们的稿"就完全取决于 `Presentations.Open(abs_pptx, …)` 的语义：

- 该路径的威胁模型在契约里写得很清楚（§3.4-2）：「`Dispatch` 会附着到用户正在用的 PowerPoint 实例，随后的 `Quit()` 会关掉用户没保存的稿子」。
- 实现为 `Quit()` 补了两道守卫，但**同一个威胁在 `Close()` 上一字未设防**。
- 触发场景：**用户自己正开着这份 `deck.pptx`**（做演示时尤其常见），此时 `pre_count ≥ 1` → `Quit` 被第一道守卫挡住（正确），但 `present.Close()` 在 `finally` 里照跑。若 PowerPoint 的 `Open` 对"已打开的文件"返回**已存在的那份 Presentation**，`Close()` 关掉的就是用户的窗口；有未保存修改时 PowerPoint 通常弹保存对话框 → 自动化里表现为挂起或静默丢弃。

**为什么我标"条件性"而不是无条件严重**：风险成立与否取决于 PowerPoint 对"重复 Open 已打开文件"的真实语义（返回已有实例 / 报错 / 只读再开一份）。这一条**必须启动 PowerPoint 才能定论**，按任务卡纪律本次跳过。

**建议**
1. 让 `Close` 与 `Quit` 对称：`Open` 之前记下 `app.Presentations` 的名字集合，`finally` 里**只关自己新开的那一份**（用 `pres.Name` 比对；不在差集里就不关）。
2. 把一个可复现实验补进 `tools/probes/`：**先手工开一份未保存的稿并保持打开** → 跑 `export_pages(同一份文件)` → 确认该稿仍在（这正是契约 V8 没覆盖的那半）。

**证据**（静态；`Quit`/`Close` 的全部调用点）

```
$ grep -n "\.Quit()\|\.Close()\|pre_count\|had_powerpoint\|Presentations\.Count" pptx_io.py
599:    had_powerpoint = _powerpoint_running()
613:        pre_count = app.Presentations.Count
648:                pres.Close()
653:            if app is not None and not had_powerpoint and app.Presentations.Count == 0:
654:                app.Quit()
```

`pre_count` **被赋值后从未被读取**（第 613 行赋值，全文件再无引用）——守卫实际用的是 `had_powerpoint` + 结束时的 `Count == 0`。契约 §3.4-5 的伪码写的是 `pre_count == 0`，实现的第二道守卫换成了"结束时 Count==0"（更严），这是**未申报的偏离（属加固方向，非缺陷）**，但同时说明 `pre_count` 成了死变量（提示级）。

---

### M1（中）· `a:br` 段落：几何用的 `_paragraph_lines` 与溢出 QA 分叉，而"防漂移"闸门盖不住它

**位置**：`hl_layout.py::_paragraph_lines`（`hl_layout.py:170-182`）与 `hl_layout.py::_build_units._text`。

**问题**：契约 §2.3 行 4 规定 `a:br → "\n"` 存进 `ParaInfo.text`；§4.3 又规定"`word_wrap is False` → `text.split("\n")`；**否则 `wrap_lines(text, size_pt, inner_w_pt)`**"。但 `qa._tokenize` 把 `"\n"` 当作**一个 1.0em 宽的 CJK 字形**（既不是空格、也不是 alnum → 落进 `else` 分支），它**不构成硬换行**。于是：

- 按契约字面照做 → 软换行被吃掉（渲染错）；
- 实现的 `_paragraph_lines` 额外加了一条 `if "\n" in text: 先硬拆再各自贪心` → 渲染对，但**与 `qa` 对同一段算出不同行数**。

**契约 §4.2 的强制性等价断言**（`assert len(wrap_lines(t,s,w)) == qa.measure_text_lines(t,s,w)`）测的是 `wrap_lines`——它和 `qa` 一致（我独立复验过，见 §3）；**真正产出几何的是 `_paragraph_lines`，而它没有任何等价测试**。契约"两套算法一旦漂移，溢出告警和高亮位置就会互相打架"的防线，在 a:br 上失效。

**证据 ·【已实测复现】**

```
$ PYTHONIOENCODING=utf-8 ./.venv/Scripts/python.exe tools/probes/audit_layout.py
（§3 段）
  text='第一行\n第二行'          size=18.0 行宽=300.0pt → _paragraph_lines=2 wrap_lines=1 qa=1（与 QA 差 +1）
  text='A\nB'                  size=18.0 行宽=300.0pt → _paragraph_lines=2 wrap_lines=1 qa=1（与 QA 差 +1）
  text='前缀字×20\n后缀'          size=18.0 行宽=300.0pt → _paragraph_lines=3 wrap_lines=2 qa=2（与 QA 差 +1）
```

**现实性边界（必须说清，避免夸大）**：本次在**仓库内三份稿**里统计 a:br 均为 **0**（`b_multislide.pptx` 44 段 / `fixtures_cover.pptx` 25 段 / 某真实 LLM 稿 51 段，全部 `a:br=0 a:fld=0`，见 `audit_pptx_io.py` §7）。所以这条**在当前素材上不发作**；但手工制作的真实 PPT 里软换行很常见（`pptx_io._para_text` 的注释自己也这么写："软换行在汇报稿里很常见"），一旦接真实稿就会踩到。

**定性**：这是**契约自身的内部矛盾**（§2.3 要求 `a:br→"\n"` 与 §4.3 要求把整段交给 `wrap_lines` 不可兼得）。实现选择了"渲染正确"的一侧，但**未在 IMPL_REPORT 申报**（D1–D9 里没有这一条）→ 按任务卡口径"未申报的偏离算缺陷"。

**建议**（三选一，需架构 agent 定）
1. 修 `qa._tokenize`：让 `"\n"` 走硬换行（行数 +1、宽度清零），然后 `_paragraph_lines` 可直接退回契约字面写法，两边回到同源。代价是要评估对既有 248 基线里溢出告警的影响。
2. 保持实现、把契约 §4.3 改成"先按 `\n` 硬拆"，并把等价测试改成覆盖 `_paragraph_lines`。
3. 若 `\n` 由 `a:br` 产生（而非本段结尾），在 IR 里标出"硬换行"语义，两边各自处理。

---

### M2（中）· zip bomb 无闸门（契约 §8.3 点名项）

**位置**：`pptx_io.py::read_pages`（`Presentation(abs_path)`）—— 解析链路全在 python-pptx 里，本模块不做任何包体校验。

**证据 ·【已实测复现】**（内存构造，未落盘）

```
$ PYTHONIOENCODING=utf-8 ./.venv/Scripts/python.exe tools/probes/audit_pptx_io.py
（§2 段）
  构造出的包体积 = 111,679 B（109.1 KB），内含一个 64 MiB 的 [Content_Types].xml → 压缩比 ≈ 601:1
  Presentation() 结果：XMLSyntaxError：Document is empty, line 1, column 1，耗时 0.55s
  但解压轨迹 = [('[Content_Types].xml', 67108864), ('_rels/.rels', 737), ...]
  → 即使 XML 立刻解析失败，字节已在内存里展开；pptx 侧只做了 read，没有任何大小/比率校验
```

即：**一个 109 KB 的文件已让 64 MiB 进内存**。pptx 侧没有 `zipfile` 的 `file_size`/`compress_size` 阈值检查，也没有解压总量上限；把 payload 换成合法但巨大的 XML（或 42.zip 级嵌套）同样照收。触发点在 `[Content_Types].xml`（`Presentation()` 打开时**必读**，`audit_pptx_io.py` §1 的 59 条读入轨迹可证）。

**可利用性边界**：当前入口是 CLI/本地文件（步 8 的上传入口尚未落地），**还不是远程可达**；一旦 `app.py` 的 pptx 上传路由上线，就变成"上传一个 100 KB 文件把服务打挂"。契约 §8.3 已把这项列为审核项。

**建议**：`read_pages` 前置一道廉价门禁——遍历 `ZipFile.infolist()` 累计 `file_size`（这是 zip 目录里的声明值，**不需要解压**），超过阈值（如 200 MB）或单个条目 `compress_size/file_size` 比率异常（如 > 200:1）即 `PptxError("PPTX_UNREADABLE", "包体异常：解压后体积 xxx MB")`。零新依赖、零解压成本。

---

### M3（中）· `inner_w_pt <= 0` 静默丢弃形状，契约要求的 warning 无处承载

**位置**：`hl_layout.py::_UnitBuilder._text`（`hl_layout.py:280-281`）。

```python
if inner_w_pt <= 0 or inner_h_pt <= 0:
    return                                  # ← 契约 §4.3 要求"记 warning"，此处一条不留
```

契约 §4.3 原文：「`inner_w_pt <= 0` → 跳过该形状（**记 warning**，不产出 Unit）」。实现只做到了"不产出 Unit"。

**证据 ·【已实测复现】**

```
$ PYTHONIOENCODING=utf-8 ./.venv/Scripts/python.exe tools/probes/audit_layout.py
（§1 段）
形状宽 100000 EMU，左右内边距各 91440（合计 182880）→ inner_w_pt = -6.526pt（<=0）
build_units 返回单元数 = 0（预期 0 → 形状被跳过）
→ 有没有任何出口告诉调用方『这个形状为什么没有高亮』：无（契约 §4.3 要求的 warning 无处可寻）
```

**为什么这是"最危险的一类"静默失败**：`build_units` 返回 `list[Unit]`，形状被丢弃后**没有任何通道**告诉调用方"这里少了一块"。CLI 的 JSON 信封（契约 §8.1）里有 `skipped` 字段——`read_pages` 会把导入期过滤写进 `deck.skipped`，但**布局期丢弃一条都进不去**。用户在播放器里看到的只是"某些文字没有高亮"，排查时没有任何线索。

**建议**：给 `build_units` 加第二个返回值或可选收集器（如 `build_units(page, ..., skipped=None)` 追加 `{shape_id, shape_name, reason}`），并在 CLI 的 `skipped`/`warnings` 里冒泡。这与 `read_pages` 已有的 `skipped` 设计完全同构，成本极低。

---

### M4（中）· `prs.slide_width is None` → 裸 `TypeError`，错误码从 3 退化成 5

**位置**：`pptx_io.py:470-471`。

```python
    slides = list(prs.slides)          # 上面唯一被 try 包住的是 Presentation(abs_path)
    ...
    width_emu = int(prs.slide_width)   # ← 在 try 之外；slide_width 可以是 None
```

**证据 ·【已实测复现】**（内存构造：从 `presentation.xml` 移除 `<p:sldSz>`）

```
$ PYTHONIOENCODING=utf-8 ./.venv/Scripts/python.exe tools/probes/audit_pptx_io.py
（§6 段）
  已从 presentation.xml 移除 <p:sldSz>（54 字符）
  prs.slide_width = None
  int(slide_width) -> TypeError: int() argument must be a string, a bytes-like object or a real number, not 'NoneType'
  → read_pages 的 try 只包住 Presentation()，这行在 try 之外 → 崩溃会冒成 INTERNAL(5)，而不是 PPTX_UNREADABLE(3)
```

**影响**：契约 §9 的文案纪律是"message 说发生了什么、hint 说下一步做什么"，`INTERNAL(5)` 的语义是"这属于 bug，请附命令与 deck.json"。对一个**用户稿子畸形**的情形给出"这是我们的 bug"，会把排查引向错误方向。同类风险还有 `prs.slides`（`slide_width`/`slide_height` 均为 `None` 的包不罕见——手工拼的包、被工具改坏的包）。

**建议**：把 `width_emu/height_emu` 的读取并入现有的 `try`，或在读完后判空并抛 `PPTX_UNREADABLE("PPTX 缺少幻灯片尺寸定义（p:sldSz）")`。

---

### L1（低）· 老 `.ppt` 被判成"已加密"，且测试把该行为锁定

**位置**：`pptx_io.py::_classify_bad_package`（`pptx_io.py:439-445`）。

```python
if head.startswith(_OLE_MAGIC):
    return PptxError("PPTX 已加密，无法读取", "PPTX_ENCRYPTED",
                     "请先用 PowerPoint 去掉打开密码再导出")
```

**问题**：OLE2 复合文档这个魔数**不专属于"加密的 pptx"**——老 `.ppt`／`.doc`／`.xls` 全是 OLE2。用户拿一份 `.ppt` 过来会被告知"已加密，请去掉打开密码"，而 `hint` 完全不可执行（那份文件根本没有密码）。契约 §9 给 `PPTX_UNREADABLE` 写的示例文案恰好是"确认是 `.pptx`（非 `.ppt`/`.pdf`）"——正是这条魔数分支把这个正确提示抢走了。

**证据 ·【已实测复现】**（拿仓库内合成 OLE2 样本喂分类器）

```
$ PYTHONIOENCODING=utf-8 ./.venv/Scripts/python.exe tools/probes/audit_pptx_io.py
（§5 段）
  仓库内找到的真实 OLE2 样本：['...\output\spike\indep_edges\encrypted.pptx']
  判定结果：code='PPTX_ENCRYPTED' message='PPTX 已加密，无法读取' hint='请先用 PowerPoint 去掉打开密码再导出'
```

对这份**合成**样本判定是对的（契约意图）。但对真 `.ppt` 的误判属**【静态推断】**：本机没有真实 `.ppt` 样本，我按 OLE2 格式事实（`.ppt` = CFB）推断同一分支会命中。要确证只需一份任意 `.ppt`（1 分钟）。

**测试质量问题（更值得注意）**：`tests/test_pptx_io.py::test_encrypted_code` 自己拼了 `b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00"*512` 并断言 `code == "PPTX_ENCRYPTED"`——**它把"该分支对任何 OLE2 都这么答"固化成期望行为**，将来真 `.ppt` 误判也不会被测试抓到。

**建议**：区分手段用**扩展名 + CFB 目录里是否存在 `EncryptedPackage` 流**（加密 OOXML 的特征）或至少用 `"EncryptedPackage" in head_bytes`；改不动就退化为 `PPTX_UNREADABLE` 并把 hint 写成"若这是旧的 .ppt，请先在 PowerPoint 里另存为 .pptx"。

---

### L2（低）· 表格列宽/行高数组短于单元格 → 静默截断

**位置**：`hl_layout.py::_UnitBuilder._table`（`hl_layout.py:372-377`）。

```python
for r, row in enumerate(shape.table_cells):
    if r + 1 >= len(row_y):
        break                       # 剩余行静默丢弃
    for c, cell in enumerate(row):
        if c + 1 >= len(col_x):
            break                   # 该行剩余列静默丢弃
```

**证据 ·【已实测复现】**

```
$ PYTHONIOENCODING=utf-8 ./.venv/Scripts/python.exe tools/probes/audit_layout.py
（§2 段）
列宽数组完整（3 列）→ 单元数 = 6，文本 = ['A1','B1','C1','A2','B2','C2']
列宽数组截短（1 列）→ 单元数 = 2，文本 = ['A1','A2']
→ 静默少出 4 个单元，无异常、无 warning
行高数组截短（1 行）→ 单元数 = 3，文本 = ['A1','B1','C1']
```

**可达性**：`read_pages` 用 `len(table.columns)`/`len(table.rows)` 生成数组，长度**必然对齐**；所以正常导入路径不会触发。可达路径是**被手改/被别的工具生成的 `deck.json`**（契约把 deck.json 称为"唯一真源"），`load_deck` 不做几何一致性校验。与 M3 同类（静默少单元），但触发条件更窄，故定低。

---

### L3（低）· 被改过的 `deck.json` → 裸崩，而非 `IR_MISMATCH`

**位置**：`hl_layout.py::_make_rect`（除零）与 `_block`/`_text`（`None` 参与算术）。

**证据 ·【已实测复现】**

```
$ PYTHONIOENCODING=utf-8 ./.venv/Scripts/python.exe tools/probes/audit_layout.py
（§6 段）
  picture(left=None) -> TypeError: unsupported operand type(s) for /: 'NoneType' and 'int'
  canvas width_emu=0 -> ZeroDivisionError: division by zero
```

`_shape_from`（`pptx_io.py:504`）对几何字段不做校验，`None` 会被原样装进 `ShapeInfo`；`load_deck` 的 `except (KeyError, TypeError, ValueError)` 只包住**构造期**，构造期不报错，崩在 `build_units`。结果：契约 §2.1 强调"EMU 是源，只读实现不得反向由 px 推 EMU"、§5 强调 deck.json 是"唯一真源"，但这个真源被改坏时给的信号是 `INTERNAL`。同 TEST_REPORT 的 F4 属同一族（裸异常 vs `PptxError`），触发点在 `hl_layout` 而非 `build_player`。

---

### L4（低）· 播放器不校验底图存在/新鲜 → 黑屏静默通过

**位置**：`hl_anim.py::build_player`（只做 `_check_bg_path` 的**形状**校验，见 `hl_anim.py:317-318`）。

`_check_bg_path` 只回答"这个字符串是不是 out_dir 内的相对路径"，不回答"文件在不在"。缺图时：`<img>` 触发 `onerror` → `hl.ready` 里 `im.onload = im.onerror = res` **同样 resolve** → `fit()`/`render()` 照跑 → 截图得到 `.frame{background:#000}` 的黑图，**没有任何错误**。步 5 的 `shot_player` 把这种帧直接喂给 `ffmpeg`。

**【静态推断】**（未构造缺图场景实跑；机制直接读自 `hl_anim.py:211-213` 与模板 CSS）。建议：`build_player` 增加 `check_exists: bool = True`，对每个 bg 做 `os.path.isfile(os.path.join(out_dir, p))`，缺失即 `IR_MISMATCH`（文案指向"重新执行 pptgen import"）。契约 §6.5 的验收里"画面非空白"正需要这道闸门兜底。

---

### L5（低）· `had_powerpoint` 的采样窗口（TOCTOU）

`had_powerpoint = _powerpoint_running()` 在 `pptx_io.py:599`，`DispatchEx` 在 `:609`。两者之间用户**启动** PowerPoint 且**尚未打开任何稿**时：结束时 `had_powerpoint == False` 且 `app.Presentations.Count == 0` → `Quit()` 会关掉用户刚启动的 PowerPoint。D1 的加固堵住的是"事先就开着"（前一轮 TEST_REPORT §3 的 B 场景实测通过），**没堵住"这期间打开"**。【静态推断】

影响有限（PowerPoint 起手界面无未保存数据），但既然模块的立场是"宁可留一个空进程，也不关用户的东西"（`pptx_io.py:20`），这条与立场不一致。**建议**：把"资格判定"从"启动前有没有进程"改为"**我们是不是这次创建实例的人**"——例如在 `DispatchEx` **之后立刻**再探一次进程/PID 数，只有"探之前没有、探之后才有且只有一个"才认为实例归我们；或干脆改成"永不 Quit，只在 `--no-com` 之外留一个空进程"（代价是常驻一个隐形 PowerPoint）。后者最符合该模块已声明的价值取向。

---

### L6（低）· 共享 PowerPoint 实例上没有互斥

`export_pages` 全程没有锁。契约 §3.4-6 只要求线程内 `CoInitialize`，计划 R5 又要" >50 页走后台"。两次并发导出会在**同一个** PowerPoint 实例上交错（打开/导出/关闭互相干扰）。守卫方向是保守的（`Count` 非 0 就不 Quit），所以**不会破坏用户数据**；但 `Export` 可能因模态状态失败，且失败信息（"检查是否被其他程序占用"）会把责任推给用户。【静态推断】

**建议**：模块级 `threading.Lock` 串行化 `export_pages`（与 `app.py::_video_jobs` 同思路），CLI 与工作台共用。步 8 落地前处理。

---

### L7（低）· `measure_coverage` 用 `0.0` 重载三种语义

**证据 ·【已实测复现】**

```
$ PYTHONIOENCODING=utf-8 ./.venv/Scripts/python.exe tools/probes/audit_layout.py
（§5 段）
  空宽矩形 -> 0.0
  负宽矩形 -> 0.0
  完全在图外 (1000,1000,10,10) -> 0.0
  ↑ 三种都返回 0.0，与『rect 里真的没有墨迹』不可区分（验收脚本据此报低分）
  纯黑块 (16,16,32,32)（环即墨迹）-> 0.0   ← 已知盲点（用例已锁）
  同一矩形显式传底色 -> 1.0
```

契约 §4.5 只规定了"空矩形返回 0.0"。问题是**越界矩形也被裁剪后返回 0.0**：若高亮定位算错到图片外（正是 M1/L3 那类 bug 的后果），度量给出的证据与"这块是空白"完全一样——**度量本身失去了发现"定位算错"的能力**。前一轮 TEST_REPORT F6 已量化了底色盲点；这里补的是"退化/越界 → 0.0"这一条。

**建议**：越界或退化时返回 `None`（或抛 `ValueError`），让调用方必须显式处理；验收脚本据此能把"定位算错"和"本来没墨迹"分开。

---

### 提示（N1–N4）

- **N1 · 契约 §2.3 自相矛盾，实现选了 `children` 一侧**：§2.3 的 `PageShapes.shapes` 注释写"**已展开组合**、已过滤"，但同一节的 `ShapeInfo.children` 又定义"kind == group 时有效（子坐标已换算为绝对 EMU）"。实测 `read_pages` 保留 group 节点（`audit_pptx_io.py` 之外的一次直接观察：`fixtures_cover.pptx` 第 0 页顶层形状 = `[('text',…), ('group', 2)]`）。`hl_layout.walk` 依赖这一读法。**建议契约把"已展开"改成"组合以 children 嵌套表示"**；实现侧无需改。
- **N2 · payload 里有从未被消费的字段**：`build_player` 把 `CFG.title` 与每个单元的 `t`（文本）写进 HTML，但模板 JS 从不读它们（`window.hl` 只用 `l` 几何、`k` 也没用）。这解释了为什么"单元文本注入"在任何上下文都无落点（我实测 4 例全灭），是**优点**；但也意味着这些字节纯属冗余（10 页稿的 `units.json` 会明显变胖）。若将来要在播放器里显示讲解文字，**必须重新走一遍 L4 之外的 HTML 文本上下文审计**（`u.t` 一旦进 DOM，就是新的注入点）。
- **N3 · 测试数字三处不一致 + 素材被 gitignore**：`docs/PPTX_INTERFACE.md`/`PLAN`/`IMPL_REPORT` 写"248 基线"，`KNOWLEDGE.md:53` 写"当前 **244** 条"（已过时）；`IMPL_REPORT` 报"545 passed"，而本次 `--collect-only` 是 **960 collected**（前一轮测试 agent 又加了用例）。更重要的是：`output/` 整个在 `.gitignore` 内（`git check-ignore` 命中 `output/b_multislide.pptx`、`output/spike/m1/slide_1.png`、`output/fixtures/fixtures_cover.pptx`；`git ls-files output/` 为空）→ **新克隆上 `needs_pptx_src` / `needs_material` 标记的用例全部静默 skip**，其中就包括 M1 的两条核心验收与 M2 的两条验收（`test_line_coverage_beats_shape_level_by_2_5x`、`test_no_line_rect_escapes_its_shape_box`）。即"545 全绿"是**本机产物**，不是仓库可复现的事实。**建议**在 CI/交接说明里显式标注"这两条验收需要先跑 `tools/probes/accept_m1.py` 生成底图"，或把小体积素材入 git。
- **N4 · `app.py::_export_pdf_via_com` 仍是老写法**（不在本次三模块范围内，但同一仓库、同一威胁）：`app.py:753-757` 用 `Dispatch` + **无条件** `pres.Close()` + **无条件** `app.Quit()`，且失败一律 `return False`。这正是 `pptx_io` 花两道守卫防的事。既然 D1 的结论已经证实"COM server 是共享的"，这条路径**可能关掉用户的 PowerPoint**；本次审计不越权改它，但建议列入后续统一（把它改成调用 `pptx_io` 的守卫版本，或至少加 `had_powerpoint` 守卫）。

---

## 3. 已核查且未发现问题（避免后人重复劳动）

| 查了什么 | 怎么查的 | 结论 |
|---|---|---|
| **`app.Visible` 红线** | 全仓 grep `Visible\s*=[^=]` → **零命中**（只有 `tools/probes/indep_d1_com.py` 的两行**读取**）；`pptx_io.py` 里 "Visible" 只出现在注释 | ✅ 无任何代码路径写 `app.Visible`，与契约 §3.4-3 及 D1 的加强理由一致 |
| **`Quit()` 的异常路径方向性** | 逐条枚举 `finally`（`pptx_io.py:645-662`）：`pres is None`→跳过 Close；`Close()` 抛→吞掉，`Count` 仍 ≥1→不 Quit；`Presentations.Count` 抛（RPC 断开）→吞掉→不 Quit；`had_powerpoint=True`→不 Quit | ✅ **每条失败路径都偏向"不 Quit"**（宁可泄漏一个空进程），不存在"异常导致误 Quit 用户实例"的路径。唯一缺口是 S1 的 `Close` 与 L5 的 TOCTOU |
| **`_powerpoint_running()` 的探测失败语义** | 读码：`except (OSError, SubprocessError) → return True`（保守）；`"POWERPNT.EXE" in out.upper()`；`tasklist` 用 `subprocess.run([...])` 无 `shell=True` | ✅ 探不到时当作"有"，绝不 Quit；无命令注入面 |
| **底图路径穿越** | `tools/probes/audit_inject.py` §三（17 例）+ `audit_paths.py`（8 例编码形态 + 4 例对照）：`../`、`bg/../../x`、`..\..\x`、`C:/`、`C:\`、`D:relative.png`、`//srv/…`、`\\srv\…`、`/etc/passwd`、`/x.png`、`bg/a:b.png`、`bg/....//x.png`、`sub/../bg/1.png`、`bg/..%2f..%2fetc.png`、`bg/%2e%2e/x.png` | ✅ **全部按预期**：越界/盘符/UNC/绝对路径一律 `IR_MISMATCH`；合法的相对路径与"仍在 out_dir 内的 `..`"放行。特别记：`/etc/passwd` 也被拦住（Windows `ntpath.isabs("/x")` 为真）——**我最初静态怀疑它能绕过，实测证否** |
| **`%2f` 绕过（TEST_REPORT F5）** | `audit_paths.py`：`_check_bg_path("bg/..%2f..%2fetc.png")` 放行，但 `_web_path` → `bg/..%252f..%252fetc.png` | ✅ **修正 F5**：`_check_bg_path` 确实不解码百分号，但所有出口都经 `_web_path`（`%`→`%25` 先编码），解码一次只得到字面文件名 `..%2f..%2fetc.png`，**不构成路径分隔符** → 越界不可达。保留意见：若将来有别的调用方直接拿 `bg` 值拼 URL，问题就会活过来（建议仍补一道解码校验） |
| **播放器 HTML 注入** | `audit_xss_browser.py`：12 例恶意 payload（title 经典闭合 / 属性逃逸 / `<svg onload>` / `<!--`+script 的 mXSS 形状 / 大写 `</TITLE><SCRIPT>` / 占位符名 / 占位符+闭合 / bg 路径引号与闭合 / 单元文本 3 种）× **真 Chromium** `page.set_content`，判据 `window.__pwned` + `document.querySelectorAll('script').length==1` + dialog 事件 | ✅ **12/12 无可执行注入**；`<title>` 与 `<h1>` 均为 `html.escape` 后的实体；`__CONFIG__` 单遍替换未被二次替换；**阳性对照有效**（手工在模板里插入 `window.__pwned=99` 后判据立刻命中 99，证明这套判据不是"永远通过"） |
| **`<script>` 字符串上下文** | 同上 + 字符串审计：`_js_json` 对 `</`→`<\/`、`<!--`→`<\!--`；文档里**从不出现裸 `<!--`**（标题走 `html.escape`，JSON 走 `<\!--`），故 `-->` 单独出现不进入任何特殊状态 | ✅ 浏览器实测 `document.scripts == 1`；JSON 里的裸 `<script>` 字样在 script-data 状态下**不构成元素**（我的初版字符串计数器误报过 `script 开标签数=2`，浏览器判定澄清了这一点） |
| **单遍替换** | `re.sub("|".join(_PLACEHOLDERS), lambda m: mapping[m.group(0)], TEMPLATE)`——lambda 替换不解释反斜杠；`__TITLE__`/`__BGJSON__`/`__UNITS__`/`__CONFIG__` 写进 title 均不被二次替换（实测 B/C/F 三例） | ✅ 与 `anim.py` 纪律等价，无链式 replace |
| **数值参数进 JSON，不进拼接** | `dim`/`auto_step_ms`/`canvas_*` 都先 `float()/int()` 强转 + 范围校验，再 `json.dumps`；`dim` 越界（0 / 0.6 / -0.1 / 1.0）与非数值（`"深一点"`）全被用例覆盖 | ✅ 无注入面 |
| **断行等价性（防漂移闸门本身）** | 我**不复用**实现者的语料/种子：12 类语料 × 5 字号 × 5 行宽 = **350 组** + 固定种子 `424242` 的 **4000 组 fuzz**（含 `\n`、全角标点、反斜杠、`<>` 等） | ✅ **diverge = 0**。`wrap_lines` 与 `qa.measure_text_lines` 确实同源（分叉只在 `_paragraph_lines`，见 M1） |
| **XXE / billion laughs** | `audit_pptx_io.py` §3：读库源码 + 真跑一段 10 层实体炸弹 | ✅ 库内 `etree.XMLParser(remove_blank_text=True, resolve_entities=False)`；炸弹实测 0.002s 抛 `XMLSyntaxError: Maximum entity amplification factor exceeded`（lxml 自身还有放大系数上限）→ **双重防护** |
| **python-pptx 是否联网（外部关系）** | `audit_pptx_io.py` §4：遍历库源码搜 `urllib/requests/socket/http.client/urlopen/httplib` | ✅ **零命中**；`TargetMode="External"` 的关系只被记录、不会被解引用 |
| **零 iframe / 同源隔离** | `hl_anim.py` 全文无 `iframe`；用例 `test_zero_iframe` 断言 `"<iframe" not in doc and "sandbox" not in doc` | ✅ 满足契约 §6.2，比 `anim.py` 的攻击面更小（无需 sandbox 策略） |
| **COM 释放顺序的噪音** | 读 `pptx_io.py:657-660` 注释 + 前一轮 `probe_com_release_order.py` / `indep_com_release.py` 的存在 | ✅ 代码已显式说明 `RPC_E_DISCONNECTED` 是噪音且不外泄；本次未复跑（要启动 PowerPoint） |
| **`_is_hidden` 的 nv*Pr 分支** | 读码：逐 tag 试 `nvSpPr/nvPicPr/nvGraphicFramePr/nvGrpSpPr/nvCxnSpPr`，取不到一律当可见 | ✅ 各帧类元素只有一个 nv*Pr，提前 return 不丢分支；取不到时"当可见"是安全侧 |
| **新测试是否"永远通过"** | 读 3 个测试文件；本次**独立复跑**（不启动 PowerPoint 的子集）`pytest tests/test_pptx_io.py tests/test_hl_layout.py tests/test_hl_anim.py -k "not export_pages"` → **292 passed, 5 deselected** | ✅ 断言都打在真实产物上（解析生成的 JSON、比对像素、比对 EMU 精确值），未见恒真断言或被 mock 掉的被测逻辑。**唯一例外是 L1 的 `test_encrypted_code`（把误判锁成期望）** |

---

## 4. 契约一致性（D1–D9 逐条）

| # | 报告的处理 | 审计判定 |
|---|---|---|
| **D1** | `DispatchEx` 不隔离 → 加"启动前有无 POWERPNT.EXE"第二道守卫 | **方向正确、必要、但不够**：堵住了"用户开着 PowerPoint（无稿）"（前一轮真机已验），**没堵住** `Close()`（S1）与采样窗口（L5）。另外 `pre_count` 成了死变量（第 613 行赋值后全文件无引用）——契约 §3.4-5 的伪码用的是它，实现换了更严的判据但没申报。✅ 未申报但属加固 |
| **D2** | 0.35 不可达（三项分解自洽），改用相对断言 `行级/形状级 ≥ 2.5×` | **申报充分、证据可复现**（前一轮又用三套独立口径交叉验证）。**但**：把契约的 M2 验收标准换成相对口径这件事，报告里写的是"供架构 agent 定夺，我没有自行放宽契约"，而**测试文件里已经落地了这个替换**——形式上是实现者替架构 agent 做了决定。判定：**合规但有流程瑕疵**，建议架构 agent 在契约里回写一句"以 xx 为准"，否则下一位读者会以为契约仍是 0.35 |
| **D3** | `font_scale` 存储口径 ÷1000、用法 ×100，实现里 `/100` 取真实倍率 | ✅ 与报告一致，代码在 `hl_layout.py:287`；用例 `test_build_units_font_scale_scales_size_and_warns` 锁定 |
| **D4** | 标题判定第 ② 条（ph type）无 IR 字段 → 只实现 ①③ | ✅ 与报告一致（`hl_layout.py:205-216`）。影响面已申报；契约未修 |
| **D5** | 合并单元格 span 信息缺失 → origin 格按单列宽 | ✅ 与报告一致；前一轮 TEST_REPORT 已量化（本例 2 行 vs 实际 1 行） |
| **D6** | 补 `page_shapes(page, deck)` 桥接；缺 deck 时中文报错 | ✅ 实现存在（`hl_layout.py:424-439`）。**注意**它是 `ValueError` 而非 `PptxError`，与契约 §3.1"唯一的对外异常"不严格一致（低）；`build_units` 收 dict 时也抛 `ValueError`（同） |
| **D7** | Length 形态行距折算会重复计入 1.25 | ✅ 与报告一致（`pptx_io.py:205-208` 折算、`hl_layout.py:299` 使用）；本素材无显式行距 |
| **D8** | 图表 coverage 天然接近 1 不成立（实测 0.166） | ✅ 与报告一致 |
| **D9** | "40 形状"口径 | ✅ 与报告一致（41 原始 = 40 文本框 + 1 图表 → 过滤后 23） |
| **未申报的偏离（本次新增）** | — | ① **M1**：`_paragraph_lines` 的 `\n` 硬拆分支不在 §4.3 的算法里；② **S1 旁的 `pre_count`**：守卫改用 `Count==0` 未申报；③ **§2.3"已展开组合"** 与 `children` 读法冲突（N1）；④ **§6.3 要求 `html.escape(p, quote=True)`**，实现改用"百分号编码 + JS `setAttribute`"，**强度等价或更强**（浏览器实测通过），但报告只列了测试、未声明这是对契约写法的替换；⑤ `read_pages` 增了 `mode`/`export_width_px` 两个可选参数（报告有提及，属申报） |

**契约与实现的其它一致性抽查（均一致）**：`export_pages` 返回绝对路径且顺序与页序一致（用例锁定）；`slide_i.png` 从 1 起；高度按画布比例取整；`out_dir` 用 `makedirs(exist_ok=True)` 且不清理同名文件；线程内 `CoInitialize/CoUninitialize`；每页导出后校验存在且非空；过滤顺序（`no_xfrm`→`zero_size`→`hidden`→`empty_text`）与 §2.3 表一致；`kind=="other"` 保号进 `shapes` 并记 `unsupported`（报告已申报该处理）；组合儿童坐标绝对化 + 旋转叠加 + 逐层累乘（用例 `test_nested_group_mapping_is_cumulative` 锁定）。

---

## 5. 我无法验证的 / 建议后续补测

1. **PowerPoint 对"已打开文件再次 `Open`"的真实语义** —— 直接决定 S1 是"严重"还是"无风险"。需要真机 + 一份**可牺牲**的稿：先手工打开它（含未保存修改），再跑 `export_pages(同一文件)`，观察该稿是否被关/是否弹保存框。**按任务卡纪律，本次未启动 PowerPoint。**
2. **`DispatchEx` 是否附着用户实例（D1 的前提）** —— 我未复跑；证据来自实现者的 `verify_v8c_isolation.py` 与前一轮 `indep_d1_com.py`（两者结论一致）。本报告把 S1/L5 建立在"该结论成立"之上。
3. **老 `.ppt` 的实际误判**（L1）—— 需要一个真实 `.ppt` 文件；本次只能用合成 OLE2 头验证分类器行为。
4. **`bg/..%2f..%2fetc.png` 在浏览器端是否真取不到**（F5 修正的最后一环）—— 本次只做到"编码链算术 + 不落文件"；完整验证需要在 `file://` 下起播放器观察 `naturalWidth`。
5. **全量 `pytest tests/ -q`** —— 会启动 PowerPoint（5 条 `export_pages` 用例）与浏览器（`test_shot_settle`），超出"只读、不启动 PowerPoint"的纪律；本次只做了 `--collect-only`（960 条）与不含 COM 的 292 条子集。
6. **>50 页后台路径的并发行为**（R5 / L6）—— 步 6/8 未实现，无法测。
7. **真实手工 PPT 的字体替换、SmartArt、嵌入视频**（R6/V12 的已知边界，前一轮已声明）—— 仓库内没有此类素材；M1 的 a:br 分叉同理只在合成段落上验过。

---

## 6. 给下一角色的最小行动清单

| 优先级 | 动作 | 位置 |
|---|---|---|
| P0 | 让 `Close` 与 `Quit` 对称：记下 `Open` 前 `app.Presentations` 的名字集合，`finally` 里只关自己新开的那份 | `pptx_io.py:645-650` |
| P0 | 补 V8 的缺口实验：用户**开着同一份稿**时跑 `export_pages`，确认该稿存活（写进 `tools/probes/`） | 新增探针 |
| P1 | `read_pages` 加 zip 门禁（`infolist()` 累计 `file_size` + 比率上限），并把 `width_emu/height_emu` 的读取纳入 `try` | `pptx_io.py:461-471` |
| P1 | `build_units` 增加"被丢弃形状"的出口，与 `read_pages.skipped` 同构；CLI 冒泡 | `hl_layout.py:280` |
| P1 | 架构 agent 裁定 M1（改 `qa._tokenize` 认 `\n`，还是改契约 §4.3 + 把等价测试换成 `_paragraph_lines`） | `qa.py` / 契约 |
| P2 | `_classify_bad_package` 区分 `.ppt` 与加密 pptx；同步修掉把它锁为期望的测试 | `pptx_io.py:439` |
| P2 | `measure_coverage` 退化/越界返回 `None`（或抛错），不再与"无墨迹"共用 `0.0` | `hl_layout.py:469-479` |
| P2 | `build_player` 校验 bg 文件存在；`export_pages` 加模块级锁；`app.py::_export_pdf_via_com` 统一到守卫版本 | `hl_anim.py` / `pptx_io.py` / `app.py:744` |
| P3 | 修 `KNOWLEDGE.md` 的"244 条"；在 CI/交接说明里标注"素材在 gitignore 内，两条 M2 验收需先生成底图" | `KNOWLEDGE.md:53` |
