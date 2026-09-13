# pptx 遗留风险修复 · 收尾计划

## Context

**背景**：pptx 双向链路已交付（19 commit / 1028 tests），随后独立测试与只读审计发现 9 个遗留问题（F1/F2/F3/M1–M4/L1/L4）。按用户决定开启修复轮次，交给 claude agent（`ppt-impl`）执行，任务卡为 `docs/TASK_FIX.md`。

**当前进度：9 项已完成 7 项**，各自独立 commit：

| 项 | 内容 | commit |
|---|---|---|
| M2 | zip bomb 闸门（`read_pages` 前置 `_check_package_safety`） | `544e7b1` |
| M4 | 缺 `p:sldSz` → 错误码从 INTERNAL(5) 归回 PPTX_UNREADABLE(3) | `2e72493` |
| L1 | 老 `.ppt`（OLE2）不再被误判成「已加密」 | `da574f8` |
| M3 | 形状静默丢弃 → `build_units` 增 `skipped` 收集器 | `692d83d` |
| M1 | `a:br` 几何与 `qa` 分叉 → 硬换行抽成公共规则 | `9b4e3c0` |
| F2 | `wrap="none"` 导致高亮精度退化为段级 → 记 warning 声明降级 | `01cb207` |
| L4 | 播放器缺图不再静默 resolve `hl.ready` | `8bd539c` |

**未完成**：F1（探针 `tools/probes/fix_f1_apartment.py` 已写 155 行，未验证/未提交）、F3（复用实例）、`docs/FIX_REPORT.md`。

### ⚠️ 硬阻塞：venv 已不可用（同时卡住 claude）

```
pyvenv.cfg   home = ...\uv\python\cpython-3.11-windows-x86_64-none      ← uv 0.11 的旧命名
磁盘实际      ...\uv\python\cpython-3.11.15-windows-x86_64-none          ← uv 0.12+ 的新命名
```

`.venv/Scripts/python.exe` 是 uv 的 trampoline，找不到基座解释器 → 任何 python 命令都报
`uv trampoline failed to spawn Python child process / entity not found (os error 2)`。

**证据**：编排者连 `import sys` 都跑不了；claude agent 的诊断命令与 pytest 同样报此错（其 pytest 进程还因此 CPU=0 假死）。**未修复前，剩余 F1/F3 的验证一步都做不了。**

## Approach

三段推进：**先修环境**（否则一切停滞）→ 让 claude 收尾剩余项 → 编排者独立验证与收尾。

- **环境**：只改 `.venv/pyvenv.cfg` 一行（该文件已被 `.gitignore` 覆盖，不进仓库）
- **F1**：按任务卡用 `CoInitializeEx` 的 **HRESULT** 判定是否自初始化，只在 `S_OK` 时 `CoUninitialize`；必须做 git stash 前后**阳性/阴性对照**
- **F3**：模块级单例 app + `threading.Lock` + 复用路径不 `Quit`；**保留 S1（`we_opened`）与 D1（`had_powerpoint`）守卫语义**；实测收益不明显或与数据安全冲突则**放弃**（任务卡已授权）

## Files to modify

| 文件 | 改动 | 执行者 |
|---|---|---|
| `.venv/pyvenv.cfg` | `home` 指向实际存在的 `cpython-3.11.15-windows-x86_64-none` | 编排者（一行） |
| `pptx_io.py` | F1：`CoInitializeEx` HRESULT 判定 + 条件 `CoUninitialize` | claude |
| `pptx_io.py` | F3：单例 app + 锁 + 复用不 Quit + `release_app()` + CLI `--no-reuse` | claude（可放弃） |
| `cli.py` | F3 配套：`--no-reuse` 回退开关 | claude |
| `tests/` | F1/F3 的防回归用例（照 `tests/test_s1_close_guard.py` 的打桩范式） | claude |
| `tools/probes/fix_f1_apartment.py` | 补齐阳性/阴性对照 | claude |
| `tools/probes/fix_f3_reuse.py` | 新建：同进程连续两次 `export_pages` 的秒数对比 | claude |
| `docs/FIX_REPORT.md` | 四段式报告（根因/修法/前后对照实测/新增用例数） | claude |
| `KNOWLEDGE.md` | 风险表把已修项标记「已修复 + commit」 | 编排者 |
| `AGENTS.md`（知识库源） | 会话日志补修复轮次 + **新增踩坑：uv 升级导致 venv python 目录改名** | 编排者 |
| `.gitignore` | 若 `spike/` 判定为临时产物则加入忽略 | 编排者 |

## Reuse

- `tools/probes/s1_probe_v2.py` —— **探针范式**（独立进程观察者 + 安全三重闸门 + 阳性/阴性对照）
- `tools/probes/fix_f1_apartment.py` —— claude 已写的 F1 探针（155 行，待验证）
- `tests/test_s1_close_guard.py` —— **打桩 COM 的测试范式**（F1/F3 的防回归用例照它写，不启动 PowerPoint）
- `docs/TASK_FIX.md` —— 本轮任务卡，每项的修法与验收标准都在这
- `qa.py` 的字体度量（M1/F2 已复用的既有能力，勿重复造）

## Steps

- [ ] **修 `.venv/pyvenv.cfg`** 的 `home` → 验证 `.venv/Scripts/python.exe -c "import sys;print(sys.version)"` 输出 3.11.15
- [ ] 让 `ppt-impl` 继续：F1（跑 git stash 前后对照）→ F3（实测秒数，收益不明即放弃）→ `docs/FIX_REPORT.md`
- [ ] 跑全量 `.venv/Scripts/python.exe -m pytest tests/ -q`（预期 ~1061 条）确认全绿
- [ ] **编排者独立验证**：`git diff e52504f HEAD --stat -- anim.py builder.py` 必须为空；抽查 F1 对照输出与 F3 秒数
- [ ] 提交未提交产物（`docs/TASK_FIX.md` + 两个探针；`spike/` 按判定提交或清理并忽略）
- [ ] 推送（走 Windows 侧 git，WSL 连不上 github.com:443）
- [ ] 更新 `KNOWLEDGE.md` 风险表（已修项标记）；`AGENTS.md` 补会话日志 + uv 踩坑，并**重新分发 5 agent**
- [ ] 关闭 `ppt-impl`（`w9:t5`）回收现场

## Verification

| 项 | 命令 | 通过标准 |
|---|---|---|
| 环境 | `.venv/Scripts/python.exe -c "import sys;print(sys.version)"` | 打印 `3.11.15` |
| F1 | `.venv/Scripts/python.exe tools/probes/fix_f1_apartment.py` | 修复后**旧 COM 代理仍可用**（`CO_E_NOTINITIALIZED` 消失）；stash 前后对照成立 |
| F3 | `.venv/Scripts/python.exe tools/probes/fix_f3_reuse.py` | 第二次 `export_pages` 显著快于第一次，**给出实测秒数**；无孤儿 `POWERPNT.EXE` |
| 回归 | `.venv/Scripts/python.exe -m pytest tests/ -q` | **全绿**（~1061 条） |
| 红线 | `git diff e52504f HEAD --stat -- anim.py builder.py qa.py shot.py template.py` | **空** |
| S1 未回归 | `python tools/probes/s1_probe_v2.py` | 仍报 `COUNT=1`（用户稿未被关） |
| 端到端 | `pptgen import/animate/video/export` + 工作台上传 pptx | 四命令 `ok:true`；播放器 HTTP 200 |

## 风险与对策

| 风险 | 对策 |
|---|---|
| **F3 与 S1/D1 冲突**（复用实例正是 S1 的根源） | 保留原守卫语义 + 加锁；实测收益不显著或引入数据风险则**放弃**并给数据 |
| `spike/` 目录来源不明（未跟踪，52K） | 判定为 claude 的临时探针产物 → 完成后清理或加入 `.gitignore` |
| venv 修复影响面 | 只改 `.venv/pyvenv.cfg`（已 gitignore），不触碰仓库文件 |
| claude 上下文已用 ~75k tokens | 若它中途上下文耗尽，剩余项由编排者直接完成（不依赖外部额度） |
| 两个 pytest 并发抢 PowerPoint COM | 同一时刻只允许一个 pytest 在跑（F1 失败用例曾因此假失败） |

## 交付后保留的已知风险（本轮不动）

- **F3**：若放弃，理由与实测数据写进 `docs/FIX_REPORT.md` 与 `KNOWLEDGE.md`
- 其余 L2/L3/L5/L6/L7（低危，证据已在 `docs/AUDIT_REPORT.md`）
