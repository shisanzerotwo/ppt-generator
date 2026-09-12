"""独立验证 ⑥ 附录 · 把 `export_pages` 端到端 ~27s 的账做平。

已知分量：DispatchEx 启动 ~5.8s、Open ~0.2s、10 页 Export ~1.5s、Close ~0.1s、
Quit() 立即返回但进程延迟退出、CoUninitialize ~2.7s。合计 ~10s，与实测 ~27s 仍差 ~17s。

本探针隔离两个剩余嫌疑：
  1. **附着到已运行实例**时 export_pages 要多久（此时不启动、不 Quit）→ 得到"纯导出"成本；
  2. `_powerpoint_running()` 里的 `tasklist` 子进程成本（冷/热各测一次）。

⚠️ 前置闸门：启动前已有 POWERPNT.EXE 则整条中止。

一条命令：./.venv/Scripts/python.exe tools/probes/indep_export_cost.py
"""

import os
import subprocess
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

import pptx_io  # noqa: E402

SRC = os.path.join(ROOT, "output", "b_multislide.pptx")
OUT = os.path.join(ROOT, "output", "spike", "indep_cost")

TASKLIST = ["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/NH"]


def pids() -> list[str]:
    out = subprocess.run(TASKLIST, capture_output=True, text=True).stdout
    return sorted(ln.split()[1] for ln in out.splitlines()
                  if "POWERPNT.EXE" in ln.upper() and len(ln.split()) >= 2)


def main() -> int:
    pre = pids()
    print(f"[闸门] 启动前 POWERPNT.EXE PIDs = {pre or '[]（无）'}")
    if pre:
        print("[ABORT] 已有 PowerPoint 进程，整条中止。")
        return 3

    print("\n== tasklist 子进程成本（_powerpoint_running 用的就是它）==")
    for label in ("cold", "warm#1", "warm#2"):
        t0 = time.time()
        subprocess.run(TASKLIST, capture_output=True, text=True)
        print(f"  {label:8} = {time.time() - t0:.2f}s")

    import pythoncom
    import win32com.client

    os.makedirs(OUT, exist_ok=True)
    pythoncom.CoInitialize()

    print("\n== 场景甲：自己起实例（冷）==")
    t0 = time.time()
    shots = pptx_io.export_pages(SRC, os.path.join(OUT, "cold"), width=1920)
    t_cold = time.time() - t0
    print(f"  export_pages = {t_cold:.2f}s（{len(shots)} 张）")
    time.sleep(2.0)
    print(f"  之后 PIDs = {pids() or '[]（无）'}")

    print("\n== 场景乙：先留一个『用户实例』，再让 export_pages 附着 ==")
    # 场景甲之后 apartment 很可能已被拆（见 TEST_REPORT 的 D1 附录），这里必须先重新 init
    try:
        win32com.client.Dispatch("PowerPoint.Application")
    except Exception as exc:
        print(f"  [证据] 场景甲之后直接 Dispatch 失败：{type(exc).__name__}: {exc}")
        pythoncom.CoInitialize()
        print("  已重新 CoInitialize()，继续")
    app = win32com.client.Dispatch("PowerPoint.Application")
    print(f"  用户实例就绪（Count={app.Presentations.Count}），PIDs={pids()}")
    t0 = time.time()
    shots2 = pptx_io.export_pages(SRC, os.path.join(OUT, "warm"), width=1920)
    t_warm = time.time() - t0
    print(f"  export_pages = {t_warm:.2f}s（{len(shots2)} 张）← 纯导出成本（无启动/无 Quit）")
    alive = pids()
    print(f"  之后 PIDs = {alive or '[]（无）'}  → 用户实例是否被保住："
          f"{'是' if alive else '否'}")

    print("\n== 账目对照 ==")
    print(f"  冷调用端到端      = {t_cold:.2f}s")
    print(f"  附着调用端到端    = {t_warm:.2f}s")
    print(f"  两者之差（启动+Quit+释放）= {t_cold - t_warm:.2f}s")
    print(f"  → 契约『≤1.5s/页』在附着场景：{t_warm / len(shots2):.2f}s/页；"
          f"冷调用场景：{t_cold / len(shots):.2f}s/页")

    # 收尾
    try:
        pythoncom.CoInitialize()
        a = win32com.client.Dispatch("PowerPoint.Application")
        while a.Presentations.Count > 0:
            a.Presentations(1).Close()
        a.Quit()
    except Exception as exc:  # noqa: BLE001
        print(f"  Quit 失败：{type(exc).__name__}")
    time.sleep(3.0)
    print(f"  收尾 PIDs = {pids() or '[]（无）'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
