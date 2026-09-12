"""独立验证 ② 附录 · 影响面评估：apartment 被拆之后，谁会真的坏掉？

已知：`export_pages()` 返回后，同线程 COM 处于未初始化状态（除非重新 CoInitialize），
调用方在调用前持有的 PowerPoint 代理也变成 RPC_E_DISCONNECTED。

本探针回答两个"会不会真出事故"的问题：
  1. 连续两次 `export_pages()`（同线程）—— 第二次还能跑吗？（export_pages 自身是否自愈）
  2. 调用方在 export_pages **之前**持有的 COM 代理（模拟 Flask 线程里的长生命周期对象），
     调用之后还能用吗？

⚠️ 前置闸门：启动前已有 POWERPNT.EXE 则整条中止。

一条命令：./.venv/Scripts/python.exe tools/probes/indep_d1_impact.py
"""

import os
import subprocess
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

import pptx_io  # noqa: E402

SRC = os.path.join(ROOT, "output", "b_multislide.pptx")
OUT = os.path.join(ROOT, "output", "spike", "indep_impact")


def pids() -> list[str]:
    out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/NH"],
                         capture_output=True, text=True).stdout
    return sorted(ln.split()[1] for ln in out.splitlines()
                  if "POWERPNT.EXE" in ln.upper() and len(ln.split()) >= 2)


def main() -> int:
    pre = pids()
    print(f"[闸门] 启动前 POWERPNT.EXE PIDs = {pre or '[]（无）'}")
    if pre:
        print("[ABORT] 已有 PowerPoint 进程，整条中止。")
        return 3

    import pythoncom
    import win32com.client

    os.makedirs(OUT, exist_ok=True)

    # ---------------------------------------------------- 1. 连续两次 export_pages
    print("\n=== 1. 同线程连续两次 export_pages（自身是否自愈）===")
    pythoncom.CoInitialize()
    for i in (1, 2):
        t0 = time.time()
        try:
            shots = pptx_io.export_pages(SRC, os.path.join(OUT, f"run{i}"), width=320)
            print(f"  第 {i} 次：OK，{len(shots)} 张，{time.time() - t0:.1f}s")
        except Exception as exc:  # noqa: BLE001
            print(f"  第 {i} 次：失败 {type(exc).__name__}: {exc}")
    print(f"  两次之后 PIDs = {pids() or '[]（无）'}")

    # ---------------------------------------------------- 2. 调用方持有的代理
    print("\n=== 2. 调用方在 export_pages 之前持有的 COM 代理，之后还能用吗 ===")
    try:
        pythoncom.CoInitialize()
    except Exception:  # noqa: BLE001
        pass
    fso = win32com.client.Dispatch("Scripting.FileSystemObject")
    try:
        print(f"  调用前持有 FSO 代理并可用：Drives={fso.Drives.Count}")
    except Exception as exc:  # noqa: BLE001
        print(f"  调用前就不可用：{exc}")
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    try:
        print(f"  调用前持有 PPT 代理并可用：Count={ppt.Presentations.Count}")
    except Exception as exc:  # noqa: BLE001
        print(f"  调用前 PPT 代理不可用：{exc}")

    shots = pptx_io.export_pages(SRC, os.path.join(OUT, "run3"), width=320)
    print(f"  export_pages 返回 {len(shots)} 张")

    for label, fn in (("FSO 代理", lambda: fso.Drives.Count),
                      ("PPT 代理", lambda: ppt.Presentations.Count)):
        try:
            print(f"  调用后 {label}：仍可用（{fn()}）")
        except Exception as exc:  # noqa: BLE001
            code = exc.args[0] if exc.args else type(exc).__name__
            print(f"  调用后 {label}：**失效**（{code}: {exc.args[1] if len(exc.args) > 1 else ''}）")

    # ---------------------------------------------------- 收尾
    print("\n=== 收尾 ===")
    try:
        pythoncom.CoInitialize()      # 修复被拆的 apartment 才能收尾
        a = win32com.client.Dispatch("PowerPoint.Application")
        while a.Presentations.Count > 0:
            a.Presentations(1).Close()
        a.Quit()
        print("  已 Quit 我方实例")
    except Exception as exc:  # noqa: BLE001
        print(f"  Quit 失败：{type(exc).__name__}: {exc}")
    time.sleep(2.0)
    print(f"  收尾 PIDs = {pids() or '[]（无）'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
