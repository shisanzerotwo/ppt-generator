"""独立验证 ② 附录 · 刻画 `export_pages` 之后 COM 代理变成 RPC_E_DISCONNECTED 的现象。

在 `indep_d1_com.py` 里观察到：`export_pages` 成功返回、PowerPoint 进程仍活，
但**调用方自己持有的 `app` 代理失效**（`对象没有连接到服务器` / 0x80010108）。
实现者在代码注释里把它归为"噪音，不用管"——本探针量化它：
  1. 现象是否稳定复现（代理失效 vs 进程存活）；
  2. 失效后能否**重新取回**可用实例（用户的应用是否真的还健康）；
  3. 归因：是 `pythoncom.CoUninitialize()` 造成，还是"释放任一 COM 代理"造成。

⚠️ 前置闸门：启动前已有 POWERPNT.EXE 则整条中止。

一条命令：./.venv/Scripts/python.exe tools/probes/indep_d1_disconnect.py
"""

import gc
import os
import subprocess
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

import pptx_io  # noqa: E402

SRC = os.path.join(ROOT, "output", "b_multislide.pptx")
OUT = os.path.join(ROOT, "output", "spike", "indep_d1_disconnect")


def pids() -> list[str]:
    out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/NH"],
                         capture_output=True, text=True).stdout
    return sorted(ln.split()[1] for ln in out.splitlines()
                  if "POWERPNT.EXE" in ln.upper() and len(ln.split()) >= 2)


def probe_handle(app, label):
    """试读一个 COM 代理还活着吗。返回 (ok, 描述)。"""
    try:
        return True, f"Count={app.Presentations.Count}"
    except Exception as exc:  # noqa: BLE001
        return False, f"{type(exc).__name__}: {exc}"


def main() -> int:
    pre = pids()
    print(f"[闸门] 启动前 POWERPNT.EXE PIDs = {pre or '[]（无）'}")
    if pre:
        print("[ABORT] 已有 PowerPoint 进程，整条中止。")
        return 3

    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    try:
        # ---------------------------------------------------- 1. 现象复现
        print("\n=== 1. export_pages 前后，调用方代理的健康度 ===")
        app = win32com.client.Dispatch("PowerPoint.Application")
        ok0, d0 = probe_handle(app, "before")
        print(f"  export_pages 前：proxy_ok={ok0}  {d0}")
        print(f"  PIDs = {pids()}")

        os.makedirs(OUT, exist_ok=True)
        shots = pptx_io.export_pages(SRC, OUT, width=640)
        print(f"  export_pages 返回 {len(shots)} 张")

        ok1, d1 = probe_handle(app, "after")
        print(f"  export_pages 后：proxy_ok={ok1}  {d1}")
        alive = pids()
        print(f"  PIDs = {alive or '[]（无）'}   ← 进程是否还活着")
        print(f"  → 现象：代理 {'失效' if not ok1 else '仍可用'} / 进程 "
              f"{'存活' if alive else '已退出'}")

        # ---------------------------------------------------- 2. 重新取回
        print("\n=== 2. 失效后能否重新取回可用实例（用户应用是否仍健康）===")
        app2 = win32com.client.Dispatch("PowerPoint.Application")
        ok2, d2 = probe_handle(app2, "resh")
        print(f"  重新 Dispatch：proxy_ok={ok2}  {d2}")
        try:
            app3 = win32com.client.GetActiveObject("PowerPoint.Application")
            ok3, d3 = probe_handle(app3, "gao")
            print(f"  GetActiveObject：proxy_ok={ok3}  {d3}")
        except Exception as exc:  # noqa: BLE001
            ok3, d3 = False, f"{type(exc).__name__}: {exc}"
            print(f"  GetActiveObject 失败：{d3}")
        if ok2:
            pres = app2.Presentations.Open(os.path.abspath(SRC), ReadOnly=True,
                                          Untitled=False, WithWindow=False)
            print(f"  重新取回的实例能打开稿：Count={app2.Presentations.Count}")
            pres.Close()
        print(f"  → 用户应用是否仍健康可用：{'是' if ok2 else '否'}")

        # 收尾：用活着的代理 Quit
        for a in (app2,):
            try:
                a.Quit()
            except Exception:  # noqa: BLE001
                pass
        del app, app2
        gc.collect()
        time.sleep(1.5)
        print(f"  Quit 后 PIDs = {pids() or '[]（无）'}")

        # ---------------------------------------------------- 3. 归因
        print("\n=== 3. 归因：CoUninitialize 还是『释放代理』？===")
        print("  3a. 只做 CoInitialize/CoUninitialize，不碰 PowerPoint 业务：")
        print(f"    在此之前 PIDs = {pids() or '[]（无）'}")
        appX = win32com.client.Dispatch("PowerPoint.Application")
        okx, dx = probe_handle(appX, "x")
        print(f"    Dispatch 后 proxy_ok={okx} {dx}")
        pythoncom.CoInitialize()
        pythoncom.CoUninitialize()
        okx2, dx2 = probe_handle(appX, "x2")
        print(f"    CoInitialize+CoUninitialize 之后 proxy_ok={okx2} {dx2}"
              f"  → 归因{'成立' if not okx2 else '不成立'}")

        print("\n  3b. 释放第二个代理（不改业务）是否让第一个失效：")
        appY = win32com.client.DispatchEx("PowerPoint.Application")
        print(f"    第二个代理就绪：{probe_handle(appY, 'y')[1]}")
        del appY
        gc.collect()
        time.sleep(0.5)
        okx3, dx3 = probe_handle(appX, "x3")
        print(f"    释放第二个代理后，第一个 proxy_ok={okx3} {dx3}")

        try:
            appX.Quit()
        except Exception:  # noqa: BLE001
            pass
        del appX
        gc.collect()
        time.sleep(1.5)
        print(f"\n收尾 PIDs = {pids() or '[]（无）'}")
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:  # noqa: BLE001
            pass
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
