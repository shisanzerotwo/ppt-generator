"""独立验证 ⑥ 附录 · 收尾成本：`Close()` / `Quit()` / `CoUninitialize()` 各要多久？

`indep_m1_com.py` 实测 `export_pages(1920)` 端到端 ~27s，
`indep_m1_com_page.py` 复刻同样的启动+打开+逐页导出只要 7.45s —— 差额必然在收尾。

⚠️ 前置闸门：启动前已有 POWERPNT.EXE 则整条中止。

一条命令：./.venv/Scripts/python.exe tools/probes/indep_com_quit_cost.py
"""

import os
import subprocess
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

SRC = os.path.join(ROOT, "output", "b_multislide.pptx")
OUT = os.path.join(ROOT, "output", "spike", "indep_quit")


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
    pythoncom.CoInitialize()

    t0 = time.time()
    app = win32com.client.DispatchEx("PowerPoint.Application")
    t_launch = time.time() - t0
    t0 = time.time()
    pres = app.Presentations.Open(os.path.abspath(SRC), ReadOnly=True,
                                 Untitled=False, WithWindow=False)
    t_open = time.time() - t0
    t0 = time.time()
    pres.Slides(1).Export(os.path.abspath(os.path.join(OUT, "s1.png")), "PNG", 1920, 1080)
    t_export = time.time() - t0

    t0 = time.time()
    pres.Close()
    t_close = time.time() - t0

    t0 = time.time()
    app.Quit()
    t_quit = time.time() - t0

    # 进程真正退出要多久
    t0 = time.time()
    gone_at = None
    while time.time() - t0 < 30:
        if not pids():
            gone_at = time.time() - t0
            break
        time.sleep(0.25)

    t0 = time.time()
    app = pres = None
    pythoncom.CoUninitialize()
    t_uninit = time.time() - t0

    print("\n== 成本（秒）==")
    print(f"  DispatchEx 启动      = {t_launch:.2f}")
    print(f"  Presentations.Open   = {t_open:.2f}")
    print(f"  单页 Export          = {t_export:.3f}")
    print(f"  pres.Close()         = {t_close:.2f}")
    print(f"  app.Quit()           = {t_quit:.2f}")
    print(f"  进程实际退出          = {gone_at if gone_at is not None else '>30s 未退出'}")
    print(f"  CoUninitialize       = {t_uninit:.3f}")
    total = t_launch + t_open + t_export + t_close + t_quit + (gone_at or 0)
    print(f"  —— 合计 ≈ {total:.2f}s（对照 export_pages 端到端实测 ~27s）")
    print("\n  注：`_powerpoint_running()` 的 tasklist 子进程另有开销（未计入）。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
