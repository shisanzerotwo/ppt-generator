"""独立验证 ② 附录 · 逐步复刻 export_pages 的内部序列，二分定位 apartment 被拆的那一步。

已知事实：
  · `export_pages()` 返回后同线程 apartment 已死（FSO 都报 CO_E_NOTINITIALIZED）；
  · 单独"释放 PPT 代理"、单独"一对 init/uninit"、两者组合 —— 都**不**触发（见 indep_com_release.py）。

本探针把 export_pages 的每一步在探针里逐条复刻，并在每步之后用
`Dispatch("Scripting.FileSystemObject")` 探一次 apartment 是否还活着。

⚠️ 前置闸门：启动前已有 POWERPNT.EXE 则整条中止。

一条命令：./.venv/Scripts/python.exe tools/probes/indep_apartment_bisect.py
"""

import os
import subprocess
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

SRC = os.path.join(ROOT, "output", "b_multislide.pptx")
OUT = os.path.join(ROOT, "output", "spike", "indep_bisect")


def pids() -> list[str]:
    out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/NH"],
                         capture_output=True, text=True).stdout
    return sorted(ln.split()[1] for ln in out.splitlines()
                  if "POWERPNT.EXE" in ln.upper() and len(ln.split()) >= 2)


def check(step, log):
    import win32com.client
    try:
        win32com.client.Dispatch("Scripting.FileSystemObject")
        log.append(f"  [{step}] apartment = 存活")
    except Exception as exc:  # noqa: BLE001
        code = exc.args[0] if exc.args else type(exc).__name__
        log.append(f"  [{step}] apartment = 已死（{code}）")


def main() -> int:
    pre = pids()
    print(f"[闸门] 启动前 POWERPNT.EXE PIDs = {pre or '[]（无）'}")
    if pre:
        print("[ABORT] 已有 PowerPoint 进程，整条中止。")
        return 3

    import pythoncom
    import win32com.client

    log = []
    pythoncom.CoInitialize()
    check("1 我方 CoInitialize", log)
    app = win32com.client.Dispatch("PowerPoint.Application")
    check("2 Dispatch 建实例", log)

    os.makedirs(OUT, exist_ok=True)
    subprocess.run(["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/NH"],
                   capture_output=True, text=True)          # = _powerpoint_running()
    check("3 subprocess.run(tasklist)", log)

    pythoncom.CoInitialize()                                # export_pages 内部
    check("4 export_pages 的 CoInitialize", log)

    appX = win32com.client.DispatchEx("PowerPoint.Application")
    check("5 DispatchEx", log)

    pre_count = appX.Presentations.Count
    check("6 Presentations.Count", log)

    pres = appX.Presentations.Open(os.path.abspath(SRC), ReadOnly=True,
                                   Untitled=False, WithWindow=False)
    check("7 Presentations.Open", log)

    w_emu = int(pres.PageSetup.SlideWidth)
    h_emu = int(pres.PageSetup.SlideHeight)
    height = round(640 * h_emu / w_emu)
    total = int(pres.Slides.Count)
    check("8 PageSetup/Slides.Count", log)

    pres.Slides(1).Export(os.path.abspath(os.path.join(OUT, "s1.png")), "PNG", 640, height)
    check("9 Slides(1).Export", log)

    pres.Close()
    check("10 pres.Close()", log)

    appX = None
    pres = None
    import gc
    gc.collect()
    check("11 释放 pres/appX 代理", log)

    pythoncom.CoUninitialize()                              # export_pages 的 finally
    check("12 CoUninitialize", log)

    print()
    for ln in log:
        print(ln)
    print()
    dead = [ln for ln in log if "已死" in ln]
    print("== 结论 ==")
    if dead:
        print(f"  apartment 在第 {dead[0].split(']')[0].strip('[')} 步之后死亡")
    else:
        print("  复刻序列全程 apartment 存活 —— 拆毁不是这些步骤的简单组合造成，")
        print("  需考虑 Export 的时序/多页循环或 PowerPoint 服务端主动断开。")

    # 收尾
    try:
        a = win32com.client.Dispatch("PowerPoint.Application")
        while a.Presentations.Count > 0:
            a.Presentations(1).Close()
        a.Quit()
    except Exception as exc:  # noqa: BLE001
        print(f"  Quit 失败：{type(exc).__name__}: {exc}")
    time.sleep(2.0)
    print(f"  收尾 PIDs = {pids() or '[]（无）'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
