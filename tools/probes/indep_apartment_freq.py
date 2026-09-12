"""独立验证 ② 附录 · 频次量化：apartment 被拆是稳定的还是偶发的？

前几个探针出现了**互相矛盾**的观测：
  · indep_com_apartment2.py / indep_apartment_bisect.py：export_pages 之后 apartment 已死；
  · indep_d1_impact.py：同一个 export_pages 调用后，调用方的 FSO/PPT 代理都还活着。

本探针在 3 个**独立线程**里各跑一遍"复刻序列"（CompInitialize → Dispatch → Open →
Export 1 页 → Close → 释放 → CoUninitialize → 探 apartment），统计死亡次数。

⚠️ 前置闸门：启动前已有 POWERPNT.EXE 则整条中止。

一条命令：./.venv/Scripts/python.exe tools/probes/indep_apartment_freq.py
"""

import os
import subprocess
import sys
import threading
import time

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

SRC = os.path.join(ROOT, "output", "b_multislide.pptx")
OUT = os.path.join(ROOT, "output", "spike", "indep_freq")
ATTEMPTS = 3


def pids() -> list[str]:
    out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/NH"],
                         capture_output=True, text=True).stdout
    return sorted(ln.split()[1] for ln in out.splitlines()
                  if "POWERPNT.EXE" in ln.upper() and len(ln.split()) >= 2)


def alive() -> str:
    import win32com.client
    try:
        win32com.client.Dispatch("Scripting.FileSystemObject")
        return "存活"
    except Exception as exc:  # noqa: BLE001
        code = exc.args[0] if exc.args else type(exc).__name__
        return f"已死({code})"


def sequence(log):
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    app = win32com.client.Dispatch("PowerPoint.Application")
    _ = app.Presentations.Count
    pythoncom.CoInitialize()                     # export_pages 内部那一次
    appX = win32com.client.DispatchEx("PowerPoint.Application")
    pres = appX.Presentations.Open(os.path.abspath(SRC), ReadOnly=True,
                                  Untitled=False, WithWindow=False)
    png = os.path.abspath(os.path.join(OUT, f"s_{threading.get_ident()}.png"))
    pres.Slides(1).Export(png, "PNG", 320, 180)
    pres.Close()
    appX = pres = None
    pythoncom.CoUninitialize()                   # export_pages 的 finally
    log.append(alive())


def main() -> int:
    pre = pids()
    print(f"[闸门] 启动前 POWERPNT.EXE PIDs = {pre or '[]（无）'}")
    if pre:
        print("[ABORT] 已有 PowerPoint 进程，整条中止。")
        return 3
    os.makedirs(OUT, exist_ok=True)

    results = {}

    def run(i):
        log = []
        try:
            sequence(log)
        except Exception as exc:  # noqa: BLE001
            log.append(f"异常 {type(exc).__name__}: {exc}")
        finally:
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except Exception:  # noqa: BLE001
                pass
        results[i] = log[0] if log else "无结果"

    threads = [threading.Thread(target=run, args=(i,)) for i in range(ATTEMPTS)]
    for t in threads:
        t.start()
    for t in threads:
        t.join()

    print(f"\n=== 复刻序列 × {ATTEMPTS}（每例一个独立线程/apartment）===")
    dead = 0
    for i in range(ATTEMPTS):
        print(f"  第 {i + 1} 例：CoUninitialize 之后 apartment = {results.get(i)}")
        if "已死" in str(results.get(i)):
            dead += 1
    print(f"\n  → 死亡 {dead}/{ATTEMPTS} 次："
          f"{'稳定复现' if dead == ATTEMPTS else ('偶发' if dead else '本次未复现')}")

    # 收尾
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
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
