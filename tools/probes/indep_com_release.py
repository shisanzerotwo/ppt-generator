"""独立验证 ② 附录 · 触发点定位：拆掉 apartment 的到底是哪一步？

已知（indep_com_apartment2.py 实测）：`export_pages()` 返回后同线程 apartment 已死
（连 `Scripting.FileSystemObject` 都报 CO_E_NOTINITIALIZED），重新 CoInitialize 即恢复。
已知（indep_com_apartment.py T2）：**单独**多一对 init/uninit 不会拆 apartment。

于是只剩两个候选触发点，本探针在**各自独立的线程**里分别验证：
  A. 只"创建 + 释放" PowerPoint COM 代理（不调用 export_pages、不做 CoUninitialize）
  B. 只做 init/uninit 那一对（不碰 PowerPoint 代理）

一条命令：./.venv/Scripts/python.exe tools/probes/indep_com_release.py
"""

import gc
import os
import subprocess
import sys
import threading
import time

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)


def pids() -> list[str]:
    out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/NH"],
                         capture_output=True, text=True).stdout
    return sorted(ln.split()[1] for ln in out.splitlines()
                  if "POWERPNT.EXE" in ln.upper() and len(ln.split()) >= 2)


def _try(progid):
    import win32com.client
    try:
        win32com.client.Dispatch(progid)
        return "OK"
    except Exception as exc:  # noqa: BLE001
        code = exc.args[0] if exc.args else type(exc).__name__
        return f"FAIL({code})"


def case_release(log):
    """A：CoInitialize → 建 PPT 代理 → 释放代理 → 再建 FSO。"""
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    app = win32com.client.Dispatch("PowerPoint.Application")
    log.append(f"  建 PPT 代理：{'OK' if app is not None else 'FAIL'}（Count={app.Presentations.Count}）")
    log.append(f"  释放前 FSO = {_try('Scripting.FileSystemObject')}")
    del app
    gc.collect()
    time.sleep(0.3)
    log.append(f"  释放代理后 FSO = {_try('Scripting.FileSystemObject')}")
    log.append(f"  释放代理后 PPT = {_try('PowerPoint.Application')}")


def case_pair(log):
    """B：CoInitialize →（不碰 PPT）→ CoInitialize+CoUninitialize → 再建 FSO。"""
    import pythoncom
    pythoncom.CoInitialize()
    pythoncom.CoInitialize()
    pythoncom.CoUninitialize()
    log.append(f"  纯一对 init/uninit 后 FSO = {_try('Scripting.FileSystemObject')}")
    log.append(f"  纯一对 init/uninit 后 PPT = {_try('PowerPoint.Application')}")


def case_live_ppt_proxy(log):
    """D：CoInitialize → **持有** PPT 代理（不释放）→ 一对 init/uninit → FSO。"""
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    app = win32com.client.Dispatch("PowerPoint.Application")
    _ = app.Presentations.Count
    pythoncom.CoInitialize()
    pythoncom.CoUninitialize()
    log.append(f"  持有 PPT 代理 + 一对 init/uninit 后 FSO = {_try('Scripting.FileSystemObject')}")
    log.append(f"  （代理类型=跨进程；app 仍被引用）")


def case_live_fso_proxy(log):
    """E：CoInitialize → **持有** 进程内 FSO 代理 → 一对 init/uninit → FSO。"""
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    fso = win32com.client.Dispatch("Scripting.FileSystemObject")
    _ = fso.Drives.Count
    pythoncom.CoInitialize()
    pythoncom.CoUninitialize()
    log.append(f"  持有 FSO 代理 + 一对 init/uninit 后 FSO = {_try('Scripting.FileSystemObject')}")
    log.append("  （代理类型=进程内；fso 仍被引用）")


def case_release_then_pair(log):
    """C：CoInitialize → 建并释放 PPT 代理 → CoInitialize+CoUninitialize → FSO。"""
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    app = win32com.client.Dispatch("PowerPoint.Application")
    _ = app.Presentations.Count
    del app
    gc.collect()
    pythoncom.CoInitialize()
    pythoncom.CoUninitialize()
    log.append(f"  释放代理 + 一对 init/uninit 后 FSO = {_try('Scripting.FileSystemObject')}")


def _run(name, body, out):
    def worker():
        log = []
        try:
            body(log)
        except Exception as exc:  # noqa: BLE001
            log.append(f"  [异常] {type(exc).__name__}: {exc}")
        finally:
            try:
                import pythoncom
                for _ in range(4):
                    pythoncom.CoUninitialize()
            except Exception:  # noqa: BLE001
                pass
        out[name] = log
    t = threading.Thread(target=worker)
    t.start()
    t.join()


def main() -> int:
    pre = pids()
    print(f"[闸门] 启动前 POWERPNT.EXE PIDs = {pre or '[]（无）'}")
    if pre:
        print("[ABORT] 已有 PowerPoint 进程，整条中止。")
        return 3

    out = {}
    _run("A. 只建+释放 PPT 代理", case_release, out)
    _run("B. 只做一对 init/uninit", case_pair, out)
    _run("C. 释放代理 + 一对 init/uninit", case_release_then_pair, out)
    _run("D. 持有 PPT 代理（不释放）+ 一对 init/uninit", case_live_ppt_proxy, out)
    _run("E. 持有 FSO 代理（不释放）+ 一对 init/uninit", case_live_fso_proxy, out)

    print()
    for name, log in out.items():
        print(f"== {name} ==")
        for ln in log:
            print(ln)
        print()

    def verdict(key, idx=0):
        return out[key][idx]

    print("== 结论 ==")
    print(f"  A 只释放 PPT 代理            → {verdict('A. 只建+释放 PPT 代理', 2)}")
    print(f"  B 只做 init/uninit 一对      → {verdict('B. 只做一对 init/uninit')}")
    print(f"  C 释放代理 + 一对 init/uninit → {verdict('C. 释放代理 + 一对 init/uninit')}")
    print(f"  D 持有 PPT 代理 + 一对        → {verdict('D. 持有 PPT 代理（不释放）+ 一对 init/uninit')}")
    print(f"  E 持有 FSO 代理 + 一对        → {verdict('E. 持有 FSO 代理（不释放）+ 一对 init/uninit')}")
    print("  判定：出现 FAIL(2147221008) 的组合即触发条件。")

    # 收尾：清掉我们建的实例
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        app = win32com.client.Dispatch("PowerPoint.Application")
        while app.Presentations.Count > 0:
            app.Presentations(1).Close()
        app.Quit()
    except Exception as exc:  # noqa: BLE001
        print(f"  Quit 失败：{type(exc).__name__}: {exc}")
    time.sleep(1.5)
    print(f"\n  收尾 PIDs = {pids() or '[]（无）'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
