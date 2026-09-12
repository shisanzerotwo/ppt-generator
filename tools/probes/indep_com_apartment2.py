"""独立验证 ② 附录 · 插桩定位：`export_pages()` 之后到底是什么失效了？

`indep_d1_disconnect.py` 观测到 export_pages 返回后，同线程里新建
`Dispatch("PowerPoint.Application")` 报 `CO_E_NOTINITIALIZED（尚未调用 CoInitialize）`；
但 `indep_com_apartment.py` 的 T2 证明"外部先 init + 再多一对 init/uninit"不会拆 apartment。
两者矛盾 → 必须插桩分辨：

  · 是整个线程 apartment 死了（连 Scripting.FileSystemObject 都建不出来），
  · 还是只有 PowerPoint 相关的 COM 路径死了（FSO 正常）。

⚠️ 前置闸门：启动前已有 POWERPNT.EXE 则整条中止。

一条命令：./.venv/Scripts/python.exe tools/probes/indep_com_apartment2.py
"""

import os
import subprocess
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

import pptx_io  # noqa: E402

SRC = os.path.join(ROOT, "output", "b_multislide.pptx")
OUT = os.path.join(ROOT, "output", "spike", "indep_apartment2")


def pids() -> list[str]:
    out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/NH"],
                         capture_output=True, text=True).stdout
    return sorted(ln.split()[1] for ln in out.splitlines()
                  if "POWERPNT.EXE" in ln.upper() and len(ln.split()) >= 2)


def probe(progid):
    import win32com.client
    try:
        obj = win32com.client.Dispatch(progid)
        return "OK"
    except Exception as exc:  # noqa: BLE001
        return f"{exc.args[0] if exc.args else type(exc).__name__}: {exc}"


def main() -> int:
    pre = pids()
    print(f"[闸门] 启动前 POWERPNT.EXE PIDs = {pre or '[]（无）'}")
    if pre:
        print("[ABORT] 已有 PowerPoint 进程，整条中止。")
        return 3

    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    print("\n=== 逐步插桩（同一线程）===")
    print(f"  1) 我方 CoInitialize 后：")
    print(f"       FSO  = {probe('Scripting.FileSystemObject')}")
    print(f"       PPT  = {probe('PowerPoint.Application')}")

    app = win32com.client.Dispatch("PowerPoint.Application")
    print(f"  2) 持有 PPT 代理（Count={app.Presentations.Count}）")

    os.makedirs(OUT, exist_ok=True)
    shots = pptx_io.export_pages(SRC, OUT, width=640)
    print(f"  3) export_pages 返回 {len(shots)} 张；PIDs = {pids() or '[]（无）'}")

    print(f"  4) export_pages 之后（同一线程、未重新 init）：")
    print(f"       旧 PPT 代理 = {_touch(app)}")
    print(f"       FSO         = {probe('Scripting.FileSystemObject')}")
    print(f"       新 PPT      = {probe('PowerPoint.Application')}")

    print(f"  5) 显式重新 CoInitialize 之后：")
    try:
        pythoncom.CoInitialize()
        print(f"       CoInitialize 返回 = ok")
    except Exception as exc:  # noqa: BLE001
        print(f"       CoInitialize 抛错：{exc}")
    print(f"       FSO         = {probe('Scripting.FileSystemObject')}")
    print(f"       新 PPT      = {probe('PowerPoint.Application')}")

    # 收尾
    try:
        a2 = win32com.client.Dispatch("PowerPoint.Application")
        while a2.Presentations.Count > 0:
            a2.Presentations(1).Close()
        a2.Quit()
        print("\n  已 Quit 我方实例")
    except Exception as exc:  # noqa: BLE001
        print(f"\n  Quit 失败：{type(exc).__name__}: {exc}")
    time.sleep(1.5)
    print(f"  收尾 PIDs = {pids() or '[]（无）'}")
    return 0


def _touch(app):
    try:
        return f"OK（Count={app.Presentations.Count}）"
    except Exception as exc:  # noqa: BLE001
        return f"{exc.args[0] if exc.args else type(exc).__name__}: {exc}"


if __name__ == "__main__":
    raise SystemExit(main())
