"""S1 实测探针：`export_pages()` 的 `pres.Close()` 会不会关掉「用户」打开着的同一份稿。

审计报告 S1（严重·条件性）的机理：
    PowerPoint 对**同一路径**的稿返回同一个 Presentation 实例。`export_pages`
    在 finally 里无条件 `pres.Close()`，于是当用户自己正开着这份稿时，被关掉的
    可能是**用户的稿**（未保存的修改即丢失）。`Quit()` 有两道守卫，`Close()` 一道没有。

安全设计（三重）：
  1. **闸门**：探针启动前若已有 POWERPNT.EXE 进程 → 整条中止，一次 COM 都不调。
  2. **只用临时副本**：复制到 tempfile 目录操作，绝不触碰任何真实文件。
  3. **只清理自己启动的实例**：探针自己启动 PowerPoint，结束时只 Quit 这一个。

判定：调用前后各重新附着一次取 `Presentations.Count`（不能复用代理——审计 F1 已证实
`export_pages` 内部的 CoUninitialize 会拆掉调用方 apartment，令旧代理失效）。

用法（Windows 侧）：
    .venv/Scripts/python.exe tools/probes/s1_close_guard_probe.py
"""

import os
import shutil
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", ".."))
sys.path.insert(0, ROOT)


def powerpoint_running() -> bool:
    """闸门用：此刻机器上是否已有 PowerPoint 进程（不区分是不是我们的）。"""
    try:
        out = subprocess.run(["tasklist"], capture_output=True, text=True, timeout=30).stdout
    except Exception:
        return True  # 探测不了就当作"有"，偏保守
    return "POWERPNT.EXE" in out


def main() -> int:
    if powerpoint_running():
        print("SKIP: 探针启动前已有 POWERPNT.EXE 在运行 —— 按安全闸整条中止，"
              "一次 COM 都不调（请先手动关闭 PowerPoint 再重跑）")
        return 0

    import pythoncom
    import win32com.client

    src = os.path.join(ROOT, "output", "b_multislide.pptx")
    if not os.path.isfile(src):
        print(f"SKIP: 缺少素材 {src}")
        return 0

    tmp = tempfile.mkdtemp(prefix="s1probe_")
    deck = os.path.join(tmp, "user_deck.pptx")
    shutil.copyfile(src, deck)          # 只用副本
    bg_dir = os.path.join(tmp, "bg")

    pythoncom.CoInitialize()
    app = None
    launched_by_us = False
    try:
        # ---- 1. 模拟「用户」打开这份稿 ----
        app = win32com.client.Dispatch("PowerPoint.Application")
        launched_by_us = True
        user_pres = app.Presentations.Open(deck, ReadOnly=False, Untitled=False,
                                           WithWindow=False)
        user_pres_name = user_pres.Name
        count_before = app.Presentations.Count
        print(f"[1] 模拟用户已打开：{os.path.basename(deck)}")
        print(f"    Presentations.Count = {count_before}")

        # ---- 2. 对**同一路径**调用 export_pages ----
        # 此刻机器上已有 POWERPNT.EXE → export_pages 内部的 had_powerpoint=True
        # → Quit 守卫会挡住 Quit。于是本实验隔离出的正是 Close() 的**单独**效果。
        import pptx_io
        try:
            shots = pptx_io.export_pages(deck, bg_dir, 1280)
            print(f"[2] export_pages 成功，导出 {len(shots)} 页底图")
        except Exception as exc:                                   # noqa: BLE001
            print(f"[2] export_pages 抛异常：{type(exc).__name__}: {exc}")

        # ---- 3. 重新附着取最新状态（不能复用旧代理，见 F1） ----
        # F1 实证：export_pages 内部的 CoUninitialize 已拆掉本线程 apartment，
        # 不重新 CoInitialize 就直接 Dispatch 会抛 0x800401F0（尚未调用 CoInitialize）。
        # 本次探针实际撞到过这个错误 —— 这是 F1 的第三方独立复现。
        pythoncom.CoInitialize()
        app2 = win32com.client.Dispatch("PowerPoint.Application")
        count_after = app2.Presentations.Count
        alive_names = []
        for i in range(1, count_after + 1):
            try:
                alive_names.append(app2.Presentations(i).Name)
            except Exception:                                      # noqa: BLE001
                alive_names.append("<unreadable>")
        print(f"[3] 调用后 Presentations.Count = {count_after}，仍打开的稿：{alive_names}")

        # ---- 4. 判定 ----
        print("-" * 68)
        if count_after < count_before:
            print("判定：**S1 成立（复现）** —— 用户的稿被 export_pages 关掉了。")
            print(f"      用户原稿 '{os.path.basename(user_pres_name)}' 已不在 PowerPoint 中；"
                  f"若用户当时有未保存修改，即已丢失。")
        else:
            print("判定：**S1 未复现** —— 用户的稿仍在（Close 未作用于它）。")
        print(f"      PowerPoint 进程仍在（Quit 守卫生效）：{powerpoint_running()}")
        print("-" * 68)
        return 0
    finally:
        # 只清理我们自己启动的实例：此时进程是我们拉起来的，且用户稿已不在其中
        if launched_by_us:
            try:
                win32com.client.Dispatch("PowerPoint.Application").Quit()
            except Exception:                                      # noqa: BLE001
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:                                          # noqa: BLE001
            pass
        shutil.rmtree(tmp, ignore_errors=True)
        print(f"清理：临时目录已删除；残留 PowerPoint 进程 = {powerpoint_running()}")


if __name__ == "__main__":
    raise SystemExit(main())
