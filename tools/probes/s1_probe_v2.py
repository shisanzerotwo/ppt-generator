"""S1 探针 v2：用**独立进程**当「用户」，把 `Close()` 的效果与 `CoUninitialize` 的效果分开。

v1（同进程模拟用户）的致命缺陷：
    探针自己用 `Presentations.Open()` 打开「用户稿」，再调 `export_pages()`。
    但 `export_pages` 内部会 `pythoncom.CoUninitialize()`（审计 F1）——**主进程的
    COM apartment 被拆掉**，PowerPoint 收回该客户端持有的文档引用，于是
    `Presentations.Count` 归 0。**这与 `pres.Close()` 完全无关**，v1 的「S1 复现」
    因此可能是假阳性。

v2 设计：把「用户」放到**另一个进程**里
    helper 子进程：CoInitialize → Dispatch → Open(deck) → 保持连接不动
    主进程      ：调 export_pages()（可能 Close / CoUninitialize / Quit）
    helper 再读 ：`Presentations.Count` —— 它的 apartment 全程有效，
                  所以这个数字才真实反映「用户的稿还在不在」

判定：
    COUNT=1 → 用户的稿仍在（Close 没有动它）
    COUNT=0 → 用户的稿被关掉（S1 成立，真凶待定：Close 或别的）

安全：三重闸门（启动前有 POWERPNT.EXE 就整条中止 / 只用临时副本 / 只 Quit 自己启动的实例）。

用法：
    .venv/Scripts/python.exe tools/probes/s1_probe_v2.py            # 主流程
    .venv/Scripts/python.exe tools/probes/s1_probe_v2.py helper F   # 内部用
"""

import os
import shutil
import subprocess
import sys
import tempfile
import time

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", ".."))
sys.path.insert(0, ROOT)


def powerpoint_running() -> bool:
    try:
        out = subprocess.run(["tasklist"], capture_output=True, text=True, timeout=30).stdout
    except Exception:
        return True
    return "POWERPNT.EXE" in out


def helper_main(deck: str) -> int:
    """子进程：扮演「用户」，打开稿后原地待命，全程不释放 COM。"""
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    app = win32com.client.Dispatch("PowerPoint.Application")
    pres = app.Presentations.Open(deck, ReadOnly=False, Untitled=False, WithWindow=False)
    print(f"HELPER_OPENED COUNT={app.Presentations.Count}", flush=True)

    sys.stdin.readline()                     # 等主进程的「检查」信号
    try:
        count = app.Presentations.Count
    except Exception as exc:                 # noqa: BLE001
        print(f"CHECK_ERROR {type(exc).__name__}: {exc}", flush=True)
        return 1
    alive = []
    for i in range(1, count + 1):
        try:
            alive.append(os.path.basename(app.Presentations(i).Name))
        except Exception:                    # noqa: BLE001
            alive.append("<unreadable>")
    print(f"CHECK COUNT={count} ALIVE={alive}", flush=True)
    try:
        pres.Name                             # 若被关闭，这里会抛
        print("HELPER_REF OK", flush=True)
    except Exception as exc:                  # noqa: BLE001
        print(f"HELPER_REF DEAD {type(exc).__name__}", flush=True)
    return 0


def main() -> int:
    if len(sys.argv) >= 3 and sys.argv[1] == "helper":
        return helper_main(sys.argv[2])

    if powerpoint_running():
        print("SKIP: 启动前已有 POWERPNT.EXE 在运行 —— 按安全闸整条中止")
        return 0

    src = os.path.join(ROOT, "output", "b_multislide.pptx")
    if not os.path.isfile(src):
        print(f"SKIP: 缺少素材 {src}")
        return 0

    tmp = tempfile.mkdtemp(prefix="s1v2_")
    deck = os.path.join(tmp, "user_deck.pptx")
    shutil.copyfile(src, deck)
    bg_dir = os.path.join(tmp, "bg")

    proc = None
    try:
        # ---- 1. 「用户」在独立进程里打开这份稿 ----
        proc = subprocess.Popen(
            [sys.executable, os.path.abspath(__file__), "helper", deck],
            stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, encoding="utf-8", bufsize=1,
        )
        first = proc.stdout.readline().strip()
        print(f"[1] helper(用户进程)：{first}")
        if "HELPER_OPENED" not in first:
            print("     helper 未能打开稿，终止实验")
            return 1
        time.sleep(1.0)

        # ---- 2. 主进程对**同一路径**调 export_pages ----
        import pptx_io
        try:
            shots = pptx_io.export_pages(deck, bg_dir, 1280)
            print(f"[2] export_pages 成功，导出 {len(shots)} 页底图")
        except Exception as exc:                                   # noqa: BLE001
            print(f"[2] export_pages 抛异常：{type(exc).__name__}: {exc}")

        # ---- 3. 让「用户」报告他的稿还在不在（他的 apartment 全程有效） ----
        proc.stdin.write("check\n")
        proc.stdin.flush()
        check = proc.stdout.readline().strip()
        print(f"[3] helper 复核：{check}")

        # ---- 4. 判定 ----
        print("-" * 70)
        if "COUNT=0" in check:
            print("判定：**S1 成立（用户的稿被关掉）** —— 且这次是独立进程观察到的，")
            print("      不是主进程 CoUninitialize 造成的假象。")
        elif "COUNT=" in check:
            print("判定：**S1 未成立** —— 用户的稿仍在，Close 没有动它。")
        else:
            print(f"判定：**无法判定**（helper 回复异常：{check}）")
        print(f"      PowerPoint 进程仍在：{powerpoint_running()}")
        print("-" * 70)
        return 0
    finally:
        if proc is not None:
            try:
                proc.stdin.write("exit\n")
                proc.stdin.flush()
                proc.wait(timeout=15)
            except Exception:                                      # noqa: BLE001
                proc.kill()
        # 清理我们自己启动的 PowerPoint（此时上面闸门保证了它是我们拉起的）
        try:
            import pythoncom
            import win32com.client
            pythoncom.CoInitialize()
            app = win32com.client.Dispatch("PowerPoint.Application")
            if app.Presentations.Count == 0:
                app.Quit()
        except Exception:                                          # noqa: BLE001
            pass
        shutil.rmtree(tmp, ignore_errors=True)
        time.sleep(3)
        print(f"清理：临时目录已删除；残留 PowerPoint 进程 = {powerpoint_running()}")


if __name__ == "__main__":
    raise SystemExit(main())
