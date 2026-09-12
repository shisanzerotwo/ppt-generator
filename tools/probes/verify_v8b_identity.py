"""V8b 探针：Dispatch vs DispatchEx 到底是不是两个进程？

V8 主探针里 DispatchEx 的实例 Open 前 `Presentations.Count == 1`（等于用户实例
的稿数）——这暗示它可能**附着**到了用户实例。若属实，契约 §3.4-5 的
「pre_count == 0 才 Quit」守卫前提（DispatchEx 必开新实例）就是错的，
必须实测出真实行为再决定实现怎么写。

本探针只读性质地测身份：比较两个实例的 HWND 与宿主进程 PID，
以及 DispatchEx 前后 POWERPNT.EXE 进程数。测完立即清理，不 Quit 用户实例。
"""

import os
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))


def _pids():
    out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/FO", "CSV", "/NH"],
                         capture_output=True, text=True).stdout
    pids = []
    for line in out.splitlines():
        parts = [p.strip('"') for p in line.split('","')]
        if len(parts) >= 2 and parts[0].upper().startswith("POWERPNT"):
            pids.append(parts[1])
    return pids


def _pid_of_window(hwnd):
    import win32process
    try:
        return win32process.GetWindowThreadProcessId(hwnd)[1]
    except Exception as exc:  # noqa: BLE001
        return f"(取不到: {exc})"


def main() -> int:
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    print(f"启动前 POWERPNT.EXE PIDs = {_pids()}")

    user_app = win32com.client.Dispatch("PowerPoint.Application")
    user_pres = user_app.Presentations.Add()
    user_pres.Slides.Add(1, 1)
    user_hwnd = user_app.HWND
    print(f"用户实例: HWND={user_hwnd} PID={_pid_of_window(user_hwnd)} "
          f"Count={user_app.Presentations.Count}")
    print(f"用户实例建好后 PIDs = {_pids()}")

    other_app = win32com.client.DispatchEx("PowerPoint.Application")
    other_hwnd = other_app.HWND
    print(f"DispatchEx 实例: HWND={other_hwnd} PID={_pid_of_window(other_hwnd)} "
          f"Count={other_app.Presentations.Count}")
    print(f"DispatchEx 后 PIDs = {_pids()}")

    print(f"\n两个实例 HWND 相同? {user_hwnd == other_hwnd}")
    print(f"两个实例 PID  相同? {_pid_of_window(user_hwnd) == _pid_of_window(other_hwnd)}")
    print(f"两实例 Presentations 是同一个集合? "
          f"{user_app.Presentations.Count == other_app.Presentations.Count}")

    # 反向再测一次：如果 Dispatch 在已有实例时也能附着，说明进程是共享的
    third = win32com.client.Dispatch("PowerPoint.Application")
    print(f"再 Dispatch 一次: HWND={third.HWND} 与用户相同? {third.HWND == user_hwnd}")

    # 清理：不 Quit，避免影响用户实例；只关掉我们 Add 的那份稿
    try:
        user_pres.Close()
    except Exception as exc:  # noqa: BLE001
        print(f"(清理提示) user_pres.Close() 失败：{exc}")
    try:
        user_app.Quit()
    except Exception as exc:  # noqa: BLE001
        print(f"(清理提示) user_app.Quit() 失败：{exc}")
    pythoncom.CoUninitialize()
    print(f"清理后 PIDs = {_pids()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
