"""V8c 探针：判定 DispatchEx 是否真的开新进程（V8b 的更正版）。

V8b 的两个疑点要修：
  1. `app.HWND` 在 win32com 动态绑定下返回的是 bound method，必须 `app.HWND()` 调用；
  2. 判定"是不是同一实例"要用**决定性证据**：往用户实例里再 Add 一份稿，
     看 DispatchEx 实例看到的 Count 是否跟着变（跟着变 = 同一个实例）。

测完只清自己造的东西。
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


def main() -> int:
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    print(f"启动前 PIDs={_pids()}")

    user_app = win32com.client.Dispatch("PowerPoint.Application")
    p1 = user_app.Presentations.Add()
    p2 = user_app.Presentations.Add()
    print(f"用户实例: Count={user_app.Presentations.Count}")
    print(f"用户实例建好后 PIDs={_pids()}")

    ex_app = win32com.client.DispatchEx("PowerPoint.Application")
    print(f"DispatchEx: Count={ex_app.Presentations.Count}")
    print(f"DispatchEx 后 PIDs={_pids()}")

    print("\n== 决定性测试：用户实例再加 1 份稿，看 DispatchEx 是否跟着变 ==")
    p3 = user_app.Presentations.Add()
    print(f"用户 Count={user_app.Presentations.Count}，DispatchEx 看到的 Count={ex_app.Presentations.Count}")
    shared = user_app.Presentations.Count == ex_app.Presentations.Count
    print(f"→ DispatchEx 与用户实例{'是同一个' if shared else '是两个独立实例'}")

    # 逆向：DispatchEx 直接对已有实例调 Quit 会怎样？——不做（破坏性）。
    # 改为验证「Close 我们自己的那份」是否安全：Open 一份稿再 Close，看用户稿数不变。
    print("\n== Close 安全性：Open + Close 自己的稿，用户稿数应不变 ==")
    before = user_app.Presentations.Count
    pres = ex_app.Presentations.Open(os.path.abspath("output/b_multislide.pptx"),
                                     ReadOnly=True, Untitled=False, WithWindow=False)
    mid = user_app.Presentations.Count
    pres.Close()
    after = user_app.Presentations.Count
    print(f"用户实例稿数: Open前={before} Open后={mid} Close后={after}"
          f" → {'未受影响' if before == after else '被影响!'}")

    for p in (p1, p2, p3):
        try:
            p.Close()
        except Exception as exc:  # noqa: BLE001
            print(f"(清理提示) Close 失败：{exc}")
    try:
        user_app.Quit()
    except Exception as exc:  # noqa: BLE001
        print(f"(清理提示) Quit 失败：{exc}")
    pythoncom.CoUninitialize()
    print(f"清理后 PIDs={_pids()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
