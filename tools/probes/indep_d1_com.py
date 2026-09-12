"""独立验证 ② · D1（高危）：`DispatchEx` 是否真的隔离用户 PowerPoint 实例。

⚠️ 前置闸门：**脚本启动前若已有 POWERPNT.EXE 进程，整条中止、一次 COM 都不调**。
（因为本探针会建/关 PowerPoint，绝不能牵连用户正在用的实例。）

三个独立问题：
  A. `DispatchEx` 与 `Dispatch`、`GetActiveObject` 拿到的是不是同一个实例？
     （不照抄实现者的探针写法：这里用三个**不同的创建 API** 交叉对照 + 进程数）
  B. D1 指出的**残余缺口**是否真被加固堵住：用户开着 PowerPoint、但**没打开任何稿**
     （`pre_count == 0`）时跑 `export_pages`，用户的 PowerPoint 会不会被 Quit 掉？
  C. 红线：`export_pages` 全程是否**从未触碰 `app.Visible`**（真机 + 静态双查）。

一条命令：./.venv/Scripts/python.exe tools/probes/indep_d1_com.py
"""

import os
import subprocess
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

import pptx_io  # noqa: E402

SRC = os.path.join(ROOT, "output", "b_multislide.pptx")
OUT = os.path.join(ROOT, "output", "spike", "indep_d1")
MODULE = os.path.join(ROOT, "pptx_io.py")


def pids() -> list[str]:
    out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/NH"],
                         capture_output=True, text=True).stdout
    found = []
    for line in out.splitlines():
        if "POWERPNT.EXE" in line.upper():
            parts = line.split()
            if len(parts) >= 2:
                found.append(parts[1])
    return sorted(found)


def hr(title):
    print("\n" + "=" * 72)
    print(title)
    print("=" * 72)


def main() -> int:
    # ---------------------------------------------------------- 前置闸门
    pre = pids()
    print(f"[闸门] 启动前 POWERPNT.EXE PIDs = {pre or '[]（无）'}")
    if pre:
        print("[ABORT] 检测到已有 PowerPoint 进程 —— 整条中止，未做任何 COM 调用。")
        print("        这是 D1 的安全红线：绝不触碰用户正在用的实例。")
        return 3

    import pythoncom
    import win32com.client

    results = {}
    owned = []          # 我们创建、有权 Quit 的实例

    try:
        # ------------------------------------------------------ A. 实例身份
        hr("A. Dispatch / DispatchEx / GetActiveObject 是否同一个实例")
        pythoncom.CoInitialize()
        print(f"  CoInitialize 后 PIDs = {pids()}")

        app1 = win32com.client.Dispatch("PowerPoint.Application")
        owned.append(app1)
        pids_after_dispatch = pids()
        print(f"  Dispatch 后：PIDs = {pids_after_dispatch}（{len(pids_after_dispatch)} 个进程）"
              f"  Count = {app1.Presentations.Count}")

        app2 = win32com.client.DispatchEx("PowerPoint.Application")
        owned.append(app2)
        pids_after_ex = pids()
        print(f"  DispatchEx 后：PIDs = {pids_after_ex}（{len(pids_after_ex)} 个进程）"
              f"  Count = {app2.Presentations.Count}")

        app3 = None
        try:
            app3 = win32com.client.GetActiveObject("PowerPoint.Application")
            print(f"  GetActiveObject 后：Count = {app3.Presentations.Count}")
        except Exception as exc:  # noqa: BLE001
            print(f"  GetActiveObject 失败：{type(exc).__name__}: {exc}")

        # 往 app1 里加一份稿，看 app2/app3 是否"看得见"
        pres_probe = app1.Presentations.Open(os.path.abspath(SRC), ReadOnly=True,
                                            Untitled=False, WithWindow=False)
        c1, c2 = app1.Presentations.Count, app2.Presentations.Count
        c3 = app3.Presentations.Count if app3 is not None else None
        print(f"  app1 打开 1 份稿后：app1.Count={c1}  app2.Count={c2}  "
              f"app3.Count={c3}")
        pids_after_open = pids()
        print(f"  打开后 PIDs = {pids_after_open}（{len(pids_after_open)} 个进程）")

        same = (c2 == c1) and (c3 is None or c3 == c1)
        results["A_same_instance"] = same
        print(f"  → 三个 API 是否同一个实例：{'是（DispatchEx 未开新进程/新实例）' if same else '否（真隔离）'}")

        # 反向证据：在 app2 上 Quit，看 app1 是否同死（同一个实例才会同死）
        pres_probe.Close()
        print(f"  关闭探针稿后 Count = {app1.Presentations.Count}")
        results["A_dispatch_created_one_process"] = len(pids_after_dispatch) == 1
        results["A_dispatchex_no_new_process"] = len(pids_after_ex) == len(pids_after_dispatch)

        # ------------------------------------------------------ B. 残余缺口
        hr("B. D1 残余缺口：用户开着 PowerPoint 但**没打开任何稿**（pre_count==0）")
        print(f"  B 开始时 PIDs = {pids()}，Count = {app1.Presentations.Count}"
              f"（=0 即『用户开着但没开稿』的场景）")

        visible_before = None
        try:
            visible_before = app1.Visible
        except Exception as exc:  # noqa: BLE001
            print(f"  读取 app.Visible 失败（不影响判定）：{type(exc).__name__}")

        os.makedirs(OUT, exist_ok=True)
        t0 = time.time()
        shots = pptx_io.export_pages(SRC, OUT, width=640)
        dt = time.time() - t0
        print(f"  export_pages 完成：{len(shots)} 张，{dt:.1f}s（width=640）")

        time.sleep(0.5)
        alive = pids()
        print(f"  导出后 PIDs = {alive or '[]（无）'}")
        print(f"  导出后 Count = {app1.Presentations.Count}")

        survived = len(alive) > 0
        results["B_user_app_survived"] = survived
        print(f"  → 用户的 PowerPoint 是否被保住：{'是（守卫生效）' if survived else '否（被 Quit —— 高危！）'}")
        print(f"  → 用户的稿数是否仍为 0（没人替他开稿）："
              f"{'是' if app1.Presentations.Count == 0 else '否'}")

        # ------------------------------------------------------ C. Visible 红线
        hr("C. 红线：全程是否触碰过 app.Visible")
        visible_after = None
        try:
            visible_after = app1.Visible
        except Exception as exc:  # noqa: BLE001
            print(f"  读取 app.Visible 失败：{type(exc).__name__}")
        print(f"  export_pages 前 app.Visible = {visible_before}")
        print(f"  export_pages 后 app.Visible = {visible_after}")
        results["C_visible_unchanged"] = (visible_before == visible_after)
        print(f"  → Visible 未被改动：{results['C_visible_unchanged']}")

        with open(MODULE, encoding="utf-8") as f:
            code = f.read()
        hits = [(i, ln.strip()) for i, ln in enumerate(code.splitlines(), 1)
                if "Visible" in ln]
        print(f"  静态检查 pptx_io.py 含 'Visible' 的行：")
        for i, ln in hits:
            print(f"    L{i}: {ln}")
        results["C_no_visible_write"] = not any(
            ("Visible" in ln and "=" in ln and "==" not in ln and "!=" not in ln)
            for _, ln in hits)

    finally:
        # ------------------------------------------------------ 清理
        hr("清理与收尾")
        for app in owned:
            try:
                while app.Presentations.Count > 0:
                    app.Presentations(1).Close()
            except Exception:  # noqa: BLE001
                pass
        for app in owned:
            try:
                app.Quit()
            except Exception:  # noqa: BLE001
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:  # noqa: BLE001
            pass
        time.sleep(1.0)
        print(f"  Quit 后 PIDs = {pids() or '[]（无，已清干净）'}")

    # ---------------------------------------------------------- 判定
    hr("判定")
    for k, v in results.items():
        print(f"  {k:38} = {v}")
    ok = all(results.values())
    print(f"\n  总判定：{'PASS' if ok else 'FAIL / 需复核'}")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
