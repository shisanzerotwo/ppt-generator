"""独立验证 ⑥ 附录 · 逐页计时：把"冷启动 vs 单页导出"的成本拆开。

已知（indep_m1_com.py）：两次连续 `export_pages(1920)` 各要 ~27s —— 因为每次调用
自己起 PowerPoint、自己 Quit，"第二次"其实是又一次冷启动。
本探针在一个实例内逐页计时，回答：冷启动/打开各占多少、单页导出到底多快。

⚠️ 前置闸门：启动前已有 POWERPNT.EXE 则整条中止。

一条命令：./.venv/Scripts/python.exe tools/probes/indep_m1_com_page.py
"""

import os
import statistics
import subprocess
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

SRC = os.path.join(ROOT, "output", "b_multislide.pptx")
OUT = os.path.join(ROOT, "output", "spike", "indep_m1_page")


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
    t_begin = time.time()
    app = win32com.client.DispatchEx("PowerPoint.Application")
    t_launch = time.time() - t_begin
    pre_count = app.Presentations.Count
    t_open0 = time.time()
    pres = app.Presentations.Open(os.path.abspath(SRC), ReadOnly=True,
                                 Untitled=False, WithWindow=False)
    t_open = time.time() - t_open0
    w_emu = int(pres.PageSetup.SlideWidth)
    h_emu = int(pres.PageSetup.SlideHeight)
    height = round(1920 * h_emu / w_emu)
    total = int(pres.Slides.Count)
    per_page = []
    for i in range(1, total + 1):
        t0 = time.time()
        pres.Slides(i).Export(os.path.abspath(os.path.join(OUT, f"slide_{i}.png")),
                              "PNG", 1920, height)
        per_page.append(time.time() - t0)
    t_total = time.time() - t_begin
    pres.Close()
    print("\n== 成本拆分（同一实例内）==")
    print(f"  PowerPoint 启动（DispatchEx→可用） = {t_launch:.2f}s")
    print(f"  Presentations.Open                  = {t_open:.2f}s")
    print(f"  逐页 Export：{['%.3f' % t for t in per_page]}")
    print(f"  单页中位 = {statistics.median(per_page):.3f}s，最大 = {max(per_page):.3f}s，"
          f"合计 = {sum(per_page):.2f}s")
    print(f"  一次调用端到端 = {t_total:.2f}s（{t_total / total:.2f}s/页）")
    print(f"  → 冷启动+打开（{t_launch + t_open:.2f}s）占端到端的 "
          f"{(t_launch + t_open) / t_total * 100:.0f}%")

    try:
        if pre_count == 0:
            app.Quit()
    except Exception:  # noqa: BLE001
        pass
    pythoncom.CoUninitialize()
    time.sleep(1.5)
    print(f"\n  收尾 PIDs = {pids() or '[]（无）'}")
    print("\n== 判定 ==")
    print(f"  契约验收『≤1.5s/页』：单页稳态中位 {statistics.median(per_page):.3f}s → "
          f"{'达标' if statistics.median(per_page) <= 1.5 else '不达标'}；")
    print(f"  但端到端 {t_total / total:.2f}s/页（每次 CLI 调用都重付冷启动）→ 用户体感不达标。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
