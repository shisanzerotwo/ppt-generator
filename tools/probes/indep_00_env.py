"""独立验证 · 环境勘察：跑 COM/浏览器相关探针前的闸门与可用性清单。

一条命令：./.venv/Scripts/python.exe tools/probes/indep_00_env.py
"""

import importlib.util
import os
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
SRC = os.path.join(ROOT, "output", "b_multislide.pptx")
BG = os.path.join(ROOT, "output", "spike", "m1")


def powerpnt_pids() -> list[str]:
    out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/NH"],
                         capture_output=True, text=True).stdout
    pids = []
    for line in out.splitlines():
        if "POWERPNT.EXE" in line.upper():
            parts = line.split()
            if len(parts) >= 2:
                pids.append(parts[1])
    return pids


def main() -> int:
    print("== 解释器 ==")
    print("  exe =", sys.executable)
    print("  cwd =", os.getcwd())
    print("  platform =", sys.platform)

    print("\n== PowerPoint 进程（COM 探针的前置闸门）==")
    pids = powerpnt_pids()
    print("  POWERPNT.EXE PIDs =", pids or "[]（无）")
    print("  → COM 探针", "必须整条中止" if pids else "可以运行")

    print("\n== 关键依赖 ==")
    for name in ("pptx", "PIL", "fontTools", "win32com.client", "pythoncom",
                 "playwright", "flask"):
        spec = importlib.util.find_spec(name)
        print(f"  {name:18} {'OK' if spec else 'MISSING'}")

    print("\n== 素材 ==")
    for label, p in (("b_multislide.pptx", SRC), ("bg/slide_1.png", os.path.join(BG, "slide_1.png"))):
        print(f"  {label:20} {'存在' if os.path.isfile(p) else '缺失'}  {p}")

    print("\n== 浏览器（真实渲染校验用）==")
    for p in (r"C:\Program Files\Google\Chrome\Application\chrome.exe",
              r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
              r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"):
        print(f"  {'有' if os.path.isfile(p) else '无'}  {p}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
