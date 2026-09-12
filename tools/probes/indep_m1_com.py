"""独立验证 ⑥ 附录 · COM 导出的页数/尺寸/耗时口径（真机）。

实现者报：10 张 1920×1080；首轮 27.03s（含 PowerPoint 冷启动），稳定后中位 0.146s/页，
最大 0.340s/页。本探针独立量：冷启动一轮 + 紧跟一轮（热），逐文件核对尺寸，
并检查"我们自己起的 PowerPoint 是否被正常 Quit"（与 D1 的残余缺口场景互补）。

⚠️ 前置闸门：启动前已有 POWERPNT.EXE 则整条中止。

一条命令：./.venv/Scripts/python.exe tools/probes/indep_m1_com.py
"""

import os
import subprocess
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

import pptx_io  # noqa: E402

SRC = os.path.join(ROOT, "output", "b_multislide.pptx")
OUT = os.path.join(ROOT, "output", "spike", "indep_m1")


def pids() -> list[str]:
    out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/NH"],
                         capture_output=True, text=True).stdout
    return sorted(ln.split()[1] for ln in out.splitlines()
                  if "POWERPNT.EXE" in ln.upper() and len(ln.split()) >= 2)


def run_once(label, width):
    d = os.path.join(OUT, f"w{width}_{label}")
    t0 = time.time()
    shots = pptx_io.export_pages(SRC, d, width=width)
    dt = time.time() - t0
    from PIL import Image
    sizes = set()
    for s in shots:
        with Image.open(s) as im:
            sizes.add(im.size)
    names = [os.path.basename(s) for s in shots]
    abs_ok = all(os.path.isabs(s) for s in shots)
    print(f"  [{label}] width={width}：{len(shots)} 张，{dt:.2f}s"
          f"（{dt / max(len(shots), 1):.3f}s/页）尺寸集合={sizes} 全绝对路径={abs_ok}")
    print(f"           文件名={names}")
    return {"n": len(shots), "dt": dt, "sizes": sizes, "abs": abs_ok}


def main() -> int:
    pre = pids()
    print(f"[闸门] 启动前 POWERPNT.EXE PIDs = {pre or '[]（无）'}")
    if pre:
        print("[ABORT] 已有 PowerPoint 进程，整条中止（COM 部分不做）。")
        print("        可先关闭 PowerPoint 再重跑本探针。")
        return 3

    print("\n== COM 导出 ==")
    cold = run_once("cold", 1920)
    warm = run_once("warm", 1920)
    time.sleep(1.0)
    after = pids()
    print(f"\n  两次导出后 PIDs = {after or '[]（无）'}")
    print(f"  → 我们自己起的 PowerPoint 被 Quit：{'是' if not after else '否（仍留进程）'}")

    print("\n== 与实现者口径对照 ==")
    print(f"  实现者：10 张 1920×1080；首轮 27.03s（冷启动）；稳定后中位 0.146s/页")
    print(f"  我实测：{cold['n']} 张 {cold['sizes']}；冷 {cold['dt']:.2f}s、热 {warm['dt']:.2f}s"
          f"（热态 {warm['dt'] / max(warm['n'], 1):.3f}s/页）")
    ok = (cold["n"] == 10 and cold["sizes"] == {(1920, 1080)}
          and cold["abs"] and not after)
    print(f"\n== 判定：{'通过（页数/尺寸/绝对路径/Quit 均符）' if ok else '需复核'} ==")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
