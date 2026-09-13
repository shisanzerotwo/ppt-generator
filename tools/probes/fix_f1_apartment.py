"""F1 对照探针：`export_pages` 在**调用方已初始化 COM** 时还会不会拆人家的 apartment。

背景（TEST_REPORT §6 F1）
------------------------
`export_pages` 原先无条件 `CoInitialize()` + `finally: CoUninitialize()`。调用它的线程
若已经初始化过 COM（Flask worker、>50 页后台路径），那句 `CoUninitialize` 就把**调用方
的**计数一起减掉；调用方之后任何 COM 调用都可能报 `CO_E_NOTINITIALIZED`。

**为什么不靠"复现症状"判定**：测试 agent 实测它是**偶发**的（3 线程频次探针 1/3 死亡，
同线程连续两次调用又都正常）。靠抓偶发症状做前后对照不可靠。所以本探针换成一个
**确定性**判据 —— 给 `pythoncom.CoUninitialize` 装间谍，直接数：
    "调用方已初始化" 时，`export_pages` 该不该调它？
      修前：无条件调（计数 1）→ 拆的是别人的
      修后：S_FALSE 判定"不归我们管" → 计数 0
    "谁都没初始化" 时，`export_pages` 自己初始化了 → 理应自己收尾（计数 1）。

两组各跑在**独立子进程**里，保证 apartment 初态干净（同进程里 pywin32 的 Dispatch
会悄悄初始化 apartment 且不配对释放，把实验污染掉 —— 这正是本探针第一版误报 FIXED 的原因）。

安全三重闸门：① 启动前有 POWERPNT.EXE 即整条中止 ② 只用 tempfile 副本
③ 只 Quit 我们自己启动的实例。

跑法：
    .venv/Scripts/python.exe tools/probes/fix_f1_apartment.py            # 主流程（起两个子进程）
    .venv/Scripts/python.exe tools/probes/fix_f1_apartment.py <group>    # 内部用
"""

import ctypes
import os
import shutil
import subprocess
import sys
import tempfile
import time

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", ".."))
sys.path.insert(0, ROOT)

COINIT_APARTMENTTHREADED = 2
SRC_DECK = os.path.join(ROOT, "output", "b_multislide.pptx")


def powerpoint_running() -> bool:
    try:
        out = subprocess.run(["tasklist"], capture_output=True, text=True, timeout=30).stdout
    except Exception:  # noqa: BLE001
        return True
    return "POWERPNT.EXE" in out


def co_init() -> int:
    """调用 ctypes 直接拿 HRESULT（pywin32 的 CoInitializeEx 恒返回 None，读不到）。"""
    return ctypes.WinDLL("ole32").CoInitializeEx(None, COINIT_APARTMENTTHREADED)


def probe_com_usable() -> tuple:
    """当前线程还能不能正常建/用 COM 对象（apartment 被拆了就建不出来）。"""
    try:
        import win32com.client
        fso = win32com.client.Dispatch("Scripting.FileSystemObject")
        fso.GetSpecialFolder(2).Name
        return True, "可用"
    except Exception as exc:  # noqa: BLE001
        return False, f"{type(exc).__name__}: {exc}"


def run_group(group: str, deck: str, bg_dir: str) -> int:
    import pptx_io
    import pythoncom

    if group == "caller-init":
        hr = co_init()
        print(f"GROUP caller-init 调用方 CoInitializeEx → "
              f"{'S_OK' if hr == 0 else f'0x{hr & 0xFFFFFFFF:08X}（已初始化）'}")

    calls = {"n": 0}
    seen = {"we_initialized": None}
    real_uninit = pythoncom.CoUninitialize
    real_co_init = pptx_io._co_initialize

    def spy_uninit():
        calls["n"] += 1
        return real_uninit()

    def spy_co_init():
        val = real_co_init()
        seen["we_initialized"] = val
        return val

    pythoncom.CoUninitialize = spy_uninit
    pptx_io._co_initialize = spy_co_init
    try:
        shots = pptx_io.export_pages(deck, bg_dir, 1280)
        print(f"GROUP {group} export_pages 成功，{len(shots)} 页")
    except Exception as exc:  # noqa: BLE001
        print(f"GROUP {group} export_pages 抛异常：{type(exc).__name__}: {exc}")
    finally:
        pythoncom.CoUninitialize = real_uninit
        pptx_io._co_initialize = real_co_init

    ok, msg = probe_com_usable()
    print(f"GROUP {group} WE_INITIALIZED={seen['we_initialized']}")
    print(f"GROUP {group} COUNINIT_CALLS={calls['n']}")
    print(f"GROUP {group} 之后调用方还能用 COM：{msg}")
    return 0


def main() -> int:
    if len(sys.argv) >= 2 and sys.argv[1] in ("caller-init", "no-init"):
        return run_group(sys.argv[1], sys.argv[2], sys.argv[3])

    if powerpoint_running():
        print("SKIP: 启动前已有 POWERPNT.EXE 在运行 —— 按安全闸整条中止")
        return 0
    if not os.path.isfile(SRC_DECK):
        print(f"SKIP: 缺少素材 {SRC_DECK}")
        return 0

    tmp = tempfile.mkdtemp(prefix="f1probe_")
    deck = os.path.join(tmp, "deck.pptx")
    shutil.copyfile(SRC_DECK, deck)
    bg_dir = os.path.join(tmp, "bg")
    results = {}
    try:
        for group in ("no-init", "caller-init"):
            print(f"\n===== 子进程：{group} =====")
            out = subprocess.run(
                [sys.executable, os.path.abspath(__file__), group, deck, bg_dir],
                capture_output=True, text=True, encoding="utf-8", timeout=300)
            print(out.stdout.strip() or out.stderr.strip()[-800:])
            for line in out.stdout.splitlines():
                if "COUNINIT_CALLS=" in line:
                    results[group] = int(line.split("COUNINIT_CALLS=")[1].split()[0])
                if "WE_INITIALIZED=" in line:
                    results[group + "_we"] = line.split("WE_INITIALIZED=")[1].split()[0]

        print("\n" + "=" * 72)
        print("不变量：**不是自己加的计数就不要减** ——"
              "CoUninitialize 的调用次数必须 == (we_initialized ? 1 : 0)")
        bad = []
        for g in ("caller-init", "no-init"):
            n = results.get(g)
            we = results.get(g + "_we")
            want = 1 if we == "True" else 0
            flag = "OK" if n == want else "**违反**"
            print(f"  {g:<12} we_initialized={we:<5} CoUninitialize 调用={n} "
                  f"应为 {want}  {flag}")
            if n != want:
                bad.append(g)
        print(f"\n  caller-init 组：调用方**已经**初始化过 COM —— 修前这里恒为 1 次"
              f"（减的是别人加的计数）")
        if not bad:
            print("判定：**FIXED** —— 每个分支都只收自己开的计数")
        elif results.get("caller-init") == 1:
            print("判定：**REPRODUCED** —— 调用方已初始化时仍被 CoUninitialize（F1 成立）")
        else:
            print(f"判定：**异常** → {results}")
        print("=" * 72)
        return 0
    finally:
        try:
            import pythoncom
            import win32com.client
            pythoncom.CoInitialize()
            app = win32com.client.Dispatch("PowerPoint.Application")
            if app.Presentations.Count == 0:
                app.Quit()
        except Exception:  # noqa: BLE001
            pass
        shutil.rmtree(tmp, ignore_errors=True)
        time.sleep(3)
        print(f"清理：临时目录已删；残留 PowerPoint 进程 = {powerpoint_running()}")


if __name__ == "__main__":
    raise SystemExit(main())
