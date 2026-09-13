"""F3 对照探针：复用 PowerPoint 实例**收益有多大**，以及**这个方案能不能落地**。

背景（TEST_REPORT §6 F3 + KNOWLEDGE.md 的复查）：`export_pages` 每次调用都自己
`DispatchEx` 起一个 PowerPoint、结束时 `Quit`。冷调用端到端 **2.5s/页**（10 页 25s），
而稳态逐页只要 0.13s。工作台实测上传一次 27 秒，也是这个原因。

本探针**不做任何修改**，量三件事：
  A 冷调用      ：调现成的 `export_pages`（起实例 + 导出 + Quit）
  B 常驻·首次   ：自己 DispatchEx 一次 → Open/Export/Close（**不 Quit**）
  C 常驻·复用   ：用**同一个** app 再来一遍 Open/Export/Close
  D 线程亲和性  ：在**另一个线程**（先自建 apartment）里用同一个 app

结论（实测，见输出）：
  · 复用确实快 —— C 相对 A 提速约 19×，每次省 ~24s；
  · 但 CLI 每次 import 都是**新进程**，只导一次 → 复用收益为 **0**；
  · 唯一的受益者是工作台（长驻进程、多次导入），而 app.py 的 pptx worker 是
    `threading.Thread(target=...)` **每请求新起线程** → D 段实测
    `RPC_E_WRONG_THREAD`（-2147417842，"已为另一线程整理的接口"）→
    **任务卡给的"模块级单例 + Lock"对真正的受益者直接不可用**。
  因此本轮**放弃 F3**（任务卡允许），理由与数据见 docs/FIX_REPORT.md。

安全三重闸门：① 启动前有 POWERPNT.EXE 即整条中止 ② 只用 tempfile 副本
③ 只 Quit 我们自己启动的实例，结束用 tasklist 复核残留。

跑法：.venv/Scripts/python.exe tools/probes/fix_f3_reuse.py
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

SRC_DECK = os.path.join(ROOT, "output", "b_multislide.pptx")


def powerpoint_running() -> bool:
    try:
        out = subprocess.run(["tasklist"], capture_output=True, text=True, timeout=30).stdout
    except Exception:  # noqa: BLE001
        return True
    return "POWERPNT.EXE" in out


def main() -> int:
    if powerpoint_running():
        print("SKIP: 启动前已有 POWERPNT.EXE 在运行 —— 按安全闸整条中止")
        return 0
    if not os.path.isfile(SRC_DECK):
        print(f"SKIP: 缺少素材 {SRC_DECK}")
        return 0

    tmp = tempfile.mkdtemp(prefix="f3probe_")
    deck = os.path.join(tmp, "deck.pptx")
    shutil.copyfile(SRC_DECK, deck)

    import pythoncom
    import win32com.client
    import pptx_io

    app = None
    try:
        # ---- A：现成的 export_pages（起实例 + 导出 + Quit）----
        print("== A 冷调用 export_pages ==")
        t0 = time.perf_counter()
        shots = pptx_io.export_pages(deck, os.path.join(tmp, "bgA"), width=1920)
        ta = time.perf_counter() - t0
        print(f"  {len(shots)} 页，{ta:.2f}s → {ta / len(shots):.2f}s/页")

        # ---- B/C：手写常驻实例 ----
        pythoncom.CoInitialize()
        os.makedirs(os.path.join(tmp, "bgB"), exist_ok=True)
        os.makedirs(os.path.join(tmp, "bgC"), exist_ok=True)
        print("\n== B 常驻·首次（DispatchEx + Open + Export×10 + Close，**不 Quit**）==")
        t0 = time.perf_counter()
        app = win32com.client.DispatchEx("PowerPoint.Application")
        t_start = time.perf_counter() - t0
        t0 = time.perf_counter()
        pres = app.Presentations.Open(deck, ReadOnly=True, Untitled=False, WithWindow=False)
        t_open = time.perf_counter() - t0
        t0 = time.perf_counter()
        for i in range(1, 11):
            pres.Slides(i).Export(os.path.join(tmp, "bgB", f"s{i}.png"), "PNG", 1920, 1080)
        t_exp = time.perf_counter() - t0
        pres.Close()
        tb = t_start + t_open + t_exp
        print(f"  启动 {t_start:.2f}s + Open {t_open:.2f}s + 导出 {t_exp:.2f}s = {tb:.2f}s"
              f" → {tb / 10:.2f}s/页")

        print("\n== C 常驻·复用同一个 app（Open + Export×10 + Close）==")
        t0 = time.perf_counter()
        pres = app.Presentations.Open(deck, ReadOnly=True, Untitled=False, WithWindow=False)
        t_open2 = time.perf_counter() - t0
        t0 = time.perf_counter()
        for i in range(1, 11):
            pres.Slides(i).Export(os.path.join(tmp, "bgC", f"s{i}.png"), "PNG", 1920, 1080)
        t_exp2 = time.perf_counter() - t0
        pres.Close()
        tc = t_open2 + t_exp2
        print(f"  Open {t_open2:.2f}s + 导出 {t_exp2:.2f}s = {tc:.2f}s → {tc / 10:.2f}s/页")

        print("\n" + "=" * 72)
        print(f"  A 冷调用（现状）        = {ta:6.2f}s  （{ta / 10:.2f}s/页）")
        print(f"  B 常驻·首次             = {tb:6.2f}s  （含一次性启动）")
        print(f"  C 常驻·第 2 次起（复用） = {tc:6.2f}s  （{tc / 10:.2f}s/页）")
        print(f"  → 复用相对冷调用提速 = {ta / tc:.1f}×，每次省 {ta - tc:.1f}s")
        print("  → 但**只有同进程里反复导出**才吃得到：CLI 每次 import 都是新进程，收益 0")

        # ---- D：跨线程复用（决定这个方案对"真正受益者"是否可行）----
        print("\n== D 线程亲和性：在**另一个线程**里用同一个 app ==")
        print("  背景：app.py 的 pptx worker 是 threading.Thread(target=...) —— "
              "**每次请求新起一个线程**；")
        print("        COM 的 STA 代理不跨线程，所以模块级单例在这里未必可用。")
        import threading
        box = {}

        def worker():
            # 关键：新线程**必须先自建 apartment**，否则任何 COM 调用都只会报
            # CO_E_NOTINITIALIZED，测不出"代理能不能跨线程用"这件事。
            pythoncom.CoInitialize()
            try:
                box["count"] = app.Presentations.Count
                try:
                    pres2 = app.Presentations.Open(deck, ReadOnly=True, Untitled=False,
                                                   WithWindow=False)
                    box["open"] = f"打开成功（{pres2.Slides.Count} 页）"
                    pres2.Close()
                except Exception as exc:  # noqa: BLE001
                    box["open"] = f"打开失败：{type(exc).__name__}: {exc}"
                box["ok"] = True
            except Exception as exc:  # noqa: BLE001
                box["ok"] = False
                box["err"] = f"{type(exc).__name__}: {exc}"
            finally:
                pythoncom.CoUninitialize()

        t = threading.Thread(target=worker)
        t.start()
        t.join(timeout=120)
        if box.get("ok"):
            print(f"  另一个线程（已自建 apartment）读 Count = {box['count']} → 代理可用")
            print(f"  再试着 Open：{box.get('open')}")
            print("  → 本环境下**跨线程复用是可行的**（pywin32 会做自动编组）；"
                  "所以单例方案在技术上不因线程而不可行。")
        else:
            print(f"  另一个线程（已自建 apartment）使用同一个 app → **失败**：{box.get('err')}")
            print("  → 结论：模块级单例对 Flask 的「每请求新线程」模型直接不可用")
        print("=" * 72)
        return 0
    finally:
        try:
            if app is not None and app.Presentations.Count == 0:
                app.Quit()
        except Exception:  # noqa: BLE001
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:  # noqa: BLE001
            pass
        shutil.rmtree(tmp, ignore_errors=True)
        time.sleep(3)
        print(f"清理：临时目录已删；残留 PowerPoint 进程 = {powerpoint_running()}")


if __name__ == "__main__":
    raise SystemExit(main())
