"""独立验证 ② 附录 · 归因实验：`CoInitialize()`/`CoUninitialize()` 是否真的能拆掉
调用线程**本已存在**的 COM apartment。

背景（indep_d1_disconnect.py 实测）：`pptx_io.export_pages()` 返回后，调用线程里
连一次全新的 `win32com.client.Dispatch(...)` 都会失败：
    com_error: (-2147221008, '尚未调用 CoInitialize。')
即 export_pages 的 `finally: pythoncom.CoUninitialize()` 把**调用方**的 apartment
一起卸掉了。本探针把归因做成三个对照实验（每个跑在**独立线程**里，apartment 互不干扰）：

  T1 基线：CoInitialize → 建对象 → CoUninitialize → 再建对象      （1 对 1，应失败）
  T2 实现者模式：CoInitialize → 建对象 → CoInitialize+CoUninitialize → 再建对象
                                                              （若失败 = 真 bug）
  T3 双初始化：CoInitialize ×2 → CoUninitialize ×1 → 再建对象   （看 pywin32 是否计数）

一条命令：./.venv/Scripts/python.exe tools/probes/indep_com_apartment.py
"""

import sys
import threading

import pythoncom
import win32com.client

PROGID = "Scripting.FileSystemObject"   # 常驻 Windows 的轻量 COM 对象


def _try_dispatch():
    try:
        win32com.client.Dispatch(PROGID)
        return True, "OK"
    except Exception as exc:  # noqa: BLE001
        return False, f"{type(exc).__name__}: {exc}"


def _co_init():
    try:
        return "ok", repr(pythoncom.CoInitialize())
    except Exception as exc:  # noqa: BLE001
        return "raised", f"{type(exc).__name__}: {exc}"


def _run_case(name, body, out):
    """在独立线程里跑一个实验（独立 apartment）。"""
    def worker():
        log = []
        try:
            body(log)
        except Exception as exc:  # noqa: BLE001
            log.append(f"  [异常] {type(exc).__name__}: {exc}")
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:  # noqa: BLE001
                pass
        out[name] = log

    t = threading.Thread(target=worker)
    t.start()
    t.join()


def case_t1(log):
    log.append(f"  CoInitialize #1 → {_co_init()}")
    log.append(f"  建对象 #1      → {_try_dispatch()}")
    pythoncom.CoUninitialize()
    log.append("  CoUninitialize #1")
    log.append(f"  建对象 #2      → {_try_dispatch()}")


def case_t2(log):
    log.append(f"  CoInitialize #1 → {_co_init()}")
    log.append(f"  建对象 #1      → {_try_dispatch()}")
    log.append(f"  CoInitialize #2 → {_co_init()}")
    pythoncom.CoUninitialize()
    log.append("  CoUninitialize #1")
    log.append(f"  建对象 #2      → {_try_dispatch()}   ← 实现者模式下的关键观测")


def case_t3(log):
    log.append(f"  CoInitialize #1 → {_co_init()}")
    log.append(f"  CoInitialize #2 → {_co_init()}")
    pythoncom.CoUninitialize()
    log.append("  CoUninitialize #1")
    log.append(f"  建对象          → {_try_dispatch()}")


def main() -> int:
    print("Python:", sys.version.split()[0])
    print(f"探针对象 ProgID = {PROGID}\n")
    out = {}
    _run_case("T1 基线（1 对 1）", case_t1, out)
    _run_case("T2 实现者模式（外部已 init，再多一对）", case_t2, out)
    _run_case("T3 双 init 单 uninit", case_t3, out)

    for name, log in out.items():
        print(f"== {name} ==")
        for ln in log:
            print(ln)
        print()

    t2 = out["T2 实现者模式（外部已 init，再多一对）"]
    verdict = next(("失败" if "False" in ln or "com_error" in ln else "成功")
                   for ln in t2 if "建对象 #2" in ln)
    print("== 结论 ==")
    print(f"  T2 里『外部先 init、再跑一遍 init/uninit』的最后一次建对象：{verdict}")
    if verdict == "失败":
        print("  → export_pages 的 CoUninitialize **会拆掉调用线程已有的 apartment**：")
        print("     ① 调用方持有的 PowerPoint 代理变 RPC_E_DISCONNECTED；")
        print("     ② 该线程后续任何 COM 调用都报 CO_E_NOTINITIALIZED，除非重新 init。")
        print("    这不止是'噪音'：契约 §3.4-6 的『无条件 CoUninitialize』写法在此场景是错的。")
    else:
        print("  → 未复现拆毁；indep_d1_disconnect.py 观测到的失败需另找归因。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
