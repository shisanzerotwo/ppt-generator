"""V8 探针：DispatchEx 是否真的隔离用户正在用的 PowerPoint 实例。

风险：若用 Dispatch("PowerPoint.Application") 会附着到用户已开的实例，
随后的 Quit() 会关掉用户**没保存**的稿子。这是本次唯一有「破坏用户数据」
后果的点，必须先实测再写 pptx_io.export_pages。

做法（自动化模拟「用户开着未保存的稿」）：
  1. 先用 Dispatch 建一个「用户实例」，Add() 一份未保存的空稿；
  2. 用 DispatchEx 走一遍 export_pages 的原型逻辑（Open/Export/Close/Quit）；
  3. 断言用户实例与那份未保存的稿仍在、内容没变；
  4. 清理探针自己造的东西（不动用户实例里那份稿）。

跑法：.venv/Scripts/python.exe tools/probes/verify_v8_com.py
"""

import os
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

SRC = os.path.abspath("output/b_multislide.pptx")
OUT = os.path.abspath("output/spike/v8")


def _pptx_process_count() -> int:
    out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE"],
                         capture_output=True, text=True).stdout
    return out.count("POWERPNT.EXE")


def main() -> int:
    import win32com.client
    import pythoncom

    if not os.path.isfile(SRC):
        print(f"[FAIL] 素材不存在：{SRC}")
        return 1

    pythoncom.CoInitialize()
    os.makedirs(OUT, exist_ok=True)

    print("== 阶段 0：模拟用户实例（含一份未保存的稿） ==")
    user_app = win32com.client.Dispatch("PowerPoint.Application")
    user_pres = user_app.Presentations.Add()
    user_pres.Slides.Add(1, 1)  # ppLayoutTitle
    user_name = user_pres.Name
    user_count_before = user_app.Presentations.Count
    print(f"用户实例已开：未保存稿名={user_name!r}，Presentations.Count={user_count_before}")
    print(f"PowerPoint 进程数={_pptx_process_count()}")

    rc = 0
    try:
        print("\n== 阶段 1：DispatchEx 走 export 原型（Open/Export/Close/Quit） ==")
        app = win32com.client.DispatchEx("PowerPoint.Application")
        pre_count = app.Presentations.Count
        print(f"新实例 Presentations.Count(Open 前) = {pre_count}")

        pres = app.Presentations.Open(SRC, ReadOnly=True, Untitled=False, WithWindow=False)
        print(f"WithWindow=False 打开成功：Slides={pres.Slides.Count}")

        # 画布比例 → 目标像素高
        w_emu, h_emu = 12191695, 6858000
        width_px = 1920
        height_px = round(width_px * h_emu / w_emu)
        png = os.path.join(OUT, "v8_slide_1.png")
        pres.Slides(1).Export(png, "PNG", width_px, height_px)
        ok_export = os.path.isfile(png) and os.path.getsize(png) > 0
        print(f"WithWindow=False 下 Export 可用：{ok_export}，"
              f"文件大小={os.path.getsize(png) if os.path.isfile(png) else 0} 字节，"
              f"目标={width_px}x{height_px}")

        # 资源收尾（照契约 §3.4-5）
        pres.Close()
        print("pres.Close() 完成")
        if pre_count == 0:
            app.Quit()
            print("pre_count==0 → app.Quit() 完成（该实例是我们自己建的）")
    except Exception as exc:  # noqa: BLE001
        print(f"[FAIL] DispatchEx 路径异常：{type(exc).__name__}: {exc}")
        rc = 1
    finally:
        try:
            app = None
        except Exception:
            pass

    print("\n== 阶段 2：确认用户实例与未保存的稿仍在 ==")
    try:
        still_open = user_app.Presentations.Count
        names = [user_app.Presentations(i + 1).Name for i in range(still_open)]
        alive = user_name in names
        print(f"用户实例 Presentations.Count={still_open}（期望 {user_count_before}）")
        print(f"稿子清单={names}")
        print(f"未保存的稿 {user_name!r} 仍在 → {alive}")
        if not alive or still_open != user_count_before:
            print("[FAIL] DispatchEx 影响到了用户实例！")
            rc = 1
        else:
            print("[PASS] DispatchEx 未触碰用户实例（V8 通过）")
    except Exception as exc:  # noqa: BLE001
        print(f"[FAIL] 用户实例已不可访问 → {type(exc).__name__}: {exc}")
        rc = 1
    finally:
        # 清理探针自己造的那份未保存稿；不保存退出
        try:
            user_pres.Close()
        except Exception as exc:  # noqa: BLE001
            print(f"(清理提示) user_pres.Close() 失败：{exc}")
        try:
            user_app.Quit()
        except Exception as exc:  # noqa: BLE001
            print(f"(清理提示) user_app.Quit() 失败：{exc}")

    pythoncom.CoUninitialize()
    print(f"\n退出码={rc}")
    return rc


if __name__ == "__main__":
    raise SystemExit(main())
