"""试几种 COM 代理释放顺序，看哪种不触发 RPC_E_DISCONNECTED(0x80010108)。

pytest 默认开 faulthandler，会把 SEH 异常打成 "Windows fatal exception" 到 stderr。
本脚本同样打开 faulthandler，对比三种 finally 写法。
"""

import faulthandler
import os
import sys

faulthandler.enable()

import pythoncom  # noqa: E402
import win32com.client  # noqa: E402

SRC = os.path.abspath("output/b_multislide.pptx")


def variant(order: str):
    print(f"\n===== 变体 {order} =====", flush=True)
    app = pres = None
    pythoncom.CoInitialize()
    try:
        app = win32com.client.DispatchEx("PowerPoint.Application")
        pres = app.Presentations.Open(SRC, ReadOnly=True, Untitled=False, WithWindow=False)
        pres.Slides(1).Export(os.path.abspath("output/spike/com_order.png"), "PNG", 640, 360)
        pres.Close()
    finally:
        if order == "A_先松pres再Quit再松app":
            try:
                pres = None
            except Exception as e:
                print("  pres=None 抛:", type(e).__name__, e)
            try:
                app.Quit()
            except Exception as e:
                print("  Quit 抛:", type(e).__name__, e)
            try:
                app = None
            except Exception as e:
                print("  app=None 抛:", type(e).__name__, e)
        elif order == "B_Quit后CoUninit再松":
            try:
                app.Quit()
            except Exception as e:
                print("  Quit 抛:", type(e).__name__, e)
            pythoncom.CoUninitialize()
            try:
                pres = None
            except Exception as e:
                print("  pres=None 抛:", type(e).__name__, e)
            try:
                app = None
            except Exception as e:
                print("  app=None 抛:", type(e).__name__, e)
            print("  (CoUninitialize 已提前调用)")
            return
        elif order == "C_完全不Quit只松引用":
            try:
                pres = None
            except Exception as e:
                print("  pres=None 抛:", type(e).__name__, e)
            try:
                app = None
            except Exception as e:
                print("  app=None 抛:", type(e).__name__, e)
        pythoncom.CoUninitialize()
    print(f"  变体 {order} 结束")


if __name__ == "__main__":
    os.makedirs("output/spike", exist_ok=True)
    for name in ("A_先松pres再Quit再松app", "C_完全不Quit只松引用", "B_Quit后CoUninit再松"):
        variant(name)
    print("\n全部变体跑完（上面若出现 'Windows fatal exception' 即为噪音）")
