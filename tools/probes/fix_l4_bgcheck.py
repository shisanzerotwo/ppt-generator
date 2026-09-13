"""L4 对照探针：播放器缺底图时会安静地出黑帧，`hl.ready` 照样 resolve。

背景（AUDIT_REPORT §2 L4）：`hl_anim` 的底图 `<img>` 没有 `onerror` 处理，
缺图时浏览器只是不显示，`window.hl.ready` 仍然 resolve → 截图拿到全黑画面 →
**安静地喂进 ffmpeg**。调用侧（`shot_player` / `cli animate`）确实会校验，
但播放器本身对"自己能不能正确显示"是不设防的。

安全：本探针只用 Playwright 打开本地 HTML，**不碰 PowerPoint**（无需 COM 闸门）。

判定：
  修复前 → hl.ready resolve；帧几乎全黑（最大亮度低于阈值）
  修复后 → hl.ready **reject**（带缺图清单）；shot_player 抛 PptxError(IR_MISMATCH)

跑法：.venv/Scripts/python.exe tools/probes/fix_l4_bgcheck.py
"""

import os
import shutil
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", ".."))
sys.path.insert(0, ROOT)


def build_player_with_missing_bg(tmp: str) -> str:
    from PIL import Image
    import hl_anim
    import hl_layout
    from hl_layout import Rect, Unit

    out = os.path.join(tmp, "player_dir")
    os.makedirs(os.path.join(out, "bg"), exist_ok=True)
    Image.new("RGB", (64, 36), (10, 30, 60)).save(os.path.join(out, "bg", "slide_1.png"))
    # 第 2 页的底图**故意不创建**
    rect = Rect(0, 0, 12700 * 200, 12700 * 40, 0.0, 0.0, 400.0, 80.0)
    unit = Unit(order=0, page_index=0, text="要点", kind="body", rect=rect,
                lines=[rect], size_pt=18.0, align="LEFT", shape_id=1,
                shape_name="S", is_estimated=False, warnings=[])
    return hl_anim.build_player(out, ["bg/slide_1.png", "bg/slide_2_missing.png"],
                                [[unit], [unit]], canvas_width_px=640,
                                canvas_height_px=360)


def main() -> int:
    from pathlib import Path
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("SKIP: 未安装 playwright")
        return 0
    import shot

    tmp = tempfile.mkdtemp(prefix="l4probe_")
    try:
        player = build_player_with_missing_bg(tmp)
        url = Path(os.path.abspath(player)).as_uri()
        print(f"播放器：{player}（第 2 页底图**故意缺失**）")

        with sync_playwright() as p:
            try:
                browser = shot._launch_browser(p)
            except RuntimeError as exc:
                print(f"SKIP: 无可用浏览器（{exc}）")
                return 0
            try:
                page = browser.new_page(viewport={"width": 640, "height": 360})
                page.goto(url, wait_until="load")
                print("\n== 1) window.hl.ready 的行为 ==")
                resolved = True
                detail = ""
                try:
                    page.evaluate("window.hl.ready")
                except Exception as exc:  # noqa: BLE001
                    resolved = False
                    detail = str(exc).splitlines()[0]
                print(f"  ready {'resolve（**静默通过**）' if resolved else 'reject'}"
                      + (f" → {detail}" if detail else ""))

                page.wait_for_timeout(400)
                png = os.path.join(tmp, "frame.png")
                page.locator("#frame").screenshot(path=png)
                from PIL import Image
                with Image.open(png) as im:
                    rgb = im.convert("RGB")
                    px = rgb.load()
                    mx = max(max(px[x, y]) for y in range(0, rgb.height, 7)
                             for x in range(0, rgb.width, 7))
                print(f"\n== 2) 画面亮度（第 2 页缺图，应显示为黑）==")
                print(f"  最大通道值 = {mx}（阈值 40：低于即判黑帧）")
            finally:
                browser.close()

        print("\n== 3) shot_player 的反应 ==")
        import hl_anim
        import pptx_io
        try:
            hl_anim.shot_player(player, os.path.join(tmp, "frames"), [(1, 1)])
            print("  结果：**没有报错** —— 黑帧被安静地截出来")
        except pptx_io.PptxError as exc:
            print(f"  结果：PptxError code={exc.code}")
            print(f"        message={exc.message}")
        except Exception as exc:  # noqa: BLE001
            print(f"  结果：抛了别的异常 {type(exc).__name__}: {exc}")

        print("\n" + "-" * 68)
        if not resolved:
            print("判定：**FIXED** —— 缺图会让 hl.ready reject，不再静默出黑帧")
        else:
            print("判定：**REPRODUCED** —— hl.ready 静默 resolve，黑帧会流进 MP4（L4 成立）")
        print("-" * 68)
        return 0
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    raise SystemExit(main())
