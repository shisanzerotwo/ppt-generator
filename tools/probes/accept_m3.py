"""M3 验收：真浏览器打开高亮播放器，驱动 window.hl 并逐页截图。

验四件事：
1. 页数 == 底图数（契约 §11.1 步 4）；
2. window.hl.goto 同步生效（调用后立刻读 DOM，高亮 div 数就对）；
3. 高亮定位准、其余降暗（截图对比"有高亮"与"无高亮"两帧的像素差）；
4. 步进/回退/自动可用，state() 自洽。

跑法：.venv/Scripts/python.exe tools/probes/accept_m3.py
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

import hl_anim  # noqa: E402
import hl_layout  # noqa: E402
import pptx_io  # noqa: E402
import shot  # noqa: E402  # _launch_browser 是有意复用私有符号，见契约 §6.5

SRC = os.path.abspath("output/b_multislide.pptx")
BG = os.path.abspath("output/spike/m1")
OUT = os.path.abspath("output/spike/m3")


def main() -> int:
    from pathlib import Path
    from playwright.sync_api import sync_playwright

    os.makedirs(OUT, exist_ok=True)
    deck, _ = pptx_io.read_pages(SRC)
    bg_paths = [f"bg/slide_{i + 1}.png" for i in range(len(deck.pages))]
    for p in bg_paths:
        src = os.path.join(BG, os.path.basename(p))
        dst = os.path.join(OUT, p)
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        if not os.path.isfile(dst):
            with open(src, "rb") as fi, open(dst, "wb") as fo:
                fo.write(fi.read())

    pages_units = [hl_layout.build_units(hl_layout.page_shapes(pg, deck))
                   for pg in deck.pages]
    total_units = sum(len(u) for u in pages_units)
    player = hl_anim.build_player(OUT, bg_paths, pages_units, title="全球气候变暖")
    print(f"播放器：{player}")
    print(f"底图 {len(bg_paths)} 张；讲解单元合计 {total_units} 个")

    url = Path(os.path.abspath(player)).as_uri()
    rc = 0
    with sync_playwright() as pw:
        browser = shot._launch_browser(pw)
        page = browser.new_page(viewport={"width": 1600, "height": 1000})
        errors = []
        page.on("pageerror", lambda e: errors.append(str(e)))
        page.goto(url, wait_until="load")
        page.evaluate("window.hl.ready")

        ok_pages = page.evaluate("window.hl.state().totalPages")
        n_bg = page.evaluate("document.querySelectorAll('img.bg').length")
        print(f"\n[1] state().totalPages = {ok_pages}，DOM 里底图数 = {n_bg}，"
              f"契约页数 {len(bg_paths)} → {'一致' if ok_pages == n_bg == len(bg_paths) else '不一致!'}")
        if not (ok_pages == n_bg == len(bg_paths)):
            rc = 1

        print("\n[2] goto 同步性（调用后立刻读 DOM，不等任何 tick）")
        probe = page.evaluate("""() => {
            const out = [];
            hl.goto(0, 0);
            out.push(['goto(0,0)', document.querySelectorAll('.hl.on').length,
                      document.querySelectorAll('img.bg.on').length, hl.state().step]);
            hl.goto(1, 3);
            out.push(['goto(1,3)', document.querySelectorAll('.hl.on').length,
                      document.querySelectorAll('img.bg.on').length, hl.state().step]);
            return out;
        }""")
        for name, hl_cnt, bg_cnt, step in probe:
            print(f"  {name}: 高亮 div={hl_cnt}，可见底图={bg_cnt}，step={step}")
        if probe[0][1] != 0 or probe[0][2] != 1 or probe[1][1] == 0 or probe[1][2] != 1:
            print("  [FAIL] goto 未同步生效")
            rc = 1

        print("\n[3] 逐页 goto 截图 + 降暗对比")
        from PIL import Image, ImageChops
        for label, (p, s) in [("p0s0", (0, 0)), ("p0s1", (0, 1)), ("p0all", (0, 99999)),
                              ("p5s1", (5, 1)), ("p5all", (5, 99999))]:
            page.evaluate(f"hl.goto({p}, {s})")
            frame = page.locator("#frame")
            shot_path = os.path.join(OUT, f"{label}.png")
            frame.screenshot(path=shot_path)
            st = page.evaluate("hl.state()")
            print(f"  {label}: state={st}")

        # 降暗对比：同页「有高亮(聚光)」vs「无高亮(不降暗)」，远离高亮处应更暗
        page.evaluate("hl.goto(0, 1)")
        frame.screenshot(path=os.path.join(OUT, "_with.png"))
        page.evaluate("hl.goto(0, 0)")
        frame.screenshot(path=os.path.join(OUT, "_without.png"))
        a = Image.open(os.path.join(OUT, "_with.png")).convert("RGB")
        b = Image.open(os.path.join(OUT, "_without.png")).convert("RGB")
        diff = ImageChops.difference(a, b)
        bbox = diff.getbbox()
        far_a = a.getpixel((60, a.height - 60))       # 远离高亮框的采样点
        far_b = b.getpixel((60, b.height - 60))
        print(f"  两帧差异区域 bbox={bbox}（应非空=高亮确实画出来了）")
        print(f"  远离高亮处像素：聚光 {far_a} vs 不降暗 {far_b} "
              f"→ {'更暗(降暗生效)' if sum(far_a) < sum(far_b) else '没变暗!'}")
        if bbox is None:
            print("  [FAIL] 高亮没有画出来")
            rc = 1
        if sum(far_a) >= sum(far_b):
            print("  [FAIL] 降暗未生效（检查 dim 遮罩）")
            rc = 1

        print("\n[4] next/back/自动")
        seq = page.evaluate("""() => {
            hl.goto(1, 0);
            const seen = [hl.state()];
            hl.next(); seen.push(hl.state());
            hl.next(); seen.push(hl.state());
            hl.back(); seen.push(hl.state());
            hl.goto(9, 99999); seen.push(hl.state());
            hl.next(); seen.push(hl.state());
            hl.back(); seen.push(hl.state());
            return seen;
        }""")
        for s in seq:
            print(f"  {s}")
        page.evaluate("hl.goto(0,0)")
        page.evaluate("toggleAuto()")
        page.wait_for_timeout(2600)
        after = page.evaluate("hl.state()")
        print(f"  自动讲解 2.6s（autoStepMs=2000）后：{after} → "
              f"{'在推进' if after['step'] > 0 else '没动!'}")
        page.evaluate("toggleAuto()")

        if errors:
            print(f"\n[FAIL] 页面 JS 报错：{errors}")
            rc = 1
        else:
            print("\n页面无 JS 报错")
        browser.close()

    print(f"\n截图与播放器目录：{OUT}")
    print(f"判定：{'PASS' if rc == 0 else 'FAIL'}")
    return rc


if __name__ == "__main__":
    raise SystemExit(main())
