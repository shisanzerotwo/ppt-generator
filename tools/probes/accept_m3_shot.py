"""M3 尾验收：对真实稿跑完整步进序列截图，验"截图数 == Σ units、画面非空白"。

跑法：.venv/Scripts/python.exe tools/probes/accept_m3_shot.py
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

import hl_anim  # noqa: E402
import hl_layout  # noqa: E402
import pptx_io  # noqa: E402

SRC = os.path.abspath("output/b_multislide.pptx")
BG = os.path.abspath("output/spike/m1")
OUT = os.path.abspath("output/spike/m3shot")


def main() -> int:
    from PIL import Image

    deck, _ = pptx_io.read_pages(SRC)
    bg_paths = [f"bg/slide_{i + 1}.png" for i in range(len(deck.pages))]
    for rel in bg_paths:
        dst = os.path.join(OUT, rel)
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        if not os.path.isfile(dst):
            with open(os.path.join(BG, os.path.basename(rel)), "rb") as fi, \
                    open(dst, "wb") as fo:
                fo.write(fi.read())

    pages_units = [hl_layout.build_units(hl_layout.page_shapes(p, deck))
                   for p in deck.pages]
    total_units = sum(len(u) for u in pages_units)
    player = hl_anim.build_player(OUT, bg_paths, pages_units, title="全球气候变暖")
    print(f"播放器 {player}；页数 {len(bg_paths)}；Σ units = {total_units}")

    # 步进序列：每页逐单元推进一步（正是 video --mode step 要喂的序列）
    steps = [(pi, si)
             for pi, units in enumerate(pages_units)
             for si in range(1, len(units) + 1)]
    print(f"步进序列长度 = {len(steps)}")

    shots = hl_anim.shot_player(player, os.path.join(OUT, "frames"), steps)
    print(f"截图数 = {len(shots)}（期望 {total_units}，即 Σ units）→ "
          f"{'一致' if len(shots) == total_units else '不一致!'}")

    sizes = set()
    blank = []
    luminance = []
    for s in shots:
        with Image.open(s) as im:
            rgb = im.convert("RGB")
            sizes.add(im.size)
            px = rgb.load()
            mx = max(max(px[x, y]) for y in range(0, rgb.height, 17)
                     for x in range(0, rgb.width, 17))
            luminance.append(mx)
            if mx < 120:
                blank.append(os.path.basename(s))
    print(f"尺寸集合 = {sizes}（期望 {{(1920, 1080)}}）")
    print(f"最暗帧的最大亮度 = {min(luminance)}（阈值 120：低于即判空白）")
    print(f"空白帧 = {blank}（期望 []）")

    ok = len(shots) == total_units and sizes == {(1920, 1080)} and not blank
    print(f"\n截图目录：{os.path.join(OUT, 'frames')}")
    print(f"判定：{'PASS' if ok else 'FAIL'}")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
