"""shot.py 截图管线（阶段二视频的底图来源）：真实 Chrome 截 16:9 页图。"""

import os

import pytest

import shot


def test_shot_deck_real_screenshot(tmp_path):
    """真实启动无头 Chrome 截两页 16:9 图（依赖本机 Chrome/Edge，缺失则跳过）。"""
    deck = tmp_path / "deck.html"
    deck.write_text(
        "<style>html{scroll-snap-type:y mandatory}.slide{height:100vh;scroll-snap-align:start;"
        "font-size:40px}</style>"
        "<div class='slide'>第一页</div><div class='slide'>第二页</div>",
        encoding="utf-8",
    )
    try:
        shots = shot.shot_deck(str(deck), str(tmp_path / "shots"))
    except RuntimeError as e:
        pytest.skip(f"无可用浏览器：{e}")
    assert len(shots) == 2 and all(os.path.isfile(p) for p in shots)
    from PIL import Image
    with Image.open(shots[0]) as im:
        assert im.size == (1280, 720)  # 16:9 画幅
