"""shot.py 截图确定化回归测试：真实 Chrome 截图（无可用浏览器则跳过）。

背景（OmniRoute vision 看图实测抓到的缺陷）：设计稿的图表柱由 IntersectionObserver
在滚动后填充，而无头截图上下文里 IO 回调不可靠——实测等 3s 柱高仍是 0，
于是**空图表**被截进图片与视频。修法不是加长缓冲（实测无效），而是注入 IO shim
让回调立即派发 + `screenshot(animations="disabled")` + 稳定性轮询兜底。

这里锁四件事：①静态稿早返回 ②逐页定位/画幅不变 ③IO 驱动动画截到终态（核心回归）
④JS 驱动的永不停歇动画被上限兜住、不卡死导出。
"""

import os
import time

import pytest

import shot


def _run(html: str, tmp_path, tag: str, **kw):
    deck = tmp_path / f"{tag}.html"
    deck.write_text(html, encoding="utf-8")
    try:
        return shot.shot_deck(str(deck), str(tmp_path / tag), **kw)
    except RuntimeError as e:  # 本机没有 Chrome/Edge
        pytest.skip(f"无可用浏览器：{e}")


def _pixel(path, xy=(50, 200)):
    from PIL import Image
    with Image.open(path) as im:
        return im.convert("RGB").getpixel(xy)


def _static_deck(n=12):
    colors = ["0a0b0c", "1b2c3d", "2c3d4e", "3d4e5f", "4e5f6a", "5f6a7b",
              "6a7b8c", "7b8c9d", "8c9d0e", "9d0e1f", "a0b1c2", "b1c2d3"][:n]
    secs = "".join(f'<section class="slide" style="width:1280px;height:720px;'
                   f'background:#{c}"></section>' for c in colors)
    return f'<!doctype html><html><body style="margin:0">{secs}</body></html>'


def _io_animated_deck():
    """复刻数据页机制：IO(threshold .5) 触发后每 120ms 长 25px，至 400px 停。

    注意：没有 IO shim 时这段回调在无头截图上下文里根本不派发（实测），
    所以本用例同时是「shim 生效」的证明。
    """
    return """<!doctype html><html><body style="margin:0;background:#fff">
<section class="slide" style="width:1280px;height:720px;background:#fff">
<div id="bar" style="width:100px;height:0;background:#333"></div></section>
<script>
const bar=document.getElementById('bar');
new IntersectionObserver(es=>es.forEach(e=>{if(e.isIntersecting){
  requestAnimationFrame(()=>requestAnimationFrame(()=>{
    let h=0;const t=setInterval(()=>{h+=25;bar.style.height=h+'px';if(h>=400)clearInterval(t);},120);
  }));}}),{threshold:0.5}).observe(document.querySelector('.slide'));
</script></body></html>"""


def _js_never_settles_deck():
    """JS 定时器永不停歇（CSS 动画会被 animations='disabled' 干掉，故用 JS）。"""
    return """<!doctype html><html><body style="margin:0">
<section class="slide" style="width:1280px;height:720px;background:#fff">
<div id="p" style="width:400px;height:400px;background:#333"></div></section>
<script>let i=0;setInterval(()=>{i++;
document.getElementById('p').style.background=(i%2)?'#333':'#eee';},150);</script>
</body></html>"""


def test_static_deck_returns_early(tmp_path):
    """画面本就静态：不该跑满上限（12 页上限 24s）。"""
    t0 = time.time()
    shots = _run(_static_deck(), tmp_path, "static")
    dt = time.time() - t0
    assert len(shots) == 12
    assert dt < 40, f"静态稿不该等满上限（实测 {dt:.1f}s）"


def test_pages_full_size_and_distinct(tmp_path):
    """逐页滚动定位仍生效：12 页全 1280×720 且内容各不相同。"""
    from PIL import Image
    shots = _run(_static_deck(), tmp_path, "distinct")
    seen = set()
    for s in shots:
        with Image.open(s) as im:
            assert im.size == (1280, 720)
            seen.add(im.convert("RGB").getpixel((640, 360)))
    assert len(seen) == 12


def test_io_driven_animation_captured_at_final_state(tmp_path):
    """核心回归：IO 驱动的图表必须截到终态（柱 > 200px），不是空白。"""
    shots = _run(_io_animated_deck(), tmp_path, "ioanim")
    assert _pixel(shots[0]) == (51, 51, 51), "截到空白=IO shim 或稳定等待失效"


def test_js_never_settling_page_capped(tmp_path):
    """永不停歇的 JS 动画：到上限就用最后一帧，不抛错、不卡死（该页仍产出）。"""
    t0 = time.time()
    shots = _run(_js_never_settles_deck(), tmp_path, "nosettle",
                 min_settle_ms=0, max_settle_ms=800)
    dt = time.time() - t0
    assert len(shots) == 1 and os.path.isfile(shots[0])
    assert dt >= 0.8, "应至少等到上限才放弃"
    assert dt < 20, f"上限应兜住耗时（实测 {dt:.1f}s）"
