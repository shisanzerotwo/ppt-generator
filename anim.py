"""阶段一 M2：把逐页截图打包成零依赖单文件动画播放器（Ken Burns + 交叉淡入）。

设计取舍：
- 纯 HTML/CSS/JS、无外部资源——延续项目"零依赖单文件"哲学，双击即可放映
- 16:9 舞台 letterbox 居中（object-fit: cover 裁满），任意窗口尺寸不变形
- Ken Burns 用 CSS animation 实现（随机方向由页序种子决定，可复现）；
  页间切换用 JS 定时 + opacity 过渡（交叉淡入）
- 参数（每页时长/转场时长/图片清单）以 JSON 注入 <script>，生成端可控
"""

import json
import os

TEMPLATE = """<!DOCTYPE html>
<html lang="zh"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>__TITLE__ · 动画版</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{background:#0b0f17;color:#e8edf5;font-family:"Microsoft YaHei","PingFang SC",system-ui,sans-serif;
  min-height:100vh;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:14px}
h1{font-size:15px;font-weight:600;color:#8b98ab;letter-spacing:.05em}
.stage{position:relative;width:min(96vw,calc(92vh*16/9));aspect-ratio:16/9;background:#000;
  border-radius:10px;overflow:hidden;box-shadow:0 8px 40px rgba(0,0,0,.5)}
.slide{position:absolute;inset:0;opacity:0;transition:opacity var(--fade) ease;pointer-events:none}
.slide.on{opacity:1;pointer-events:auto}
.slide img{width:100%;height:100%;object-fit:cover;animation:kb var(--dur) ease-out forwards}
@keyframes kb{from{transform:scale(1) translate(0,0)}to{transform:var(--kb-to)}}
.kb0{--kb-to:scale(1.08) translate(1.2%,0.8%)}
.kb1{--kb-to:scale(1.08) translate(-1.2%,-0.8%)}
.kb2{--kb-to:scale(1.06) translate(0,-1%)}
.kb3{--kb-to:scale(1.06) translate(0,1%)}
.bar{width:min(96vw,calc(92vh*16/9));display:flex;align-items:center;gap:10px;font-size:13px;color:#8b98ab}
.bar button{background:#1a2333;border:1px solid #2a3548;color:#e8edf5;border-radius:6px;
  padding:4px 12px;cursor:pointer;font-size:13px}
.bar button:hover{border-color:#3b82f6}
.dots{display:flex;gap:5px;flex:1;justify-content:center;flex-wrap:wrap}
.dot{width:8px;height:8px;border-radius:50%;background:#2a3548;cursor:pointer}
.dot.on{background:#3b82f6}
.pos{min-width:44px;text-align:right}
</style></head><body>
<h1>__TITLE__ · 动画放映</h1>
<div class="stage" id="stage"></div>
<div class="bar">
  <button id="play">⏸ 暂停</button>
  <button onclick="go(cur-1)">‹</button>
  <div class="dots" id="dots"></div>
  <button onclick="go(cur+1)">›</button>
  <span class="pos" id="pos">1/__N__</span>
</div>
<script>
const CFG = __CONFIG__;
const stage = document.getElementById('stage');
let cur = 0, playing = true, timer = null;

CFG.images.forEach((src, i) => {
  const d = document.createElement('div');
  d.className = 'slide kb' + (i % 4);
  d.innerHTML = '<img src="' + src + '" alt="第' + (i+1) + '页">';
  stage.appendChild(d);
});
const slides = [...document.querySelectorAll('.slide')];
const dots = document.getElementById('dots');
CFG.images.forEach((_, i) => {
  const d = document.createElement('div');
  d.className = 'dot'; d.onclick = () => go(i); dots.appendChild(d);
});

function show(i) {
  cur = (i + CFG.images.length) % CFG.images.length;
  slides.forEach((s, k) => {
    s.classList.toggle('on', k === cur);
    // 重启 Ken Burns 动画：仅当前页播放
    const img = s.querySelector('img');
    img.style.animation = 'none'; void img.offsetWidth;
    if (k === cur) img.style.animation = '';
  });
  [...dots.children].forEach((d, k) => d.classList.toggle('on', k === cur));
  document.getElementById('pos').textContent = (cur+1) + '/' + CFG.images.length;
  if (playing) schedule();
}
function schedule() {
  clearTimeout(timer);
  if (playing) timer = setTimeout(() => show(cur + 1), CFG.seconds * 1000);
}
function go(i) { show(i); }
document.getElementById('play').onclick = function () {
  playing = !playing;
  this.textContent = playing ? '⏸ 暂停' : '▶ 播放';
  schedule();
};
document.addEventListener('keydown', e => {
  if (e.key === 'ArrowRight' || e.key === ' ') go(cur + 1);
  if (e.key === 'ArrowLeft') go(cur - 1);
  if (e.key === 'Escape') playing = false;
});
show(0);
</script>
</body></html>
"""


def build_player(out_dir: str, shots: list[str], title: str = "",
                 seconds: float = 4.0, fade: float = 0.8) -> str:
    """在 out_dir 写 index.html 动画播放器，返回其路径。

    图片用相对路径（与 index.html 同目录的 slide_N.png），整目录可搬走。
    """
    images = [os.path.basename(p) for p in shots]
    cfg = json.dumps({"images": images, "seconds": seconds, "fade": fade},
                     ensure_ascii=False)
    doc = (TEMPLATE
           .replace("__TITLE__", title or "PPT 动画")
           .replace("__N__", str(len(images)))
           .replace("__CONFIG__", cfg))
    path = os.path.join(out_dir, "index.html")
    with open(path, "w", encoding="utf-8") as f:
        f.write(doc)
    return path
