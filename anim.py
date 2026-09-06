"""阶段一 v2：教学讲解动画播放器（元素级入场动效）。

需求演进：整页 Ken Burns 轮播 → **元素级动画**（标题、要点、配图逐个出现），
服务教学讲解节奏——老师手动步进为主（点击/方向键），自动播放可选。

实现：播放器以 iframe 同源加载原设计稿（零拷贝、零截图），注入动画引擎：
- 收集每页可动画元素（标题/段落/要点条/配图/表格），加 .ae 类初始隐藏
- 「下一步」逐元素 reveal，「上一步」回退，跨页自动重置
- 设计稿文本保持可选中/可缩放，16:9 舞台等比缩放
"""

import html
import json
import os

TEMPLATE = """<!DOCTYPE html>
<html lang="zh"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>__TITLE__ · 教学动画</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{background:#0b0f17;color:#e8edf5;font-family:"Microsoft YaHei","PingFang SC",system-ui,sans-serif;
  min-height:100vh;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:12px}
h1{font-size:15px;font-weight:600;color:#8b98ab;letter-spacing:.05em}
.stage{position:relative;width:min(96vw,calc(92vh*16/9));aspect-ratio:16/9;background:#000;
  border-radius:10px;overflow:hidden;box-shadow:0 8px 40px rgba(0,0,0,.5)}
.stage iframe{width:1280px;height:720px;border:0;transform-origin:0 0;background:#fff}
.bar{width:min(96vw,calc(92vh*16/9));display:flex;align-items:center;gap:8px;font-size:13px;color:#8b98ab;flex-wrap:wrap}
.bar button{background:#1a2333;border:1px solid #2a3548;color:#e8edf5;border-radius:6px;
  padding:4px 12px;cursor:pointer;font-size:13px}
.bar button:hover{border-color:#3b82f6}
.bar button.primary{border-color:#3b82f6;color:#93c5fd}
.hint{font-size:12px;color:#5b6779}
.pos{margin-left:auto}
.dots{display:flex;gap:5px;flex-wrap:wrap}
.dot{width:9px;height:9px;border-radius:50%;background:#2a3548;cursor:pointer}
.dot.on{background:#3b82f6}
</style></head><body>
<h1>__TITLE__ · 教学动画（逐元素讲解）</h1>
<div class="stage" id="stage"><iframe id="frame" src="__DECK__"></iframe></div>
<div class="bar">
  <button id="auto" class="primary" onclick="toggleAuto()">▶ 自动讲解</button>
  <button onclick="stepBack()">⏮ 上一步</button>
  <button onclick="stepFwd()" class="primary">下一步 ⏭</button>
  <button onclick="revealAll()">显示本页全部</button>
  <button onclick="showPage(cur-1)">‹ 页</button>
  <div class="dots" id="dots"></div>
  <button onclick="showPage(cur+1)">页 ›</button>
  <span class="hint">空格/→ 下一步 · ← 回退 · 点击画面亦可步进</span>
  <span class="pos" id="pos">-/-</span>
</div>
<script>
const CFG = __CONFIG__;
const frame = document.getElementById('frame');
const doc_ = () => frame.contentDocument;
let pages = [], secs = [], cur = 0, idx = 0, auto = null, dotsBuilt = false;

/* 元素选择器：标题优先、要点逐条、配图/表格各自入场（DOM 序即讲解序） */
const SEL = 'h1,h2,h3,h4,p,li,img,svg,video,table';

frame.addEventListener('load', () => {
  const d = doc_();
  secs = [...d.querySelectorAll('.slide')];
  const st = d.createElement('style');
  st.textContent = '.ae{opacity:0;transform:translateY(16px);'
    + 'transition:opacity .45s ease,transform .45s ease}'
    + '.ae.on{opacity:1;transform:none}';
  d.head.appendChild(st);
  pages = secs.map(sec => [...sec.querySelectorAll(SEL)]
    .filter(el => el.getClientRects().length));   // 过滤 display:none 的元素
  pages.forEach(els => els.forEach(el => el.classList.add('ae')));
  if (!dotsBuilt) { buildDots(); dotsBuilt = true; }
  showPage(0);
  // 点击画面步进（教学：讲一下、点一下）
  d.body.addEventListener('click', () => stepFwd());
});

function buildDots() {
  const dots = document.getElementById('dots');
  dots.innerHTML = '';
  secs.forEach((_, i) => {
    const d = document.createElement('div');
    d.className = 'dot'; d.onclick = () => showPage(i);
    dots.appendChild(d);
  });
}

function showPage(i) {
  if (!pages.length) return;
  cur = (i + pages.length) % pages.length;
  secs[cur].scrollIntoView({behavior: 'instant', block: 'start'});
  idx = 0;
  pages[cur].forEach(el => el.classList.remove('on'));
  stepFwd();                                     // 进页先出第一个元素（标题）
  syncUI();
}

function stepFwd() {
  if (!pages.length) return;
  const els = pages[cur];
  if (idx < els.length) { els[idx].classList.add('on'); idx++; }
  else if (cur < pages.length - 1) showPage(cur + 1);
  syncUI();
}

function stepBack() {
  const els = pages[cur];
  if (idx > 1) { idx--; els[idx].classList.remove('on'); }
  else if (cur > 0) {
    showPage(cur - 1);
    revealAll();                                  // 回到上一页：全显便于衔接
  }
  syncUI();
}

function revealAll() {
  pages[cur].forEach(el => el.classList.add('on'));
  idx = pages[cur].length;
  syncUI();
}

function toggleAuto() {
  auto = auto ? null : setInterval(stepFwd, CFG.autoStepMs);
  document.getElementById('auto').textContent = auto ? '⏸ 停止自动' : '▶ 自动讲解';
  if (!auto) return;
  stepFwd();
}

function syncUI() {
  document.getElementById('pos').textContent = (cur + 1) + '/' + secs.length
    + ' · 元素 ' + idx + '/' + pages[cur].length;
  [...document.getElementById('dots').children].forEach((d, k) =>
    d.classList.toggle('on', k === cur));
}

document.addEventListener('keydown', e => {
  if (e.key === 'ArrowRight' || e.key === ' ') { e.preventDefault(); stepFwd(); }
  if (e.key === 'ArrowLeft') { e.preventDefault(); stepBack(); }
});
window.addEventListener('resize', fit);
function fit() {
  const stage = document.getElementById('stage');
  const f = frame;
  f.style.transform = 'scale(' + (stage.clientWidth / 1280) + ')';
}
frame.addEventListener('load', fit);
fit();
</script>
</body></html>
"""


def build_player(out_dir: str, deck_web_path: str, title: str = "",
                 auto_step_ms: int = 1200) -> str:
    """在 out_dir 写 index.html 教学动画播放器，返回其路径。

    deck_web_path 为设计稿的站内 URL（如 /decks/xxx.html，同源 iframe 直载）。
    title/deck 来自用户主题与路径，必须按上下文转义后再进模板（审计：主题
    含 </title><script> 可注入播放器——title 进 HTML 文本与 <title> 两个上下文，
    deck 进 src 属性上下文，CONFIG 进 <script> 字符串上下文，三处分别处理）。
    """
    safe_title = html.escape(title or "PPT 教学动画")
    safe_deck = html.escape(deck_web_path or "", quote=True)
    cfg = json.dumps({"deck": deck_web_path or "", "title": title or "PPT 教学动画",
                      "autoStepMs": auto_step_ms}, ensure_ascii=False)
    # <script> 字符串上下文：JSON 里的 </ 可能闭合 script 标签，统一转义
    cfg = cfg.replace("</", "<\\/")
    doc = (TEMPLATE
           .replace("__TITLE__", safe_title)
           .replace("__DECK__", safe_deck)
           .replace("__CONFIG__", cfg))
    path = os.path.join(out_dir, "index.html")
    os.makedirs(out_dir, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write(doc)
    return path
