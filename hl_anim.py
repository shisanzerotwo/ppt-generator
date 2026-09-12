"""高亮讲解播放器：底图（PowerPoint 导出）+ 行级高亮锚点。

与 `anim.py` 的关键差异（契约 §6.2）
------------------------------------
`anim.py` 必须用 iframe——它要把**设计稿 HTML** 连脚本一起加载再注入动画引擎。
本播放器的"幻灯片"是**一张 PNG 底图** + 绝对定位的高亮 div，同源限制在这里毫无
收益，所以**零 iframe**：攻击面比 anim.py 更小。（若将来为 redesign 预览引入
iframe，必须 `sandbox="allow-scripts"` 且**禁止** `allow-same-origin`。）

为什么底图用「多张 img 叠放 + display 切换」而不是「单张 img 换 src」
------------------------------------------------------------------------
契约 §6.4 要求 `window.hl.goto()` **同步生效**（无 CSS 过渡、无 rAF 依赖），
因为截图侧只靠 `screenshot(animations="disabled")` 兜底，不想再引入等待源。
换 `src` 会触发异步图片加载，`goto` 就不再同步；预加载全部底图后按页切
`display`，DOM 变更完全同步。

转义纪律（照抄 anim.py 的三上下文 + 单遍替换，契约 §6.3）
----------------------------------------------------------
- title 进 HTML 文本与 `<title>` 两个上下文 → `html.escape`；
- 底图路径进 `src`：除 HTML 转义外还要**百分号编码**（`#` 会被当成 fragment、
  `%` 会被当成转义引导符，裸拼 URL 会请求到错文件）；
- JSON（配置 / 底图表 / 单元表）进 `<script>` 字符串 → `json.dumps` 后把 `</`
  与 `<!--` 转义，防止提前闭合 script 或让解析器进注释态；
- 模板占位符**单遍** `re.sub` 替换——链式 `str.replace` 会让后一次替换扫到前一次
  插入的文本（title 里含 `__CONFIG__` 字样时可绕过转义注入活标签）。
"""

import html
import json
import os
import re
from urllib.parse import quote

from hl_layout import Unit
from pptx_io import PptxError

DEFAULT_CANVAS_W = 1920
DEFAULT_CANVAS_H = 1080

_PLACEHOLDERS = ("__TITLE__", "__BGJSON__", "__UNITS__", "__CONFIG__")

TEMPLATE = """<!DOCTYPE html>
<html lang="zh"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>__TITLE__ · 高亮讲解</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{background:#0b0f17;color:#e8edf5;font-family:"Microsoft YaHei","PingFang SC",system-ui,sans-serif;
  min-height:100vh;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:12px}
h1{font-size:15px;font-weight:600;color:#8b98ab;letter-spacing:.05em}
.frame{position:relative;width:min(96vw,calc(88vh*16/9));aspect-ratio:16/9;
  border-radius:10px;overflow:hidden;box-shadow:0 8px 40px rgba(0,0,0,.5);background:#000}
.stage{position:absolute;left:0;top:0;transform-origin:0 0}
.bg{position:absolute;left:0;top:0;width:100%;height:100%;display:none;user-select:none;-webkit-user-drag:none}
.bg.on{display:block}
.dim{position:absolute;left:0;top:0;width:100%;height:100%;pointer-events:none}
.hl{position:absolute;pointer-events:none;border:2px solid #ffd166;background:rgba(255,209,102,.18);
  border-radius:3px;display:none}
.hl.on{display:block}
.bar{width:min(96vw,calc(88vh*16/9));display:flex;align-items:center;gap:8px;font-size:13px;
  color:#8b98ab;flex-wrap:wrap}
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
<h1>__TITLE__ · 高亮讲解</h1>
<div class="frame" id="frame"><div class="stage" id="stage">
<svg class="dim" id="dim" preserveAspectRatio="none"><defs>
<mask id="holed" maskUnits="userSpaceOnUse"><rect x="0" y="0" width="100%" height="100%" fill="#fff"/>
<g id="holes"></g></mask></defs>
<rect x="0" y="0" width="100%" height="100%" class="dimrect" mask="url(#holed)"/></svg>
</div></div>
<div class="bar">
  <button id="auto" class="primary" onclick="toggleAuto()">▶ 自动讲解</button>
  <button onclick="hl.back()">⏮ 上一步</button>
  <button onclick="hl.next()" class="primary">下一步 ⏭</button>
  <button onclick="revealAll()">显示本页全部</button>
  <button onclick="hl.goto(hl.state().page-1, 99999)">‹ 页</button>
  <div class="dots" id="dots"></div>
  <button onclick="hl.goto(hl.state().page+1, 1)">页 ›</button>
  <span class="hint">空格/→ 下一步 · ← 回退 · 点击画面亦可步进</span>
  <span class="pos" id="pos">-/-</span>
</div>
<script>
const CFG = __CONFIG__;
const BGS = __BGJSON__;
const PAGES = __UNITS__;

const W = CFG.canvasWidth, H = CFG.canvasHeight;
const stage = document.getElementById('stage');
const frame = document.getElementById('frame');
const dimg = document.getElementById('dim');
const dimrect = dimg.querySelector('.dimrect');
const holes = document.getElementById('holes');
const dotsEl = document.getElementById('dots');

stage.style.width = W + 'px';
stage.style.height = H + 'px';
dimg.setAttribute('viewBox', '0 0 ' + W + ' ' + H);
dimg.setAttribute('width', W);
dimg.setAttribute('height', H);
const maskEl = dimg.querySelector('mask');
maskEl.setAttribute('x', 0);
maskEl.setAttribute('y', 0);
maskEl.setAttribute('width', W);
maskEl.setAttribute('height', H);
dimrect.setAttribute('fill', 'rgba(0,0,0,' + CFG.dim + ')');

/* 底图一次性建好并全部保留在 DOM 里：goto 靠切 display 同步生效，不再等图片加载 */
const imgs = BGS.map((src, i) => {
  const im = document.createElement('img');
  im.className = 'bg';
  im.setAttribute('src', src);           // 走 JS 赋值，不经过 HTML 解析
  im.alt = '第 ' + (i + 1) + ' 页';
  stage.insertBefore(im, dimg);
  return im;
});

const hlRects = [];                       // 复用的高亮 div 池
function rectAt(k) {
  while (hlRects.length <= k) {
    const d = document.createElement('div');
    d.className = 'hl';
    stage.appendChild(d);
    hlRects.push(d);
  }
  return hlRects[k];
}

let page = 0, step = 0;

function clearHl() {
  hlRects.forEach(d => { d.classList.remove('on'); });
  holes.innerHTML = '';
}

function showUnit(u) {
  let used = 0;
  u.l.forEach(box => {
    const d = rectAt(used++);
    d.style.left = box[0] + 'px';
    d.style.top = box[1] + 'px';
    d.style.width = box[2] + 'px';
    d.style.height = box[3] + 'px';
    d.classList.add('on');
    const hole = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
    hole.setAttribute('x', box[0]);
    hole.setAttribute('y', box[1]);
    hole.setAttribute('width', box[2]);
    hole.setAttribute('height', box[3]);
    hole.setAttribute('fill', '#000');    // mask 里画黑 = 挖洞（不降暗）
    holes.appendChild(hole);
  });
}

function render() {
  imgs.forEach((im, i) => im.classList.toggle('on', i === page));
  clearHl();
  const units = PAGES[page] || [];
  for (let i = 0; i < Math.min(step, units.length); i++) showUnit(units[i]);
  // 聚光模式：只在本页已有高亮时才降暗其余区域。step=0（刚进页）不降暗，
  // 否则观众第一眼看到的是整页变暗、什么都看不清。
  dimg.style.display = (step > 0 && units.length) ? 'block' : 'none';
  const total = units.length;
  document.getElementById('pos').textContent =
    (page + 1) + '/' + PAGES.length + ' · 讲解 ' + Math.min(step, total) + '/' + total;
  if (dotsEl.children[page]) {
    [...dotsEl.children].forEach((d, k) => d.classList.toggle('on', k === page));
  }
}

function clampPage(p) {
  if (!PAGES.length) return 0;
  if (p < 0) return 0;
  if (p > PAGES.length - 1) return PAGES.length - 1;
  return p;
}

/* 供程序驱动（video --mode step 依赖它）：goto 必须同步改 DOM */
window.hl = {
  goto(p, s) {
    page = clampPage(p === undefined ? page : p);
    const total = (PAGES[page] || []).length;
    step = Math.max(0, Math.min(s === undefined ? step : s, total));
    render();
    return this.state();
  },
  next() {
    const total = (PAGES[page] || []).length;
    if (step < total) step += 1;
    else if (page < PAGES.length - 1) { page += 1; step = 1; }
    render();
    return this.state();
  },
  back() {
    if (step > 0) step -= 1;
    else if (page > 0) { page -= 1; step = (PAGES[page] || []).length; }
    render();
    return this.state();
  },
  state() {
    return {page: page, step: step, totalPages: PAGES.length,
            totalUnits: (PAGES[page] || []).length};
  },
  ready: Promise.resolve().then(() => Promise.all(imgs.map(im =>
    im.complete ? null : new Promise(res => { im.onload = im.onerror = res; })
  ))).then(() => { fit(); render(); return true; })
};

function revealAll() { hl.goto(page, 99999); }
function toggleAuto() {
  if (autoTimer) { clearInterval(autoTimer); autoTimer = null; }
  else {
    autoTimer = setInterval(() => hl.next(), CFG.autoStepMs);
    hl.next();
  }
  document.getElementById('auto').textContent = autoTimer ? '⏸ 停止自动' : '▶ 自动讲解';
}
let autoTimer = null;

function buildDots() {
  dotsEl.innerHTML = '';
  PAGES.forEach((_, i) => {
    const d = document.createElement('div');
    d.className = 'dot';
    d.onclick = () => hl.goto(i, 1);
    dotsEl.appendChild(d);
  });
}

frame.addEventListener('click', () => hl.next());
document.addEventListener('keydown', e => {
  if (e.key === 'ArrowRight' || e.key === ' ') { e.preventDefault(); hl.next(); }
  if (e.key === 'ArrowLeft') { e.preventDefault(); hl.back(); }
});
window.addEventListener('resize', fit);
function fit() {
  const s = frame.clientWidth / W;
  stage.style.transform = 'scale(' + s + ')';
}

buildDots();
fit();
hl.ready.then(fit);
window.addEventListener('load', fit);
</script>
</body></html>
"""


def _js_json(value) -> str:
    """对象 → 可安全放进 <script> 字符串上下文的 JSON 字面量。"""
    text = json.dumps(value, ensure_ascii=False, separators=(",", ":"))
    # </ 会闭合 script 标签；<!-- 会让解析器进注释态使末尾 </script> 失效
    return text.replace("</", "<\\/").replace("<!--", "<\\!--")


def _web_path(path: str) -> str:
    """相对路径 → 可放进 URL 的形式：百分号编码（`#`/`%`/空格都会坏掉裸拼）。"""
    return quote(str(path).replace("\\", "/"), safe="/")


def _check_bg_path(path: str) -> None:
    """底图必须是 out_dir 内的相对路径（防被改过的 deck.json 指向任意文件）。"""
    p = str(path).replace("\\", "/")
    if os.path.isabs(p) or p.startswith("//") or ":" in p.split("/")[0]:
        raise PptxError(f"底图路径必须是相对路径：{path}", "IR_MISMATCH",
                        "重新执行 pptgen import 生成 deck.json")
    normalized = os.path.normpath(p).replace("\\", "/")
    if normalized.startswith("../") or normalized == "..":
        raise PptxError(f"底图路径逃出输出目录：{path}", "IR_MISMATCH",
                        "重新执行 pptgen import 生成 deck.json")


def _rebase_bg(path: str, out_dir: str, bg_base_dir: str) -> str:
    """把「相对 bg_base_dir」的底图路径改写成「相对 out_dir（播放器所在目录）」。

    契约 §5 的目录树是 `<out>/player/index.html` + `<out>/bg/*.png`，而 §6.1 的
    `bg_paths` 又写成"相对 out_dir"。两者对 out_dir 的所指不一致：播放器实际在
    `<out>/player/` 下，浏览器按**播放器文件所在目录**解析 src，直接用 `bg/x.png`
    会请求 `<out>/player/bg/x.png`（不存在）→ 页面全黑。这里显式做一次 rebase。
    """
    base = os.path.abspath(bg_base_dir)
    target = os.path.abspath(os.path.join(base, str(path).replace("\\", "/")))
    if target != base and not target.startswith(base + os.sep):
        raise PptxError(f"底图路径逃出输出目录：{path}", "IR_MISMATCH",
                        "重新执行 pptgen import 生成 deck.json")
    return os.path.relpath(target, os.path.abspath(out_dir)).replace("\\", "/")


def _unit_payload(u: Unit) -> dict:
    return {
        "k": u.kind,
        "t": u.text,
        "l": [[round(r.left_px, 2), round(r.top_px, 2),
               round(r.width_px, 2), round(r.height_px, 2)] for r in u.lines],
    }


def build_player(out_dir: str, bg_paths: list, pages_units: list,
                 title: str = "", dim: float = 0.25, auto_step_ms: int = 2000,
                 canvas_width_px: int = DEFAULT_CANVAS_W,
                 canvas_height_px: int = DEFAULT_CANVAS_H,
                 bg_base_dir: str | None = None) -> str:
    """写 <out_dir>/index.html，返回其路径。零截图、零 iframe、零外部依赖。

    bg_paths 默认已相对 out_dir；若它们相对别的目录（CLI 的 `<out>/player/` 布局
    就是这种情形），传 `bg_base_dir` 让本函数负责改写（见 `_rebase_bg`）。
    """
    if len(bg_paths) != len(pages_units):
        raise PptxError(
            f"底图数与讲解单元页数不一致：{len(bg_paths)} vs {len(pages_units)}",
            "IR_MISMATCH", "重新执行 pptgen import 生成 deck.json")
    try:
        dim = float(dim)
    except (TypeError, ValueError) as exc:
        raise PptxError(f"降暗比例非法：{dim!r}", "BAD_ARGS",
                        "降暗比例需在 (0, 0.6) 之间") from exc
    if not (0.0 < dim < 0.6):
        raise PptxError("降暗比例需在 (0, 0.6) 之间", "BAD_ARGS",
                        "例如 --dim 0.25")
    try:
        auto_step_ms = int(auto_step_ms)
        canvas_width_px = int(canvas_width_px)
        canvas_height_px = int(canvas_height_px)
    except (TypeError, ValueError) as exc:
        raise PptxError(f"数值参数非法：{auto_step_ms!r}", "BAD_ARGS",
                        "自动步进毫秒数与画布尺寸需为整数") from exc
    if auto_step_ms <= 0 or canvas_width_px <= 0 or canvas_height_px <= 0:
        raise PptxError("自动步进毫秒数与画布尺寸需为正整数", "BAD_ARGS", "例如 --auto-ms 2000")

    for p in bg_paths:
        _check_bg_path(p)

    safe_title = html.escape(str(title or "PPT 高亮讲解"))
    web_paths = [_rebase_bg(p, out_dir, bg_base_dir) if bg_base_dir else _web_path(p)
                 for p in bg_paths]
    bg_json = _js_json(web_paths)
    units_json = _js_json([[_unit_payload(u) for u in page] for page in pages_units])
    cfg = _js_json({
        "title": str(title or "PPT 高亮讲解"),
        "dim": round(dim, 4),
        "autoStepMs": auto_step_ms,
        "canvasWidth": canvas_width_px,
        "canvasHeight": canvas_height_px,
    })

    # 单遍替换：链式 str.replace 会让后续替换扫到先前插入的文本
    mapping = {"__TITLE__": safe_title, "__BGJSON__": bg_json,
               "__UNITS__": units_json, "__CONFIG__": cfg}
    doc = re.sub("|".join(_PLACEHOLDERS), lambda m: mapping[m.group(0)], TEMPLATE)

    os.makedirs(out_dir, exist_ok=True)
    path = os.path.join(out_dir, "index.html")
    with open(path, "w", encoding="utf-8") as f:
        f.write(doc)
    return path


# ---------------------------------------------------------------- 步进截图

def shot_player(player_html: str, out_dir: str, steps: list,
                min_settle_ms: int = 300, max_settle_ms: int = 2000) -> list[str]:
    """逐 (page, step) 驱动 `window.hl.goto` → 截图 → `step_0001.png…`，返回路径列表。

    沿用 `shot.py` 的三重确定化（KNOWLEDGE.md 红线：**别退回固定 sleep**）：
    ① `window.hl.ready`；② `screenshot(animations="disabled")`；③ 连续两帧字节一致轮询。

    两处与契约 §6.5 不同的实现选择：
    - `_launch_browser` / `_screenshot_settled` 是 `shot.py` 的私有名，这里**有意复用**
      （契约 §6.5 明列的二选一退路）。没有复制第三份确定化实现。
    - **截图前先校验底图真的加载出来了**（审计 L4）：播放器本身不校验文件存在，
      缺图时页面全黑而 `hl.ready` 照样 resolve，会把黑帧安静地截进 MP4。这里用
      `naturalWidth > 0` 兜底，一张缺失即报 IR_MISMATCH。

    截图尺寸恒等于播放器画布（先把 `.frame` 撑到画布尺寸、令缩放比为 1 再截
    `#frame`），因此不受窗口大小影响。
    """
    from pathlib import Path

    from playwright.sync_api import sync_playwright

    import shot  # 有意复用私有符号，见契约 §6.5

    steps = list(steps or [])
    if not steps:
        raise PptxError("步进序列为空，没有可截图的步骤", "BAD_ARGS", "至少给一个 (page, step)")
    if not os.path.isfile(player_html):
        raise PptxError(f"找不到播放器页面：{player_html}", "IR_MISMATCH",
                        "先执行 pptgen animate 生成 player/index.html")

    os.makedirs(out_dir, exist_ok=True)
    url = Path(os.path.abspath(player_html)).as_uri()
    paths: list[str] = []
    with sync_playwright() as p:
        browser = shot._launch_browser(p)
        try:
            page = browser.new_page(viewport={"width": 640, "height": 360})
            page.goto(url, wait_until="load")
            page.evaluate("window.hl.ready")

            size = page.evaluate(
                "() => ({w: CFG.canvasWidth, h: CFG.canvasHeight})")
            w, h = int(size["w"]), int(size["h"])
            page.set_viewport_size({"width": w, "height": h})
            # 让舞台独占视口：截图走 shot._screenshot_settled 的 page.screenshot，
            # 截的是视口而非元素，所以必须把标题/按钮栏藏掉、body 去掉居中，
            # 否则截出来是"页面"而不是"幻灯片"（画布外露着 body 底色）。
            page.evaluate(
                """([w, h]) => {
                    document.querySelector('h1').style.display = 'none';
                    document.querySelector('.bar').style.display = 'none';
                    document.body.style.cssText =
                        'margin:0;padding:0;display:block;overflow:hidden;background:#000';
                    const f = document.getElementById('frame');
                    f.style.width = w + 'px';
                    f.style.height = h + 'px';
                    f.style.aspectRatio = 'auto';
                    f.style.borderRadius = '0';
                    f.style.boxShadow = 'none';
                    fit();
                }""", [w, h])

            check = page.evaluate(
                "() => ({total: BGS.length,"
                "        failed: BGS.filter((s, i) => imgs[i].naturalWidth === 0)})")
            if check["failed"]:
                raise PptxError(
                    f"底图加载失败 {len(check['failed'])}/{check['total']}：{check['failed'][0]}",
                    "IR_MISMATCH", "重新执行 pptgen import 生成底图")

            frame = page.locator("#frame")
            for k, (pg, st) in enumerate(steps, 1):
                page.evaluate("([p, s]) => window.hl.goto(p, s)", [int(pg), int(st)])
                data = shot._screenshot_settled(page, min_settle_ms, max_settle_ms)
                out = os.path.abspath(os.path.join(out_dir, f"step_{k:04d}.png"))
                with open(out, "wb") as f:
                    f.write(data)
                paths.append(out)
        finally:
            browser.close()
    return paths
