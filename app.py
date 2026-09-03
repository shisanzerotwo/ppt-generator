"""可视化制作 Web 界面后端：python app.py 后浏览器打开 http://127.0.0.1:5000"""

import os
import re
import threading
import time

from flask import Flask, jsonify, render_template, request, send_from_directory

import builder
import image_gen
import outline

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
IMAGES_DIR = os.path.join(OUTPUT_DIR, "images")

state = {
    "topic": "",
    "phase": "idle",  # idle / outline / images / ready
    "theme": "blue",  # blue / dark / green（色板来自 ppt-maker skill）
    "slides": [],     # {title, points, image_prompt, image, imageStatus}
    "log": [],
}
lock = threading.Lock()


def _log(msg: str):
    with lock:
        state["log"].append({"time": time.strftime("%H:%M:%S"), "msg": msg})


def _phase(p: str):
    with lock:
        state["phase"] = p


def _generation_worker(topic: str):
    state["topic"] = topic
    state["slides"] = []
    state["log"] = []
    _phase("outline")
    try:
        _log(f"生成大纲：{topic}")
        slides = outline.generate_outline(topic)
        with lock:
            state["slides"] = [
                {"type": s.get("type", "content"), "title": s.get("title", ""),
                 "points": s.get("points", []), "image_prompt": s.get("image_prompt", ""),
                 "chart": s.get("chart"), "image": None, "imageStatus": "pending"}
                for s in slides
            ]
        _log(f"大纲完成，共 {len(slides)} 页，开始逐页生图")
        _phase("images")

        for i, s in enumerate(slides):
            prompt = s.get("image_prompt", "")
            if not prompt:
                with lock:
                    state["slides"][i]["imageStatus"] = "skipped"
                _log(f"页 {i + 1}：无配图提示词，跳过")
                continue
            path = os.path.join(IMAGES_DIR, f"slide_{i}.png")
            try:
                image_gen.generate_image(prompt, path)
                with lock:
                    state["slides"][i]["image"] = f"/images/slide_{i}.png"
                    state["slides"][i]["imageStatus"] = "done"
                _log(f"页 {i + 1} 图片完成")
            except Exception as e:
                with lock:
                    state["slides"][i]["imageStatus"] = "failed"
                _log(f"页 {i + 1} 图片生成失败：{e}")
        _log("全部页面就绪，可编辑后导出")
        _phase("ready")
    except Exception as e:
        _log(f"生成中断：{e}")
        _phase("idle")


app = Flask(__name__)


@app.after_request
def no_cache(resp):
    # 开发期禁用缓存，保证前端每次刷新都拿到最新页面
    resp.headers["Cache-Control"] = "no-store, max-age=0"
    return resp


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/generate", methods=["POST"])
def api_generate():
    data = request.get_json(force=True)
    topic = (data.get("topic") or "").strip()
    if not topic:
        return jsonify({"error": "主题不能为空"}), 400
    theme = data.get("theme") if data.get("theme") in builder.THEMES else "blue"
    with lock:
        if state["phase"] not in ("idle", "ready"):
            return jsonify({"error": "正在生成中，请等待完成"}), 409
        state["theme"] = theme
    os.makedirs(IMAGES_DIR, exist_ok=True)
    threading.Thread(target=_generation_worker, args=(topic,), daemon=True).start()
    return jsonify({"ok": True})


@app.route("/api/status")
def api_status():
    with lock:
        return jsonify(state)


@app.route("/api/slide/<int:i>/text", methods=["POST"])
def api_edit_text(i):
    data = request.get_json(force=True)
    with lock:
        if i < 0 or i >= len(state["slides"]):
            return jsonify({"error": "页码不存在"}), 404
        state["slides"][i]["title"] = data.get("title", state["slides"][i]["title"])
        pts = data.get("points")
        if isinstance(pts, list):
            state["slides"][i]["points"] = [str(p).strip() for p in pts if str(p).strip()]
    return jsonify({"ok": True})


@app.route("/api/slide/<int:i>/image", methods=["POST"])
def api_regen_image(i):
    data = request.get_json(force=True) if request.data else {}
    with lock:
        if i < 0 or i >= len(state["slides"]):
            return jsonify({"error": "页码不存在"}), 404
        if state["slides"][i]["imageStatus"] == "generating":
            return jsonify({"error": "该页正在生成中"}), 409
        prompt = (data.get("prompt") or "").strip() or state["slides"][i]["image_prompt"]
        state["slides"][i]["imageStatus"] = "generating"
        state["slides"][i]["image"] = None

    def worker():
        path = os.path.join(IMAGES_DIR, f"slide_{i}.png")
        try:
            image_gen.generate_image(prompt, path)
            with lock:
                state["slides"][i]["image"] = f"/images/slide_{i}.png"
                state["slides"][i]["imageStatus"] = "done"
                state["slides"][i]["image_prompt"] = prompt
            _log(f"页 {i + 1} 图片已更新")
        except Exception as e:
            with lock:
                state["slides"][i]["imageStatus"] = "failed"
            _log(f"页 {i + 1} 图片生成失败：{e}")

    threading.Thread(target=worker, daemon=True).start()
    return jsonify({"ok": True})


@app.route("/api/export", methods=["POST"])
def api_export():
    with lock:
        phase = state["phase"]
        slides = [dict(s) for s in state["slides"]]
        topic = state["topic"]
        theme = state["theme"]
    if phase != "ready":
        return jsonify({"error": "生成尚未完成，无法导出"}), 409
    if not slides:
        return jsonify({"error": "没有可导出的页面"}), 400

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    safe_topic = re.sub(r'[\\/:*?"<>| ]', "_", topic or "ppt")[:20]
    out_path = os.path.join(OUTPUT_DIR, f"{safe_topic}_{time.strftime('%Y%m%d_%H%M%S')}.pptx")
    builder.build_ppt(
        [{"type": s.get("type", "content"), "title": s["title"], "points": s["points"],
          "image_prompt": s["image_prompt"], "chart": s.get("chart")} for s in slides],
        [os.path.join(IMAGES_DIR, os.path.basename(s["image"])) if s["image"] else None for s in slides],
        out_path,
        theme=theme,
        subtitle=topic,
    )
    _log(f"已导出：{os.path.basename(out_path)}")
    return jsonify({"ok": True, "path": out_path})


# HTML 演示版色板/结构来自 ppt-maker skill（~/.agents/skills/ppt-maker）
_HTML_THEMES = {
    "blue":  {"bg": "#0f2a4a", "accent": "#3b82f6", "fg": "#ffffff", "muted": "#b6c7dc"},
    "dark":  {"bg": "#111827", "accent": "#f59e0b", "fg": "#f9fafb", "muted": "#9ca3af"},
    "green": {"bg": "#062e21", "accent": "#22c55e", "fg": "#ecfdf5", "muted": "#a7c4b6"},
}


def _html_escape(s: str) -> str:
    return (s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def _deck_bars(chart) -> str:
    labels = chart.get("labels", [])
    values = chart.get("values", [])
    mx = max(values) if values else 1
    return '<div class="chart">' + "".join(
        f'<div class="bar-row"><span class="bar-label">{_html_escape(labels[i])}</span>'
        f'<span class="bar-track"><span class="bar" style="width:{values[i] / mx * 100:.0f}%"></span></span>'
        f'<span class="bar-val">{values[i]}</span></div>'
        for i in range(min(len(labels), len(values)))) + "</div>"


def _build_html_deck(topic: str, slides: list[dict], theme: str, out_path: str) -> str:
    """把当前页面状态导出为零依赖单文件 HTML 幻灯片（含图片，按版式区分）。"""
    t = _HTML_THEMES.get(theme, _HTML_THEMES["blue"])
    sections = []
    for i, s in enumerate(slides):
        title = _html_escape(s["title"])
        stype = s.get("type", "content")
        points = s.get("points", [])
        if stype == "cover" or (i == 0 and not points):
            body = (f'<section class="slide cover"><h1 class="title">{title}</h1>'
                    f'<p class="sub">{_html_escape(topic)}</p></section>')
        elif stype == "toc":
            rows = "".join(f'<li class="toc-item">{_html_escape(p)}</li>' for p in points)
            body = f'<section class="slide"><h2 class="h">{title}</h2><ul class="toc">{rows}</ul></section>'
        elif stype == "section":
            brief = f'<p class="lead">{_html_escape(points[0])}</p>' if points else ""
            body = f'<section class="slide section"><h1 class="sect-title">{title}</h1>{brief}</section>'
        elif stype == "data" and s.get("chart"):
            body = f'<section class="slide"><h2 class="h">{title}</h2>{_deck_bars(s["chart"])}</section>'
        elif stype == "end":
            body = f'<section class="slide end"><h1 class="center">{title}</h1></section>'
        else:
            lis = "".join(f"<li>{_html_escape(p)}</li>" for p in points)
            img = f'<img class="deck-img" src="{s["image"].lstrip("/")}" alt="">' if s.get("image") else ""
            row = f'<div class="row"><ul>{lis}</ul>{img}</div>' if img else f'<ul>{lis}</ul>'
            body = f'<section class="slide"><h2 class="h">{title}</h2>{row}</section>'
        sections.append(body)

    doc = f"""<!DOCTYPE html><html lang="zh"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{_html_escape(topic)}</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:"Microsoft YaHei","PingFang SC",system-ui,sans-serif;background:{t['bg']};color:{t['fg']}}}
.slide{{min-height:100vh;display:flex;flex-direction:column;justify-content:center;padding:7vh 8vw}}
.slide.cover{{align-items:center;text-align:center}}
.slide.cover .title{{font-size:clamp(2.2rem,6vw,4.2rem);font-weight:800}}
.slide.cover .sub{{font-size:clamp(1.1rem,2.4vw,1.6rem);color:{t['muted']};margin-top:1.2rem}}
.h{{font-size:clamp(1.6rem,3.6vw,2.6rem);font-weight:800;margin-bottom:2rem;border-left:6px solid {t['accent']};padding-left:1rem}}
.row{{display:flex;gap:3rem;align-items:center}}
ul{{list-style:none}}
li{{font-size:clamp(1rem,2vw,1.35rem);line-height:1.7;padding:.55rem 0 .55rem 2rem;position:relative}}
li::before{{content:"▸";position:absolute;left:0;color:{t['accent']}}}
li.toc-item{{font-size:clamp(1.1rem,2.4vw,1.6rem);padding:.8rem 0 .8rem 2rem;border-bottom:1px solid rgba(255,255,255,.1)}}
.sect-title{{font-size:clamp(2rem,5vw,3.4rem);font-weight:800;text-align:center;border-bottom:4px solid {t['accent']};padding-bottom:1.4rem}}
.section .lead{{font-size:1.2rem;color:{t['muted']};text-align:center;margin-top:1.6rem}}
.chart{{display:flex;flex-direction:column;gap:1.1rem;margin-top:1rem}}
.bar-row{{display:flex;align-items:center;gap:.8rem;font-size:1.1rem}}
.bar-label{{width:8rem;text-align:right;color:{t['muted']};flex-shrink:0}}
.bar-track{{flex:1;background:rgba(255,255,255,.12);border-radius:8px;height:1.6rem;overflow:hidden}}
.bar{{display:block;height:100%;background:linear-gradient(90deg,{t['accent']},{t['accent']}cc);border-radius:8px}}
.bar-val{{width:3rem;font-weight:700}}
.center{{text-align:center;font-weight:800;font-size:clamp(1.8rem,4vw,3rem)}}
.slide.end .center{{color:{t['accent']}}}
.deck-img{{flex:1;max-width:40vw;max-height:70vh;border-radius:10px;box-shadow:0 4px 24px rgba(0,0,0,.35)}}
@media print{{.slide{{min-height:100vh;page-break-after:always}}}}
</style></head><body>{''.join(sections)}</body></html>"""

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(doc)
    return out_path


@app.route("/api/export_html", methods=["POST"])
def api_export_html():
    with lock:
        phase = state["phase"]
        slides = [dict(s) for s in state["slides"]]
        topic = state["topic"]
        theme = state["theme"]
    if phase != "ready":
        return jsonify({"error": "生成尚未完成，无法导出"}), 409
    if not slides:
        return jsonify({"error": "没有可导出的页面"}), 400

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    safe_topic = re.sub(r'[\\/:*?"<>| ]', "_", topic or "ppt")[:20]
    out_path = os.path.join(OUTPUT_DIR, f"{safe_topic}_{time.strftime('%Y%m%d_%H%M%S')}.html")
    _build_html_deck(topic, slides, theme, out_path)
    _log(f"已导出 HTML 演示版：{os.path.basename(out_path)}")
    return jsonify({"ok": True, "path": out_path})


@app.route("/images/<path:filename>")
def images(filename):
    return send_from_directory(IMAGES_DIR, filename)


if __name__ == "__main__":
    os.makedirs(IMAGES_DIR, exist_ok=True)
    app.run(host="127.0.0.1", port=5000, debug=False)
