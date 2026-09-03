"""可视化制作 Web 界面后端：python app.py 后浏览器打开 http://127.0.0.1:5000"""

import os
import re
import threading
import time

from flask import Flask, jsonify, render_template, request, send_from_directory

import builder
import critic
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


def _parse_upload(file) -> str:
    """解析上传文件（docx/md/txt），返回纯文本。"""
    filename = (file.filename or "").lower()
    if filename.endswith(".docx"):
        import docx
        from io import BytesIO
        d = docx.Document(BytesIO(file.read()))
        return "\n".join(p.text for p in d.paragraphs if p.text.strip())
    return file.read().decode("utf-8", errors="ignore")


def _start_generation(content: str, from_text: bool, theme: str) -> bool:
    """锁内初始化 state 并启动后台 worker，避免 phase 置位竞态。"""
    with lock:
        if state["phase"] not in ("idle", "ready"):
            return False
        state["topic"] = content
        state["theme"] = theme
        state["slides"] = []
        state["log"] = []
        state["phase"] = "outline"
    os.makedirs(IMAGES_DIR, exist_ok=True)
    threading.Thread(target=_generation_worker, args=(content, from_text), daemon=True).start()
    return True


def _generation_worker(content: str, from_text: bool = False):
    try:
        if from_text:
            _log("从文档提炼大纲…")
            slides = outline.generate_outline_from_text(content)
            if slides:
                with lock:
                    state["topic"] = slides[0].get("title", "") or content[:20]
        else:
            _log(f"生成大纲：{content}")
            slides = outline.generate_outline(content)
        with lock:
            state["slides"] = [
                {"type": s.get("type", "content"), "title": s.get("title", ""),
                 "points": s.get("points", []), "image_prompt": s.get("image_prompt", ""),
                 "chart": s.get("chart"), "layout": s.get("layout"),
                 "image": None, "imageStatus": "pending",
                 "review": {"ok": True, "reason": "", "tries": 0}}
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
            review = {"ok": True, "reason": "", "tries": 0}
            # 提议者-审核者闭环：生图 → vision 校验 → 不契合则改提示词重生（封顶 2 次）
            for attempt in range(3):
                try:
                    image_gen.generate_image(prompt, path)
                    with lock:
                        state["slides"][i]["image"] = f"/images/slide_{i}.png"
                        state["slides"][i]["imageStatus"] = "done"
                    rv = critic.review_image(s["title"], s.get("points", []), path)
                    review = {"ok": rv["ok"], "reason": rv["reason"], "tries": attempt}
                    with lock:
                        state["slides"][i]["review"] = review
                    if rv["ok"]:
                        _log(f"页 {i + 1} 图片完成（校验契合）")
                        break
                    if attempt < 2:
                        if rv["advice"]:
                            prompt = f"{s.get('image_prompt', prompt)}。注意：{rv['advice']}"
                        _log(f"页 {i + 1} 校验不契合，重生（第 {attempt + 1} 次）：{rv['reason'][:40]}")
                    else:
                        _log(f"页 {i + 1} 已重生 2 次仍未契合，保留当前图")
                except Exception as e:
                    with lock:
                        state["slides"][i]["imageStatus"] = "failed"
                        state["slides"][i]["review"] = review
                    _log(f"页 {i + 1} 图片生成失败：{e}")
                    break
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
    if not isinstance(topic, str) or not topic:
        return jsonify({"error": "主题不能为空"}), 400
    theme = data.get("theme") if data.get("theme") in builder.THEMES else "blue"
    if not _start_generation(topic, False, theme):
        return jsonify({"error": "正在生成中，请等待完成"}), 409
    return jsonify({"ok": True})


@app.route("/api/import", methods=["POST"])
def api_import():
    data = request.get_json(force=True)
    text = (data.get("text") or "").strip()
    if not isinstance(text, str) or len(text) < 30:
        return jsonify({"error": "文档内容过短，请提供更完整的文档"}), 400
    theme = data.get("theme") if data.get("theme") in builder.THEMES else "blue"
    if not _start_generation(text, True, theme):
        return jsonify({"error": "正在生成中，请等待完成"}), 409
    return jsonify({"ok": True})


@app.route("/api/import_file", methods=["POST"])
def api_import_file():
    f = request.files.get("file")
    if not f:
        return jsonify({"error": "未收到文件"}), 400
    try:
        text = _parse_upload(f)
    except Exception as e:
        return jsonify({"error": f"文件解析失败：{e}"}), 400
    if len(text.strip()) < 30:
        return jsonify({"error": "文档内容过短"}), 400
    theme = request.form.get("theme") if request.form.get("theme") in builder.THEMES else "blue"
    if not _start_generation(text.strip(), True, theme):
        return jsonify({"error": "正在生成中，请等待完成"}), 409
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


@app.route("/api/refine", methods=["POST"])
def api_refine():
    data = request.get_json(force=True)
    instruction = (data.get("instruction") or "").strip()
    if not instruction:
        return jsonify({"error": "指令不能为空"}), 400
    with lock:
        if state["phase"] not in ("ready",):
            return jsonify({"error": "请先生成 PPT，再进行对话修改"}), 409
        cur_slides = [dict(s) for s in state["slides"]]
        state["phase"] = "refining"

    def worker():
        try:
            _log(f"对话修改：{instruction}")
            new_slides = critic.refine_outline(cur_slides, instruction)
            with lock:
                state["slides"] = [
                    {"type": s.get("type", "content"), "title": s.get("title", ""),
                     "points": s.get("points", []), "image_prompt": s.get("image_prompt", ""),
                     "chart": s.get("chart"), "layout": s.get("layout"),
                     "image": None, "imageStatus": "pending",
                     "review": {"ok": True, "reason": "", "tries": 0}}
                    for s in new_slides
                ]
            _log(f"对话修改完成，共 {len(new_slides)} 页，重新生图")
            state["phase"] = "images"
            # 重新走生图 + 视觉校验
            for i, s in enumerate(new_slides):
                prompt = s.get("image_prompt", "")
                if not prompt:
                    with lock:
                        state["slides"][i]["imageStatus"] = "skipped"
                    continue
                path = os.path.join(IMAGES_DIR, f"slide_{i}.png")
                review = {"ok": True, "reason": "", "tries": 0}
                for attempt in range(3):
                    try:
                        image_gen.generate_image(prompt, path)
                        with lock:
                            state["slides"][i]["image"] = f"/images/slide_{i}.png"
                            state["slides"][i]["imageStatus"] = "done"
                        rv = critic.review_image(s["title"], s.get("points", []), path)
                        review = {"ok": rv["ok"], "reason": rv["reason"], "tries": attempt}
                        with lock:
                            state["slides"][i]["review"] = review
                        if rv["ok"]:
                            _log(f"页 {i + 1} 图片完成（校验契合）")
                            break
                        if attempt < 2:
                            if rv["advice"]:
                                prompt = f"{s.get('image_prompt', prompt)}。注意：{rv['advice']}"
                            _log(f"页 {i + 1} 校验不契合，重生（第 {attempt + 1} 次）")
                        else:
                            _log(f"页 {i + 1} 已重生 2 次仍未契合，保留")
                    except Exception as e:
                        with lock:
                            state["slides"][i]["imageStatus"] = "failed"
                            state["slides"][i]["review"] = review
                        _log(f"页 {i + 1} 图片生成失败：{e}")
                        break
            _log("对话修改完成，可导出")
            state["phase"] = "ready"
        except Exception as e:
            _log(f"对话修改失败：{e}")
            state["phase"] = "ready"

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
          "image_prompt": s["image_prompt"], "chart": s.get("chart"), "layout": s.get("layout")} for s in slides],
        [os.path.join(IMAGES_DIR, os.path.basename(s["image"])) if s["image"] else None for s in slides],
        out_path,
        theme=theme,
        subtitle=topic,
    )
    _log(f"已导出：{os.path.basename(out_path)}")
    return jsonify({"ok": True, "path": out_path})


@app.route("/api/export_txt", methods=["POST"])
def api_export_txt():
    with lock:
        phase = state["phase"]
        slides = [dict(s) for s in state["slides"]]
        topic = state["topic"]
    if phase != "ready" or not slides:
        return jsonify({"error": "生成尚未完成，无法导出"}), 409
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    safe_topic = re.sub(r'[\\/:*?"<>| ]', "_", topic or "ppt")[:20]
    out_path = os.path.join(OUTPUT_DIR, f"{safe_topic}_{time.strftime('%Y%m%d_%H%M%S')}_大纲.txt")
    lines = [topic, "=" * 40]
    for i, s in enumerate(slides):
        lines.append(f"\n[{i + 1}] {s['title']}")
        for p in s.get("points", []):
            lines.append(f"  - {p}")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    _log(f"已导出大纲：{os.path.basename(out_path)}")
    return jsonify({"ok": True, "path": out_path})


def _export_pdf_via_com(pptx_path: str, pdf_path: str) -> bool:
    """用 PowerPoint COM 把 pptx 转 PDF，成功返回 True。"""
    try:
        import win32com.client  # noqa
    except ImportError:
        return False
    import pythoncom
    try:
        pythoncom.CoInitialize()
        app = win32com.client.Dispatch("PowerPoint.Application")
        pres = app.Presentations.Open(pptx_path, WithWindow=False)
        pres.SaveAs(pdf_path, 32)  # 32 = ppSaveAsPDF
        pres.Close()
        app.Quit()
        return os.path.exists(pdf_path)
    except Exception:
        return False


@app.route("/api/export_pdf", methods=["POST"])
def api_export_pdf():
    with lock:
        phase = state["phase"]
        slides = [dict(s) for s in state["slides"]]
        topic = state["topic"]
        theme = state["theme"]
    if phase != "ready" or not slides:
        return jsonify({"error": "生成尚未完成，无法导出"}), 409
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    safe_topic = re.sub(r'[\\/:*?"<>| ]', "_", topic or "ppt")[:20]
    stamp = time.strftime('%Y%m%d_%H%M%S')
    pptx_path = os.path.join(OUTPUT_DIR, f"{safe_topic}_{stamp}.pptx")
    pdf_path = os.path.join(OUTPUT_DIR, f"{safe_topic}_{stamp}.pdf")
    builder.build_ppt(
        [{"type": s.get("type", "content"), "title": s["title"], "points": s["points"],
          "image_prompt": s["image_prompt"], "chart": s.get("chart"), "layout": s.get("layout")} for s in slides],
        [os.path.join(IMAGES_DIR, os.path.basename(s["image"])) if s["image"] else None for s in slides],
        pptx_path, theme=theme, subtitle=topic,
    )
    if _export_pdf_via_com(pptx_path, pdf_path):
        _log(f"已导出 PDF：{os.path.basename(pdf_path)}")
        return jsonify({"ok": True, "path": pdf_path})
    return jsonify({"ok": False, "error": "PDF 导出失败（需安装 pywin32 与 PowerPoint），可改用「导出 HTML」后用浏览器打印为 PDF"}), 500


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
