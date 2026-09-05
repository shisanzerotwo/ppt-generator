"""可视化制作 Web 界面后端：python app.py 后浏览器打开 http://127.0.0.1:5000"""

import json
import os
import re
import threading
import time
from urllib.parse import quote

from flask import Flask, jsonify, render_template, request, send_from_directory

import builder
import critic
import html_gen
import image_gen
import outline
import style as style_mod

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
IMAGES_DIR = os.path.join(OUTPUT_DIR, "images")
DECKS_DIR = os.path.join(OUTPUT_DIR, "decks")
PROJECTS_DIR = os.path.join(OUTPUT_DIR, "projects")

state = {
    "topic": "",
    "phase": "idle",  # idle / outline / images / designing / ready
    "theme": "blue",  # 卡片区底色（固定值，视觉风格由 AI 按主题选定）
    "style": None,    # AI 选定的风格 key（style.STYLE_LIBRARY）
    "style_name": "", # AI 选定风格名（前端展示）
    "brand": None,    # 品牌模板 {name, color}，用户自定义后覆盖风格强调色（路线图 #6）
    "slides": [],     # {title, points, image_prompt, image, imageStatus}
    "html_path": None,  # AI 自主设计的 HTML 主产物（web 路径 /decks/xxx.html）
    "await_step": None,  # 分步确认时正在等待放行的步骤："outline" / "design" / None
    "stepwise": False,   # 是否在关键节点暂停等用户审阅（默认关，用户按需开启）
    "log": [],
}
lock = threading.Lock()
resume_event = threading.Event()   # 分步确认的放行信号


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


def _start_generation(content: str, from_text: bool) -> bool:
    """锁内初始化 state 并启动后台 worker，避免 phase 置位竞态。"""
    with lock:
        if state["phase"] not in ("idle", "ready"):
            return False
        state["topic"] = content
        state["slides"] = []
        state["style"] = None
        state["style_name"] = ""
        # 注意：brand 不在此重置——它是用户级偏好，设过就跨生成生效，直到手动清除
        state["html_path"] = None
        state["log"] = []
        state["await_step"] = None
        state["phase"] = "outline"
    resume_event.clear()   # 防上一轮残留的放行信号让本次暂停被瞬间跳过
    os.makedirs(IMAGES_DIR, exist_ok=True)
    threading.Thread(target=_generation_worker, args=(content, from_text), daemon=True).start()
    return True


def _apply_brand(chosen_style, brand):
    """品牌模板（路线图 #6）：品牌主色覆盖 AI 风格强调色，并注入品牌指引。"""
    base = dict(chosen_style or style_mod.STYLE_LIBRARY["fresh-light"])
    base["name"] = f"{brand['name']}·品牌定制"
    base["accent"] = brand["color"]
    base["guidance"] = (base.get("guidance", "") +
                        f" 品牌名为「{brand['name']}」，品牌主色 {brand['color']} 必须作为全篇强调色与关键装饰色。")
    return base


def _save_project_snapshot():
    """路线图 #8：ready 后把 slides 序列化到 projects/，供历史回看与载入复用。"""
    try:
        with lock:
            payload = {
                "topic": state["topic"], "style": state.get("style"),
                "style_name": state.get("style_name"), "brand": state.get("brand"),
                "slides": [dict(s) for s in state["slides"]],
                "html_path": state.get("html_path"),
                "saved_at": time.strftime("%Y-%m-%d %H:%M:%S"),
            }
            topic = state["topic"]
        os.makedirs(PROJECTS_DIR, exist_ok=True)
        safe_topic = re.sub(r'[\\/:*?"<>| ]', "_", topic or "ppt")[:20]
        path = os.path.join(PROJECTS_DIR, f"{safe_topic}_{time.strftime('%Y%m%d_%H%M%S')}.json")
        with open(path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False)
    except Exception as e:
        _log(f"项目快照保存失败：{e}")


def _design_and_save():
    """designing 阶段：调 LLM 自主设计 HTML 并保存，失败仅记日志不阻断。"""
    with lock:
        slides = [dict(s) for s in state["slides"]]
        topic = state["topic"]
        chosen_style = style_mod.STYLE_LIBRARY.get(state.get("style"))
        brand = state.get("brand")
    if brand:
        chosen_style = _apply_brand(chosen_style, brand)
    _phase("designing")
    _log("AI 正在自主设计 HTML 幻灯片…")
    image_map = {i: f"../images/slide_{i}.png" for i, s in enumerate(slides) if s.get("image")}
    try:
        doc = html_gen.generate_html_deck(topic, slides, image_map, chosen_style)
        os.makedirs(DECKS_DIR, exist_ok=True)
        safe_topic = re.sub(r'[\\/:*?"<>| ]', "_", topic or "ppt")[:20]
        out_path = os.path.join(DECKS_DIR, f"{safe_topic}_{time.strftime('%Y%m%d_%H%M%S')}.html")
        tmp_path = out_path + ".tmp"
        with open(tmp_path, "w", encoding="utf-8") as f:
            f.write(doc)
        os.replace(tmp_path, out_path)  # 原子落盘，避免 /api/decks 列出半截文件（审计 P2-2）
        with lock:
            # basename 做 URL 编码：主题含引号/反引号等字符时不会被前端 onclick 拼接执行（审计 C1）
            state["html_path"] = f"/decks/{quote(os.path.basename(out_path))}"
        _save_project_snapshot()
        _log(f"HTML 设计完成：{os.path.basename(out_path)}")
        return True
    except Exception as e:
        _log(f"HTML 设计失败：{e}")
        return False


REVIEW_TIMEOUT = 1800  # 分步确认最长等待，超时自动放行，避免 worker 永久挂起


def _pause_gate(step: str, label: str):
    """分步确认闸门：stepwise 开启时暂停 worker，等前端 /api/continue 放行。

    超时或异常一律放行，保证生成流程不会因等待而卡死。
    """
    with lock:
        if not state["stepwise"]:
            return
        state["phase"] = "review"
        state["await_step"] = step
    _log(f"{label}已完成，等你确认后再继续（可先用 AI 协作面板提修改意见）")
    try:
        resumed = resume_event.wait(timeout=REVIEW_TIMEOUT)
        if not resumed:
            _log("确认等待超时，自动继续")
    finally:
        resume_event.clear()
        with lock:
            state["await_step"] = None
            state["phase"] = "images" if step == "outline" else "designing"


def _gen_images(indices=None):
    """逐页生图 + 视觉校验闭环（提议者-审核者，不契合改词重生封顶 2 次）。

    indices 为 None 时处理全部页；定点修改时只传被改动的页，避免整篇重生图。
    """
    with lock:
        n = len(state["slides"])
        todo = list(range(n)) if indices is None else [i for i in indices if 0 <= i < n]
    for i in todo:
        with lock:
            s = dict(state["slides"][i])
        prompt = s.get("image_prompt", "")
        if not prompt:
            with lock:
                state["slides"][i]["imageStatus"] = "skipped"
            _log(f"页 {i + 1}：无配图提示词，跳过")
            continue
        path = os.path.join(IMAGES_DIR, f"slide_{i}.png")
        review = {"ok": True, "reason": "", "tries": 0}
        for attempt in range(3):
            try:
                image_gen.generate_image(prompt, path)
                with lock:
                    state["slides"][i]["image"] = f"/images/slide_{i}.png"
                    state["slides"][i]["imageStatus"] = "done"
                rv = critic.review_image(s.get("title", ""), s.get("points", []), path)
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
        _log(f"大纲完成，共 {len(slides)} 页，AI 正在判断视觉风格")
        # 智能选风格：AI 根据主题/内容从风格库中自主判断（ppt-maker skill 规范）
        with lock:
            cur_topic = state["topic"]
        chosen = style_mod.decide_style(cur_topic, slides)
        with lock:
            state["style"] = chosen.get("key")
            state["style_name"] = chosen.get("name", "")
        reason = f"（{chosen['reason']}）" if chosen.get("reason") else ""
        _log(f"AI 选定风格：{chosen.get('name')} {reason}")
        _pause_gate("outline", "大纲与风格")
        _log("开始逐页生图")
        _phase("images")

        _gen_images()
        _pause_gate("design", "配图")
        _design_and_save()
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


@app.route("/api/continue", methods=["POST"])
def api_continue():
    """放行分步确认。仅在 review 态有效，其余情况幂等返回 409。"""
    with lock:
        if state["phase"] != "review":
            return jsonify({"error": "当前没有待确认的步骤"}), 409
        step = state["await_step"]
    if step:
        resume_event.set()
    return jsonify({"ok": True, "resumed": step})


@app.route("/api/stepwise", methods=["POST"])
def api_stepwise():
    """开关分步确认（生成过程是否在关键节点暂停等你放行）。"""
    data = request.get_json(force=True)
    enabled = bool(data.get("enabled"))
    with lock:
        state["stepwise"] = enabled
    _log(f"分步确认已{'开启' if enabled else '关闭'}")
    return jsonify({"ok": True, "stepwise": enabled})


@app.route("/api/generate", methods=["POST"])
def api_generate():
    data = request.get_json(force=True)
    topic = (data.get("topic") or "").strip()
    if not isinstance(topic, str) or not topic:
        return jsonify({"error": "主题不能为空"}), 400
    if not _start_generation(topic, False):
        return jsonify({"error": "正在生成中，请等待完成"}), 409
    return jsonify({"ok": True})


@app.route("/api/import", methods=["POST"])
def api_import():
    data = request.get_json(force=True)
    text = (data.get("text") or "").strip()
    if not isinstance(text, str) or len(text) < 30:
        return jsonify({"error": "文档内容过短，请提供更完整的文档"}), 400
    if len(text.strip()) < 30:
        return jsonify({"error": "文档内容过短，请提供更完整的文档"}), 400
    if not _start_generation(text.strip(), True):
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
    if not _start_generation(text.strip(), True):
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
        # 主流程生图中不允许换图：两者会并发写同一个 slide_<i>.png（审计发现）
        if state["phase"] not in ("ready", "review"):
            return jsonify({"error": "正在生成中，请等待完成后再换图"}), 409
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
    target = data.get("target")
    if target is not None and not isinstance(target, dict):
        return jsonify({"error": "target 需为对象 {slide, quote}"}), 400
    with lock:
        # 刻意不放行 review 态：refine 会另起 worker 重跑生图与设计，
        # 而原 worker 仍阻塞在闸门上，醒来会重复一遍（竞态 + 白烧额度）。
        # 暂停期请直接编辑卡片（saveText / 换图不触发流程），改完点「继续」即生效。
        if state["phase"] != "ready":
            return jsonify({"error": "请先生成 PPT，再进行对话修改"}), 409
        cur_slides = [dict(s) for s in state["slides"]]
        # 输入类错误在 HTTP 层即时判掉：否则只进日志，前端会把失败当成改写成功
        if target is not None:
            idx = target.get("slide")
            if isinstance(idx, bool) or not isinstance(idx, int) or not 0 <= idx < len(cur_slides):
                return jsonify({"error": "目标页码不存在"}), 400
            q = str(target.get("quote") or "").strip()
            if q and not any(q in p for p in (cur_slides[idx].get("points") or [])):
                return jsonify({"error": "该页未找到选中的文本，可能已被改动"}), 400
        state["phase"] = "refining"

    def worker():
        try:
            if target is not None:
                # 定点修改：只动被圈定的页/要点，其余页连图一起保留，不打回重做
                _log(f"定点修改（第 {int(target.get('slide', -1)) + 1} 页）：{instruction}")
                new_slides, info = critic.revise_target(cur_slides, target, instruction)
                idx = info["slide"]
                if info["level"] == "slide":
                    # critic 只产出内容字段，状态字段（image/imageStatus/review）沿用原页
                    content = {k: v for k, v in new_slides[idx].items()
                               if k in ("title", "points", "image_prompt", "chart")}
                    new_slides[idx] = {**cur_slides[idx], **content}
                with lock:
                    state["slides"] = new_slides
                    if info["regen_image"]:
                        state["slides"][idx].update(
                            {"image": None, "imageStatus": "pending",
                             "review": {"ok": True, "reason": "", "tries": 0}})
                _log(f"已改写第 {idx + 1} 页的"
                     + ("选中要点" if info["level"] == "point" else "整页内容"))
                if info["regen_image"]:
                    _phase("images")
                    _gen_images([idx])
                _design_and_save()
                _log("定点修改完成，可导出")
                _phase("ready")
                return

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
            _phase("images")
            _gen_images()
            _design_and_save()
            _log("对话修改完成，可导出")
            _phase("ready")
        except Exception as e:
            _log(f"对话修改失败：{e}")
            _phase("ready")

    threading.Thread(target=worker, daemon=True).start()
    return jsonify({"ok": True})


@app.route("/api/export", methods=["POST"])
def api_export():
    with lock:
        phase = state["phase"]
        slides = [dict(s) for s in state["slides"]]
        topic = state["topic"]
        theme = style_mod.THEME_MAP.get(state.get("style"), state["theme"])
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
        theme = style_mod.THEME_MAP.get(state.get("style"), state["theme"])
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
        theme = style_mod.THEME_MAP.get(state.get("style"), state["theme"])
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


@app.route("/api/redesign", methods=["POST"])
def api_redesign():
    """编辑页面后重新触发 AI 设计（重生成 HTML 主产物）。"""
    with lock:
        if state["phase"] not in ("ready",):
            return jsonify({"error": "请等待生成完成再重新设计"}), 409
        if not state["slides"]:
            return jsonify({"error": "没有可设计的页面"}), 400
        state["phase"] = "designing"

    def worker():
        try:
            _design_and_save()
        finally:
            state["phase"] = "ready"

    threading.Thread(target=worker, daemon=True).start()
    return jsonify({"ok": True})


@app.route("/api/slide/reorder", methods=["POST"])
def api_slide_reorder():
    """页面重排序（路线图 #7）。order 为新顺序的下标排列。"""
    data = request.get_json(force=True)
    order = data.get("order")
    with lock:
        if state["phase"] != "ready":
            return jsonify({"error": "请等待生成完成再调整页面"}), 409
        n = len(state["slides"])
        if not isinstance(order, list) or sorted(order) != list(range(n)):
            return jsonify({"error": "order 必须是 0..n-1 的完整排列"}), 400
        state["slides"] = [state["slides"][i] for i in order]
    return jsonify({"ok": True})


@app.route("/api/slide/add", methods=["POST"])
def api_slide_add():
    """在 after 下标后插入一页空白 content 页（路线图 #7）。"""
    data = request.get_json(force=True) if request.data else {}
    with lock:
        if state["phase"] != "ready":
            return jsonify({"error": "请等待生成完成再加页"}), 409
        try:
            i = len(state["slides"]) - 1 if data.get("after") is None else int(data["after"])
        except (TypeError, ValueError):
            return jsonify({"error": "after 必须是页码"}), 400
        if i < -1 or i >= len(state["slides"]):
            return jsonify({"error": "页码不存在"}), 404
        state["slides"].insert(i + 1, {
            "type": "content", "title": "新页面", "points": [], "image_prompt": "",
            "chart": None, "layout": None, "image": None, "imageStatus": "skipped",
            "review": {"ok": True, "reason": "", "tries": 0}})
    return jsonify({"ok": True, "index": i + 1})


@app.route("/api/slide/<int:i>/delete", methods=["DELETE"])
def api_slide_delete(i):
    """删除一页（路线图 #7）。"""
    with lock:
        if state["phase"] != "ready":
            return jsonify({"error": "请等待生成完成再删除页面"}), 409
        if i < 0 or i >= len(state["slides"]):
            return jsonify({"error": "页码不存在"}), 404
        if len(state["slides"]) <= 1:
            return jsonify({"error": "至少要保留一页"}), 400
        state["slides"].pop(i)
    return jsonify({"ok": True})


@app.route("/api/projects")
def api_projects():
    """历史项目列表（路线图 #8，按时间倒序，最多 30 条）。"""
    if not os.path.isdir(PROJECTS_DIR):
        return jsonify({"projects": []})
    items = []
    for name in os.listdir(PROJECTS_DIR):
        if not name.lower().endswith(".json"):
            continue
        full = os.path.join(PROJECTS_DIR, name)
        try:
            mtime = os.path.getmtime(full)
        except OSError:
            continue
        stem = name[:-5]
        parts = stem.rsplit("_", 2)
        title = parts[0].replace("_", " ") if len(parts) == 3 else stem
        items.append({"title": title, "name": name, "mtime": mtime})
    items.sort(key=lambda x: x["mtime"], reverse=True)
    for it in items:
        it["when"] = time.strftime("%m-%d %H:%M", time.localtime(it.pop("mtime")))
    return jsonify({"projects": items[:30]})


@app.route("/api/projects/load", methods=["POST"])
def api_projects_load():
    """载入历史项目：恢复 topic/slides/style 与设计稿链接。"""
    data = request.get_json(force=True)
    name = (data.get("name") or "").strip()
    if (not name or "/" in name or "\\" in name or ".." in name
            or not name.endswith(".json")):
        return jsonify({"error": "非法项目名"}), 400
    full = os.path.join(PROJECTS_DIR, name)
    if not os.path.isfile(full):
        return jsonify({"error": "项目不存在"}), 404
    try:
        with open(full, encoding="utf-8") as f:
            payload = json.load(f)
    except (OSError, ValueError):
        return jsonify({"error": "项目文件损坏"}), 500
    with lock:
        if state["phase"] not in ("idle", "ready"):
            return jsonify({"error": "正在生成中，无法载入"}), 409
        html_path = payload.get("html_path")
        if html_path and not os.path.isfile(
                os.path.join(DECKS_DIR, os.path.basename(html_path))):
            html_path = None  # 设计稿已不在，可重新设计
        state.update({
            "topic": payload.get("topic", ""), "style": payload.get("style"),
            "style_name": payload.get("style_name", ""), "brand": payload.get("brand"),
            "slides": payload.get("slides", []), "html_path": html_path,
            "phase": "ready"})
    _log(f"已载入历史项目：{name}")
    return jsonify({"ok": True})


@app.route("/api/brand", methods=["POST"])
def api_brand():
    """品牌模板（路线图 #6）：设置/清除品牌名与品牌主色。"""
    data = request.get_json(force=True) if request.data else {}
    name = (data.get("name") or "").strip()
    color = (data.get("color") or "").strip()
    if not name and not color:
        with lock:
            state["brand"] = None
        return jsonify({"ok": True, "brand": None})
    if not re.fullmatch(r"#[0-9a-fA-F]{6}", color):
        return jsonify({"error": "颜色需为 #RRGGBB 格式"}), 400
    brand = {"name": name[:20] or "品牌", "color": color}
    with lock:
        state["brand"] = brand
    return jsonify({"ok": True, "brand": brand})


@app.route("/images/<path:filename>")
def images(filename):
    return send_from_directory(IMAGES_DIR, filename)


@app.route("/decks/<path:filename>")
def decks(filename):
    return send_from_directory(DECKS_DIR, filename)


ARTIFACT_EXT = (".pptx", ".pdf", ".txt")


@app.route("/api/artifacts")
def api_artifacts():
    """列出可打开/下载的产物：output 根下 pptx/pdf/txt + decks 下 html。"""
    items = []
    if os.path.isdir(OUTPUT_DIR):
        for name in os.listdir(OUTPUT_DIR):
            if not name.lower().endswith(ARTIFACT_EXT):
                continue
            full = os.path.join(OUTPUT_DIR, name)
            if not os.path.isfile(full):
                continue
            try:
                mtime = os.path.getmtime(full)
            except OSError:
                continue
            kind = name.rsplit(".", 1)[-1].lower()
            items.append({"name": name, "kind": kind, "url": f"/files/{quote(name)}",
                          "mtime": mtime})
    if os.path.isdir(DECKS_DIR):
        for name in os.listdir(DECKS_DIR):
            if not name.lower().endswith(".html"):
                continue
            full = os.path.join(DECKS_DIR, name)
            try:
                mtime = os.path.getmtime(full)
            except OSError:
                continue
            items.append({"name": name, "kind": "html", "url": f"/decks/{quote(name)}",
                          "mtime": mtime})
    items.sort(key=lambda x: x["mtime"], reverse=True)
    for it in items:
        stem = it["name"].rsplit(".", 1)[0]
        parts = stem.rsplit("_", 2)
        it["title"] = parts[0].replace("_", " ") if len(parts) == 3 else stem
        it["when"] = time.strftime("%m-%d %H:%M", time.localtime(it.pop("mtime")))
    return jsonify({"artifacts": items[:30]})


@app.route("/files/<path:filename>")
def artifact_files(filename):
    """下载 output 根下的产物文件。send_from_directory 内部做 safe_join，越界路径被拒。"""
    if not filename.lower().endswith(ARTIFACT_EXT):
        return jsonify({"error": "该类型不支持下载"}), 403
    return send_from_directory(OUTPUT_DIR, filename, as_attachment=True)


@app.route("/api/decks")
def api_decks():
    """列出历史设计稿，供侧栏历史项目区（按时间倒序，最多 30 条）。"""
    if not os.path.isdir(DECKS_DIR):
        return jsonify({"decks": []})
    items = []
    for name in os.listdir(DECKS_DIR):
        if not name.lower().endswith(".html"):
            continue
        full = os.path.join(DECKS_DIR, name)
        try:
            mtime = os.path.getmtime(full)
        except OSError:
            continue  # 列目录与取 mtime 之间文件被删（审计 P2-1）
        stem = name[:-5]
        # 文件名格式：<主题>_<YYYYmmdd>_<HHMMSS>.html，拆出主题与时间
        parts = stem.rsplit("_", 2)
        title = parts[0].replace("_", " ") if len(parts) == 3 else stem
        items.append({
            "title": title,
            "url": f"/decks/{quote(name)}",
            "mtime": mtime,
        })
    items.sort(key=lambda x: x["mtime"], reverse=True)
    for it in items:
        it["when"] = time.strftime("%m-%d %H:%M", time.localtime(it.pop("mtime")))
    return jsonify({"decks": items[:30]})


if __name__ == "__main__":
    os.makedirs(IMAGES_DIR, exist_ok=True)
    app.run(host="127.0.0.1", port=5000, debug=False)
