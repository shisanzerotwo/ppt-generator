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
                {"title": s.get("title", ""), "points": s.get("points", []),
                 "image_prompt": s.get("image_prompt", ""), "image": None,
                 "imageStatus": "pending"}
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


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/generate", methods=["POST"])
def api_generate():
    data = request.get_json(force=True)
    topic = (data.get("topic") or "").strip()
    if not topic:
        return jsonify({"error": "主题不能为空"}), 400
    with lock:
        if state["phase"] not in ("idle", "ready"):
            return jsonify({"error": "正在生成中，请等待完成"}), 409
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
    if phase != "ready":
        return jsonify({"error": "生成尚未完成，无法导出"}), 409
    if not slides:
        return jsonify({"error": "没有可导出的页面"}), 400

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    safe_topic = re.sub(r'[\\/:*?"<>| ]', "_", topic or "ppt")[:20]
    out_path = os.path.join(OUTPUT_DIR, f"{safe_topic}_{time.strftime('%Y%m%d_%H%M%S')}.pptx")
    builder.build_ppt(
        [{"title": s["title"], "points": s["points"], "image_prompt": s["image_prompt"]} for s in slides],
        [os.path.join(IMAGES_DIR, os.path.basename(s["image"])) if s["image"] else None for s in slides],
        out_path,
    )
    _log(f"已导出：{os.path.basename(out_path)}")
    return jsonify({"ok": True, "path": out_path})


@app.route("/images/<path:filename>")
def images(filename):
    return send_from_directory(IMAGES_DIR, filename)


if __name__ == "__main__":
    os.makedirs(IMAGES_DIR, exist_ok=True)
    app.run(host="127.0.0.1", port=5000, debug=False)
