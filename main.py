"""入口：python main.py "<主题>" → 大纲 → 逐页生图 → 组装 pptx → 可视化报告。"""

import os
import re
import sys
import time

if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")

from dotenv import load_dotenv

import builder
import image_gen
import outline
import report

load_dotenv()

OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")


def main():
    if len(sys.argv) < 2:
        print('用法: python main.py "<主题文字>"，例: python main.py "青少年为什么要学编程"')
        sys.exit(1)
    topic = sys.argv[1]

    os.makedirs(os.path.join(OUTPUT_DIR, "images"), exist_ok=True)

    steps = []
    slides = []
    image_results = []
    out_path = None

    print(f"[1/3] 生成大纲: {topic}")
    t0 = time.time()
    try:
        slides = outline.generate_outline(topic)
        steps.append({"name": "1. 大纲生成（glm-4-flash）", "ok": True,
                      "seconds": time.time() - t0,
                      "detail": f"共 {len(slides)} 页"})
        print(f"      共 {len(slides)} 页: " + " / ".join(s.get("title", "") for s in slides))
    except Exception as e:
        steps.append({"name": "1. 大纲生成（glm-4-flash）", "ok": False,
                      "seconds": time.time() - t0, "detail": str(e)})
        _finish(topic, steps, slides, image_results, out_path)
        sys.exit(1)

    print("[2/3] 逐页生成图片（cogview-3-flash，串行）")
    images_dir = os.path.join(OUTPUT_DIR, "images")
    step_ok = True
    step_start = time.time()
    for i, s in enumerate(slides):
        prompt = s.get("image_prompt", "")
        path = os.path.join(images_dir, f"slide_{i}.png")
        t = time.time()
        if not prompt:
            image_results.append({"path": None, "ok": True, "seconds": 0})
            print(f"      页 {i + 1}: 无配图提示词，跳过")
            continue
        try:
            image_gen.generate_image(prompt, path)
            image_results.append({"path": path, "ok": True, "seconds": time.time() - t})
            print(f"      页 {i + 1}: {os.path.basename(path)} 完成（{time.time() - t:.1f}s）")
        except Exception as e:
            image_results.append({"path": None, "ok": False, "seconds": time.time() - t})
            step_ok = False
            print(f"      页 {i + 1}: 生成失败（{e}），留占位框")
        time.sleep(1)
    ok_count = sum(1 for r in image_results if r["ok"])
    steps.append({"name": "2. 逐页生图（cogview-3-flash）", "ok": step_ok,
                  "seconds": time.time() - step_start,
                  "detail": f"{ok_count}/{len(slides)} 页成功"})

    print("[3/3] 组装 PPT")
    t = time.time()
    try:
        safe_topic = re.sub(r'[\\/:*?"<>| ]', "_", topic)[:20]
        out_path = os.path.join(OUTPUT_DIR, f"{safe_topic}_{time.strftime('%Y%m%d_%H%M%S')}.pptx")
        builder.build_ppt(slides, [r["path"] for r in image_results], out_path)
        steps.append({"name": "3. PPT 组装（python-pptx）", "ok": True,
                      "seconds": time.time() - t, "detail": os.path.basename(out_path)})
        print(f"      完成: {out_path}")
    except Exception as e:
        steps.append({"name": "3. PPT 组装（python-pptx）", "ok": False,
                      "seconds": time.time() - t, "detail": str(e)})

    _finish(topic, steps, slides, image_results, out_path)


def _finish(topic, steps, slides, image_results, out_path):
    report_path = os.path.join(OUTPUT_DIR, f"report_{time.strftime('%Y%m%d_%H%M%S')}.html")
    report.build_report(topic, slides, image_results, steps, out_path, report_path)
    report.open_report(report_path)
    print(f"测试报告: {report_path}（已在浏览器打开）")


if __name__ == "__main__":
    main()
