"""入口：python main.py "<主题>" → 大纲 → 逐页生图 → 组装 pptx。"""

import os
import re
import sys
import time

from dotenv import load_dotenv

import builder
import image_gen
import outline

load_dotenv()

OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")


def main():
    if len(sys.argv) < 2:
        print('用法: python main.py "<主题文字>"，例: python main.py "青少年为什么要学编程"')
        sys.exit(1)
    topic = sys.argv[1]

    os.makedirs(os.path.join(OUTPUT_DIR, "images"), exist_ok=True)

    print(f"[1/3] 生成大纲: {topic}")
    slides = outline.generate_outline(topic)
    print(f"      共 {len(slides)} 页: " + " / ".join(s.get("title", "") for s in slides))

    print("[2/3] 逐页生成图片（cogview-3-flash，串行）")
    image_paths = []
    images_dir = os.path.join(OUTPUT_DIR, "images")
    for i, s in enumerate(slides):
        prompt = s.get("image_prompt", "")
        path = os.path.join(images_dir, f"slide_{i}.png")
        if not prompt:
            image_paths.append(None)
            print(f"      页 {i + 1}: 无配图提示词，跳过")
            continue
        try:
            image_gen.generate_image(prompt, path)
            image_paths.append(path)
            print(f"      页 {i + 1}: {os.path.basename(path)} 完成")
        except Exception as e:
            image_paths.append(None)
            print(f"      页 {i + 1}: 生成失败（{e}），留占位框")
        time.sleep(1)

    print("[3/3] 组装 PPT")
    safe_topic = re.sub(r'[\\/:*?"<>| ]', "_", topic)[:20]
    out_path = os.path.join(OUTPUT_DIR, f"{safe_topic}_{time.strftime('%Y%m%d_%H%M%S')}.pptx")
    builder.build_ppt(slides, image_paths, out_path)

    print(f"完成: {out_path}")


if __name__ == "__main__":
    main()
