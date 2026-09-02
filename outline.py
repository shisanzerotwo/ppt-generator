"""第一步：调 glm-4-flash 把主题文字转成结构化大纲 JSON。"""

import json
import os
import re

from dotenv import load_dotenv
from zhipuai import ZhipuAI

load_dotenv()

PROMPT_TEMPLATE = """你是一位专业的 PPT 策划师。用户会给你一个主题，请生成一份 PPT 大纲。

要求：
1. 共 6 页：第 1 页是封面页，第 2 页是目录/概览页，第 3~5 页是核心内容页，第 6 页是总结页
2. 每页输出 title（页标题）、points（要点列表，3~4 条，每条一句话）、image_prompt（与该页内容匹配的中文画面描述，用于 AI 生图，描述具体场景、风格统一为扁平插画风格）
3. 封面页 points 为空列表，image_prompt 描述一张体现主题的封面插图

只输出 JSON 数组，不要输出任何其他文字。格式：
[{{"title": "...", "points": ["...", "..."], "image_prompt": "..."}}]

主题：{topic}"""


def _extract_json(text: str) -> list:
    start = text.find("[")
    end = text.rfind("]")
    if start == -1 or end == -1 or end <= start:
        raise ValueError(f"回复中未找到 JSON 数组: {text[:200]}")
    return json.loads(text[start : end + 1])


def generate_outline(topic: str) -> list[dict]:
    api_key = os.getenv("ZHIPUAI_API_KEY")
    if not api_key or api_key == "your-key-here":
        raise RuntimeError("未配置 ZHIPUAI_API_KEY，请复制 .env.example 为 .env 并填入 Key")

    client = ZhipuAI(api_key=api_key)
    last_err = None
    for attempt in range(2):
        try:
            resp = client.chat.completions.create(
                model="glm-4-flash",
                messages=[{"role": "user", "content": PROMPT_TEMPLATE.format(topic=topic)}],
                temperature=0.7,
            )
            slides = _extract_json(resp.choices[0].message.content)
            if not isinstance(slides, list) or len(slides) < 2:
                raise ValueError(f"大纲页数异常: {len(slides) if isinstance(slides, list) else '非数组'}")
            return slides
        except Exception as e:
            last_err = e
    raise RuntimeError(f"大纲生成失败（已重试 1 次）: {last_err}")


if __name__ == "__main__":
    import sys

    result = generate_outline(sys.argv[1] if len(sys.argv) > 1 else "青少年为什么要学编程")
    print(json.dumps(result, ensure_ascii=False, indent=2))
