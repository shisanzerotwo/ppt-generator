"""第一步：调 glm-4-flash 把主题文字转成结构化大纲 JSON。"""

import json
import os
import re

from dotenv import load_dotenv
from zhipuai import ZhipuAI

load_dotenv()

PROMPT_TEMPLATE = """你是一位专业的 PPT 策划师。用户会给你一个主题，请生成一份章节化的 PPT 页面序列。

页面用 type 标记版式，每页是一个对象。结构要求（必须严格遵守）：
1. 第 1 页 type="cover"（封面，title 为主题，points 空列表）；第 2 页 type="toc"（目录，title="目录"，points 为各章节标题列表）
2. 按章节组织：每章开头一个 type="section" 页（title 为章标题，points 为该章一句话简介），其后是该章 2 个 type="content" 页
3. 最后 1 页 type="end"（总结/行动项，title 一句总结）
4. 出现可量化的对比/分布信息时，用 type="data" 页替代一个 content 页，并给 chart: {{"type": "bar|column|pie|line", "labels": ["项1","项2",...], "values": [数字, ...]}}。chart.type 按数据性质选：占比/份额用 pie、时间趋势用 line、类别对比用 column、类别横向对比用 bar；拿不准可省略 type 让系统自动选
5. 有明确的时间/阶段演进（如发展历程、历史节点）时，可用 type="timeline" 页（points 为按时间顺序的节点，每条含时间+事件，如"2015年：成立"）
6. 需要左右对照（如优劣、方案对比、前后对比）时，可用 type="compare" 页（points 前一半放左栏、后一半放右栏）

总页数必须控制在 8~10 页之间，绝对不要超过 10 页，不要过度分章。

每页必须带 image_prompt（非空）：
- cover 页：描述体现主题的封面插图
- content 页：描述该页内容的具体场景，扁平插画风格
- section 页：可给章节主题相关画面，也可留空字符串 ""
- toc / end / data / timeline / compare 页：image_prompt 一律为 ""（空字符串）

视觉规范：每页只讲一个观点；要点 2~3 条、每条不超过 20 字。

【重要】所有文本字段（title / points / image_prompt）中**禁止使用英文双引号 "**。如需引用或强调，用中文书名号《》或直接叙述，不要用任何引号。这是为了确保 JSON 合法。

只输出 JSON 数组，不要输出任何其他文字。格式：
[{{"type": "content", "title": "...", "points": ["...", "..."], "image_prompt": "..."}}]

主题：{topic}"""

FROM_TEXT_PROMPT = """你是一位专业的 PPT 策划师。用户会给你一份文档/讲稿内容，请提炼成一份章节化的 PPT 页面序列。

页面用 type 标记版式，每页是一个对象。结构要求（必须严格遵守）：
1. 第 1 页 type="cover"（封面，title 为文档主题，points 空列表）；第 2 页 type="toc"（目录，title="目录"，points 为各章节标题列表）
2. 按章节组织：每章开头一个 type="section" 页（title 为章标题，points 为该章一句话简介），其后是该章 2 个 type="content" 页
3. 最后 1 页 type="end"（总结/行动项，title 一句总结）
4. 出现可量化的对比/分布信息时，用 type="data" 页替代一个 content 页，并给 chart: {{"type": "bar|column|pie|line", "labels": [...], "values": [...]}}
5. 有明确时间演进用 type="timeline" 页（points 为时间节点）；需左右对照用 type="compare" 页（points 前一半左栏、后一半右栏）

总页数必须控制在 8~10 页之间。

【提炼要求】不要照搬原文长段落，提炼成每页一个观点的短要点（每条不超过 20 字），只保留文档核心信息。

每页 image_prompt：cover/content 页非空（扁平插画风格画面描述），toc/section/end/data/timeline/compare 页一律为 ""（空字符串）。

【重要】所有文本字段（title / points / image_prompt）中**禁止使用英文双引号 "**，用《》或直接叙述，确保 JSON 合法。

只输出 JSON 数组，不要输出任何其他文字。

文档内容：
{text}"""

VALID_TYPES = {"cover", "toc", "section", "content", "data", "timeline", "compare", "end"}


def _extract_json(text: str) -> list:
    start = text.find("[")
    end = text.rfind("]")
    if start == -1 or end == -1 or end <= start:
        raise ValueError(f"回复中未找到 JSON 数组: {text[:200]}")
    return json.loads(text[start : end + 1])


def _normalize(slides: list) -> list[dict]:
    """统一字段，非法 type 回退 content，content/cover 无配图提示词时兜底。"""
    out = []
    for i, s in enumerate(slides):
        if not isinstance(s, dict):
            continue
        stype = s.get("type", "content")
        if stype not in VALID_TYPES:
            stype = "content"
        title = str(s.get("title", "") or "")
        points = [str(p).strip() for p in (s.get("points") or []) if str(p).strip()]
        # 首页且无 type 时强制 cover
        if i == 0 and (stype == "content" or not points):
            stype = "cover"
        prompt = str(s.get("image_prompt", "") or "")
        # cover/content 必须有配图提示词，缺失时用标题兜底
        if stype in ("cover", "content") and not prompt:
            prompt = f"{title}，扁平插画风格"
        item = {"type": stype, "title": title, "points": points, "image_prompt": prompt}
        item["chart"] = _normalize_chart(s.get("chart"))
        out.append(item)
    return out


def _normalize_chart(chart) -> dict | None:
    """校验并清洗 chart：values 数值化、labels/values 等长、type 合法。"""
    if not isinstance(chart, dict):
        return None
    labels = chart.get("labels") if isinstance(chart.get("labels"), list) else []
    values = chart.get("values") if isinstance(chart.get("values"), list) else []
    cleaned = []
    for v in values:
        try:
            cleaned.append(float(v))
        except (TypeError, ValueError):
            continue
    n = min(len(labels), len(cleaned))
    if n == 0:
        return None
    ctype = chart.get("type")
    return {
        "type": ctype if ctype in ("bar", "column", "pie", "line") else None,
        "labels": [str(l) for l in labels[:n]],
        "values": cleaned[:n],
    }


def _generate_once(client, model, prompt_text, feedback=""):
    """生成一次大纲，feedback 为空表示首轮，否则带自评意见改进。"""
    fb = f"\n\n上一版的自评意见（请据此改进）：{feedback}" if feedback else ""
    resp = client.chat.completions.create(
        model=model,
        messages=[{"role": "user", "content": prompt_text + fb}],
        temperature=0.3,
    )
    slides = _extract_json(resp.choices[0].message.content)
    if not isinstance(slides, list) or len(slides) < 2:
        raise ValueError(f"大纲页数异常: {len(slides) if isinstance(slides, list) else '非数组'}")
    return _normalize(slides)


def _self_critique(client, model, label, slides) -> str:
    """自评大纲结构是否完整，返回需改进的问题描述；无问题返回空串。"""
    import json
    cur = json.dumps([{"type": s["type"], "title": s["title"]} for s in slides], ensure_ascii=False)
    prompt = (
        f"主题「{label}」的 PPT 大纲如下（只列 type 和 title）：\n{cur}\n\n"
        "判断这份大纲结构是否完整合理。重点看：是否缺封面/目录/总结页、页数是否在 8~12 之间、"
        "章节划分是否清晰。只输出两行：\n第一行：没问题 或 有问题\n第二行：若有问题，一句话描述需改进什么"
    )
    resp = client.chat.completions.create(
        model=model,
        messages=[{"role": "user", "content": prompt}],
        max_tokens=200,
    )
    text = (resp.choices[0].message.content or "").strip()
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    if lines and "有问题" in lines[0] and len(lines) > 1:
        return lines[1]
    return ""


def _client():
    api_key = os.getenv("ZHIPUAI_API_KEY")
    if not api_key or api_key == "your-key-here":
        raise RuntimeError("未配置 ZHIPUAI_API_KEY，请复制 .env.example 为 .env 并填入 Key")
    base_url = os.getenv("ZHIPUAI_BASE_URL") or None
    model = os.getenv("ZHIPUAI_CHAT_MODEL") or "agnes-2.0-flash"
    return ZhipuAI(api_key=api_key, base_url=base_url, timeout=60.0), model


def _generate_with_prompt(prompt_text: str, label: str) -> list[dict]:
    client, model = _client()
    last_err = None
    for attempt in range(2):
        try:
            slides = _generate_once(client, model, prompt_text)
            # 自评迭代：有问题则带意见再生成一轮（封顶 1 次改进）
            feedback = _self_critique(client, model, label, slides)
            if feedback:
                try:
                    slides = _generate_once(client, model, prompt_text, feedback)
                except Exception:
                    pass  # 改进失败则用首轮结果
            return slides
        except Exception as e:
            last_err = e
    raise RuntimeError(f"大纲生成失败（已重试 1 次）: {last_err}")


def generate_outline(topic: str) -> list[dict]:
    return _generate_with_prompt(PROMPT_TEMPLATE.format(topic=topic), topic)


def generate_outline_from_text(text: str) -> list[dict]:
    return _generate_with_prompt(FROM_TEXT_PROMPT.format(text=text[:6000]), "文档内容")


if __name__ == "__main__":
    import sys

    result = generate_outline(sys.argv[1] if len(sys.argv) > 1 else "青少年为什么要学编程")
    print(json.dumps(result, ensure_ascii=False, indent=2))
