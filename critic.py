"""方案 C2：视觉回看校验 + 对话式修改（提议者-审核者）。"""

import base64
import json
import os

from zhipuai import ZhipuAI

from llm_util import TEXT_MODEL, VISION_MODEL, llm_client

from outline import _extract_json, _normalize


def _client() -> ZhipuAI:
    return llm_client(60.0)


def review_image(title: str, points: list[str], image_path: str) -> dict:
    """用 vision 模型回看图片，判断是否契合本页主题。

    返回 {"ok": bool 契合, "reason": str 一句话原因, "advice": str 重绘建议}。
    校验失败时返回 ok=True（降级为跳过，不中断主流程）。
    """
    if not os.path.exists(image_path):
        return {"ok": True, "reason": "", "advice": ""}
    try:
        with open(image_path, "rb") as f:
            b64 = base64.b64encode(f.read()).decode()
        pts = "、".join(points) if points else "（无要点）"
        prompt = (
            f"这是一页 PPT 的配图。页标题是「{title}」，页面要点是：{pts}。\n"
            "请判断这张图片是否契合本页主题。只输出两行：\n"
            "第一行：契合 或 不契合\n"
            "第二行：一句话原因（若契合也简要说明画面内容）"
        )
        resp = _client().chat.completions.create(
            model=VISION_MODEL,
            messages=[{"role": "user", "content": [
                {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{b64}"}},
                {"type": "text", "text": prompt},
            ]}],
            max_tokens=300,
        )
        text = (resp.choices[0].message.content or "").strip()
        lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
        verdict = lines[0] if lines else ""
        reason = lines[1] if len(lines) > 1 else ""
        ok = "不契合" not in verdict and "契合" in verdict
        return {"ok": ok, "reason": reason, "advice": reason if not ok else ""}
    except Exception:
        # 校验失败降级为跳过，不中断主流程
        return {"ok": True, "reason": "", "advice": ""}


def _ask(prompt: str) -> str:
    """文本模型单轮调用（定点改写用）。"""
    resp = _client().chat.completions.create(
        model=TEXT_MODEL,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.7,
    )
    return (resp.choices[0].message.content or "").strip()


def revise_target(slides: list[dict], target: dict, instruction: str):
    """只改写 target 指定的那一页 / 那一条要点，其余内容服务端强制原样保留。

    target: {"slide": int, "quote": str|None}
      - 有 quote：仅重写命中的那一条要点（不重生图）
      - 无 quote：仅重写该页（title/points/image_prompt），image_prompt 变化时标记需重生图
    返回 (new_slides, info)；info = {"level","slide","regen_image"}
    """
    i = target.get("slide")
    if not isinstance(i, int) or i < 0 or i >= len(slides):
        raise ValueError("目标页码不存在")
    new_slides = [dict(s) for s in slides]
    slide = new_slides[i]
    quote = str(target.get("quote") or "").strip()

    if quote:
        pts = list(slide.get("points") or [])
        idx = next((k for k, p in enumerate(pts) if quote in p), None)
        if idx is None:
            raise ValueError("该页未找到选中的文本")
        text = _ask(
            f"这是一页 PPT（标题：{slide.get('title', '')}）中的一条要点文本：\n{pts[idx]}\n\n"
            f"用户要求：{instruction}\n\n"
            "只输出改写后的这一条要点文本本身：不要解释、不要引号、不要列表符号、不要换行。"
            "若原文是「标题：描述」格式（中文冒号分隔），改写后保持该格式。"
        )
        # 只要第一行、剥掉可能的包裹符号，防止 LLM 附加解释污染数据
        new_point = text.splitlines()[0].strip() if text else ""
        new_point = new_point.strip('"“”\'').strip()
        if not new_point:
            raise ValueError("AI 返回空内容")
        pts[idx] = new_point
        slide["points"] = pts
        return new_slides, {"level": "point", "slide": i, "regen_image": False}

    obj = _ask(
        "你是一位 PPT 策划师。下面是一页 PPT 的内容（JSON 对象）：\n"
        + json.dumps({k: slide.get(k) for k in ("type", "title", "points", "image_prompt", "chart")},
                     ensure_ascii=False)
        + f"\n\n用户要求：{instruction}\n\n"
        "请只重写这一页，输出一个只包含该页对象的 JSON 数组（元素个数必须为 1）。"
        "保持该页 type 与 chart 结构不变；若内容变化影响配图，同步更新 image_prompt"
        "（该页本就不配图则保持空字符串）。不要输出任何解释文字。"
    )
    parsed = _extract_json(obj)
    if not isinstance(parsed, list) or not parsed or not isinstance(parsed[0], dict):
        raise ValueError("AI 未返回有效的单页内容")
    normalized = _normalize(parsed)[:1]
    if not normalized:
        raise ValueError("AI 返回的页内容不合法")
    rebuilt = normalized[0]
    # _normalize 会给 content 页的空 image_prompt 塞兜底描述，故 LLM 原意须从 parsed 取，
    # 否则「LLM 留空=不改配图」会被误判成新提示词，白白重生一张图
    llm_prompt = str(parsed[0].get("image_prompt", "") or "")
    old_prompt = str(slide.get("image_prompt") or "")
    rebuilt["type"] = slide.get("type")          # 版式不允许被改写带偏
    rebuilt["chart"] = rebuilt.get("chart") or slide.get("chart")
    rebuilt["image_prompt"] = llm_prompt or old_prompt
    new_slides[i] = rebuilt
    return new_slides, {
        "level": "slide", "slide": i,
        # 有新描述且与原文不同即重生（含原本无图、改写后需配图的情况）
        "regen_image": bool(llm_prompt) and llm_prompt != old_prompt,
    }


def refine_outline(slides: list[dict], instruction: str) -> list[dict]:
    """按用户自然语言指令修改整份大纲，返回同结构 slides。"""
    import json
    cur = json.dumps(
        [{"type": s.get("type"), "title": s.get("title"), "points": s.get("points"),
          "image_prompt": s.get("image_prompt"), "chart": s.get("chart")}
         for s in slides],
        ensure_ascii=False,
    )
    prompt = (
        "你是一位 PPT 策划师。当前 PPT 页面序列如下（JSON 数组）：\n"
        f"{cur}\n\n"
        f"用户指令：{instruction}\n\n"
        "请按指令修改这份页面序列，输出修改后的完整 JSON 数组。要求：\n"
        "1. 保持每页的 type 不变（cover/toc/section/content/data/end）\n"
        "2. 保留 chart 字段结构（若有）\n"
        "3. 被修改的页，image_prompt 也要同步更新成与新内容匹配的画面描述\n"
        "4. 只输出 JSON 数组，不要输出其他文字"
    )
    resp = _client().chat.completions.create(
        model=TEXT_MODEL,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.7,
    )
    slides_out = _extract_json(resp.choices[0].message.content)
    if not isinstance(slides_out, list) or not slides_out:
        raise ValueError("对话修改未返回有效 slides")
    return _normalize(slides_out)
