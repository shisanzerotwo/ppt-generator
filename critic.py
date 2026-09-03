"""方案 C2：视觉回看校验 + 对话式修改（提议者-审核者）。"""

import base64
import os

from dotenv import load_dotenv
from zhipuai import ZhipuAI

from outline import _extract_json, _normalize

load_dotenv()

VISION_MODEL = os.getenv("ZHIPUAI_VISION_MODEL") or "agnes-2.5-pro"
TEXT_MODEL = os.getenv("ZHIPUAI_CHAT_MODEL") or "agnes-2.0-flash"


def _client() -> ZhipuAI:
    api_key = os.getenv("ZHIPUAI_API_KEY")
    if not api_key or api_key == "your-key-here":
        raise RuntimeError("未配置 ZHIPUAI_API_KEY")
    return ZhipuAI(api_key=api_key, base_url=os.getenv("ZHIPUAI_BASE_URL") or None, timeout=60.0)


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
