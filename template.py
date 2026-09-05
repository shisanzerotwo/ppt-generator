"""模板来源：内置模板库（风格色板 + 版式骨架）+ 参考稿识别（AI 看图/读 HTML 提炼风格）。

统一产出一份 style dict（bg/accent/fg/muted/mood/guidance/name/key），
交给 html_gen 作为设计基调，覆盖 style.decide_style 的 AI 自动选。
"""

import base64
import json
import os
import re

from dotenv import load_dotenv
from zhipuai import ZhipuAI

import style as style_mod

load_dotenv()

VISION_MODEL = os.getenv("ZHIPUAI_VISION_MODEL") or "agnes-2.5-pro"
TEXT_MODEL = os.getenv("ZHIPUAI_CHAT_MODEL") or "agnes-2.0-flash"

_HEX = re.compile(r"^#[0-9a-fA-F]{6}$")

# 内置模板 = 复用 style.STYLE_LIBRARY 的色板 + 追加一句版式骨架倾向，比纯风格更具体
_LAYOUT_HINT = {
    "tech-deep": "封面用超大标题＋网格/线路装饰；内容多用数据卡与编号列表；深色留白充足",
    "business-minimal": "封面居中大标题＋副标题；内容偏两栏与大数字；分隔用细线，装饰克制",
    "fresh-light": "封面明快插画感；内容多用圆角卡片网格；配色轻盈、阴影柔和",
    "warm-editorial": "封面像杂志大图＋标题压角；内容单栏叙事、首字下沉感；强调排版节奏",
    "vivid-gradient": "封面强渐变冲击＋错位大字；内容用彩色卡片；活力但保持可读",
    "nature-green": "封面有机曲线装饰；内容用图文分栏与流程；色调自然、圆润",
}


def builtin_templates() -> list[dict]:
    """列出内置模板（供界面选择），含缩略色板。"""
    out = []
    for key, s in style_mod.STYLE_LIBRARY.items():
        out.append({
            "key": key, "name": s["name"], "mood": s.get("mood", ""),
            "bg": s["bg"], "accent": s["accent"], "fg": s["fg"],
            "layout_hint": _LAYOUT_HINT.get(key, ""),
        })
    return out


def get_template_style(key: str) -> dict | None:
    """按 key 取内置模板的完整 style dict（含版式骨架并入 guidance）。"""
    s = style_mod.STYLE_LIBRARY.get(key)
    if not s:
        return None
    out = dict(s, key=key)
    hint = _LAYOUT_HINT.get(key)
    if hint:
        out["guidance"] = f"{s.get('guidance', '')}；版式骨架：{hint}"
    return out


def _sanitize(raw: dict) -> dict:
    """把 AI 识别结果清洗成合法 style dict，缺失/非法字段回退中性默认。"""
    d = style_mod.STYLE_LIBRARY["fresh-light"]  # 中性回退底座

    def hexval(k):
        v = str(raw.get(k, "") or "")
        return v if _HEX.match(v) else d[k]

    return {
        "key": "custom",
        "name": str(raw.get("name") or "参考稿风格")[:20],
        "bg": hexval("bg"), "accent": hexval("accent"),
        "fg": hexval("fg"), "muted": hexval("muted"),
        "mood": str(raw.get("mood") or "")[:40],
        "guidance": str(raw.get("guidance") or "参照参考稿的整体气质与配色")[:200],
    }


def _client() -> ZhipuAI:
    api_key = os.getenv("ZHIPUAI_API_KEY")
    if not api_key or api_key == "your-key-here":
        raise RuntimeError("未配置 ZHIPUAI_API_KEY")
    return ZhipuAI(api_key=api_key, base_url=os.getenv("ZHIPUAI_BASE_URL") or None, timeout=120.0)


_ASK = ("请分析这份演示/海报设计的视觉风格，只输出一个 JSON 对象："
        '{"name":"风格名(≤8字)","bg":"#背景主色","accent":"#强调色",'
        '"fg":"#正文色","muted":"#次要文字色","mood":"气质关键词",'
        '"guidance":"一句话描述其配色、版式与装饰特征，供仿制参考"}。'
        "颜色必须是 #RRGGBB 六位十六进制。不要输出 JSON 以外的任何文字。")


def _extract_json_obj(text: str) -> dict | None:
    """提取首个 JSON 对象；失败时退化为正则抠色值。"""
    t = (text or "").strip()
    start, end = t.find("{"), t.rfind("}")
    if start != -1 and end > start:
        try:
            obj = json.loads(t[start:end + 1])
            if isinstance(obj, dict):
                return obj
        except ValueError:
            pass
    hexes = re.findall(r"#[0-9a-fA-F]{6}", t)  # 兜底：抠出主色勉强拼一个
    if hexes:
        return {"accent": hexes[0], "bg": hexes[1] if len(hexes) > 1 else "",
                "guidance": "从参考稿提取的主色"}
    return None


def analyze_reference(kind: str, data) -> dict | None:
    """识别参考稿风格。kind='image' 时 data 为图片 bytes；'html' 时为文本。

    失败一律返回 None（调用方回退 AI 自动选风格），不抛异常打断主流程。
    """
    try:
        if kind == "image":
            b64 = base64.b64encode(data).decode()
            resp = _client().chat.completions.create(
                model=VISION_MODEL,
                messages=[{"role": "user", "content": [
                    {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{b64}"}},
                    {"type": "text", "text": _ASK},
                ]}],
                max_tokens=400,
            )
        else:
            resp = _client().chat.completions.create(
                model=TEXT_MODEL,
                messages=[{"role": "user", "content":
                           f"下面是一份网页/幻灯片的 HTML/CSS：\n{str(data)[:6000]}\n\n{_ASK}"}],
                max_tokens=400,
            )
        obj = _extract_json_obj(resp.choices[0].message.content or "")
        return _sanitize(obj) if obj else None
    except Exception:
        return None
