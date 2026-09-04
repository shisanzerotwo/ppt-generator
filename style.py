"""AI 智能选风格：根据主题/内容自主判断视觉基调，替代固定主题下拉框。

风格预设库提炼自 ppt-maker skill 的视觉规范（主题色 1 + 强调色 1 + 中性灰，
正文不超 3 色，系统无衬线字体，整套统一不乱换）。
"""

import json
import os
import re

from dotenv import load_dotenv

from zhipuai import ZhipuAI

load_dotenv()

# 与 app._HTML_THEMES / builder.THEMES 的映射（pptx 降级导出用近似色板）
THEME_MAP = {
    "tech-deep": "dark",
    "business-minimal": "dark",
    "nature-green": "green",
    "fresh-light": "blue",
    "warm-editorial": "blue",
    "vivid-gradient": "blue",
}

STYLE_LIBRARY = {
    "tech-deep": {
        "name": "深空科技",
        "bg": "#0b1220", "accent": "#38bdf8", "fg": "#f1f5f9", "muted": "#94a3b8",
        "mood": "科技感、前沿、理性",
        "guidance": "深色底 + 亮青强调，几何线条/网格/光效装饰，适合科技、编程、AI、数据主题",
    },
    "business-minimal": {
        "name": "极简商务",
        "bg": "#111827", "accent": "#f59e0b", "fg": "#f9fafb", "muted": "#9ca3af",
        "mood": "专业、克制、可信",
        "guidance": "大量留白 + 大数字排版，细线分隔，适合商务、管理、行业分析主题",
    },
    "fresh-light": {
        "name": "清新浅色",
        "bg": "#f8fafc", "accent": "#4f46e5", "fg": "#1e293b", "muted": "#64748b",
        "mood": "轻盈、明快、易读",
        "guidance": "浅色底 + 靛蓝强调，卡片圆角柔和阴影，适合教育、科普、成长主题",
    },
    "warm-editorial": {
        "name": "暖调人文",
        "bg": "#fffbeb", "accent": "#d97706", "fg": "#292524", "muted": "#78716c",
        "mood": "温暖、人文、叙事感",
        "guidance": "米白底 + 琥珀强调，衬线感标题排版，适合文化、历史、故事、公益主题",
    },
    "vivid-gradient": {
        "name": "活力渐变",
        "bg": "#1e1b4b", "accent": "#a855f7", "fg": "#faf5ff", "muted": "#c4b5fd",
        "mood": "年轻、活力、有冲击力",
        "guidance": "深紫底 + 多彩渐变（紫→粉→橙），大标题冲击排版，适合营销、创意、青春主题",
    },
    "nature-green": {
        "name": "自然墨绿",
        "bg": "#062e21", "accent": "#22c55e", "fg": "#ecfdf5", "muted": "#a7c4b6",
        "mood": "自然、健康、生命力",
        "guidance": "墨绿底 + 草木绿强调，有机曲线装饰，适合环保、健康、体育、农业主题",
    },
}

SELECT_PROMPT = """你是幻灯片视觉总监。根据 PPT 主题与页面内容，从风格库中选出最合适的一套视觉基调。

【风格库】
{library}

【主题】{topic}
【页面概要】{brief}

判断依据：主题领域、受众气质、内容情绪（理性/活泼/温暖/专业）。
只输出 JSON：{{"key": "风格key", "reason": "一句话理由"}}，不要输出其他文字。"""


def _client():
    api_key = os.getenv("ZHIPUAI_API_KEY")
    if not api_key or api_key == "your-key-here":
        raise RuntimeError("未配置 ZHIPUAI_API_KEY")
    return ZhipuAI(api_key=api_key, base_url=os.getenv("ZHIPUAI_BASE_URL") or None, timeout=60.0)


def _call_llm(prompt: str) -> str:
    resp = _client().chat.completions.create(
        model=os.getenv("ZHIPUAI_CHAT_MODEL") or "agnes-2.0-flash",
        messages=[{"role": "user", "content": prompt}],
        max_tokens=200,
        temperature=0.2,
    )
    return resp.choices[0].message.content or ""


def build_prompt(topic: str, slides: list[dict]) -> str:
    library = "\n".join(
        f"- {k}（{v['name']}）：{v['mood']}；{v['guidance']}" for k, v in STYLE_LIBRARY.items()
    )
    brief = "；".join(f"{s.get('type')}《{s.get('title', '')}》" for s in slides[:6])
    return SELECT_PROMPT.format(library=library, topic=topic, brief=brief)


def decide_style(topic: str, slides: list[dict]) -> dict:
    """AI 选风格；失败或输出非法时回退 fresh-light，绝不阻断主流程。"""
    fallback = dict(STYLE_LIBRARY["fresh-light"], key="fresh-light", reason="")
    try:
        text = _call_llm(build_prompt(topic, slides))
        m = re.search(r"\{[^{}]*\}", text, re.S)
        if not m:
            return fallback
        data = json.loads(m.group(0))
        key = data.get("key")
        if key not in STYLE_LIBRARY:
            return fallback
        return dict(STYLE_LIBRARY[key], key=key, reason=str(data.get("reason", "")))
    except Exception:
        return fallback


def style_guidance(style: dict) -> str:
    """把选定风格转成设计 prompt 的基调段落。"""
    if not style:
        return ""
    return (
        f"【本次设计基调（AI 已选定：{style.get('name')}）】\n"
        f"- 背景 {style.get('bg')}，强调色 {style.get('accent')}，正文 {style.get('fg')}，"
        f"次要文字 {style.get('muted')}，全篇严格使用这套色板（正文不超 3 色）\n"
        f"- 气质关键词：{style.get('mood')}\n"
        f"- 设计指引：{style.get('guidance')}\n"
        f"- 字体用系统无衬线栈（Microsoft YaHei / PingFang SC / system-ui），标题一位粗体、正文一个常规体\n"
        f"- 整套统一不乱换色，但每页布局仍须按内容变化"
    )
