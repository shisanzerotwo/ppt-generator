"""LLM 自主设计 HTML 幻灯片：把大纲与图片交给 LLM 直接生成单文件 HTML。

设计理念：不再用代码模板渲染，而是让 LLM 根据每页内容自主设计布局/配色/装饰，
真正"按内容适配"。默认模型走 .env 的 ZHIPUAI_CHAT_MODEL（agnes-2.0-flash），
设计任务可用 ZHIPUAI_DESIGN_MODEL 单独指定（如余额恢复后的 agnes-2.5-pro）。
"""

import json
import os

from dotenv import load_dotenv

import outline

load_dotenv()

PROMPT_TEMPLATE = """你是一位顶级幻灯片视觉设计师。根据下面提供的 PPT 大纲，直接输出一份完整的单文件 HTML 幻灯片文档。

【硬性要求】
1. 输出完整 HTML：以 <!DOCTYPE html> 开头、以 </html> 结尾，全部样式内联在文档内
2. 每页一个 <section class="slide">，占满一屏（100vh），CSS scroll-snap 滚动翻页，并附一小段内联 <script> 支持左右方向键/PageUp/PageDown 翻页
3. 只用 <style> 内联 CSS 与内联 SVG，禁止引用任何外部资源（无 CDN、无外链字体、无外链图片）
4. 图片一律使用页面数据中给定的 image 相对路径，不得虚构或改动其他路径；无 image 字段的页面不要放图

【按内容自主设计（核心）】
每页布局必须因内容而异，禁止所有页面同构：
- cover：视觉冲击排版（超大标题 + 渐变/几何装饰/光效）
- toc：目录索引设计（编号/分栏/进度感）
- section：章节过渡页（大字 + 大量留白 + 装饰图形）
- data：用纯 CSS/SVG 把 chart 数据画出来（占比用饼图/环形图、趋势用折线、类别对比用条形，或大数字卡片），必须带数值标注
- timeline：横向或纵向时间轴设计
- compare：左右对照栏（两侧可用不同底色强调差异）
- content：卡片/网格/编号列表/图文分栏等，按要点数量与性质自行选择
- end：收尾设计（总结大字 + 行动号召）

【视觉规范】
- 现代配色：统一色系（深色或浅色）+ 渐变主色 + 强调色点缀，全篇和谐
- 系统字体栈（Microsoft YaHei / PingFang SC / system-ui）
- 字号层级清晰，中文排版舒适（行高 1.6 以上，页面留白充分）
- 有图的页：图片要有设计感地融入（圆角/阴影/遮罩/裁切构图），不要简单堆放
- 可用内联 SVG 几何图形、渐变、光效、大数字、简易图标提升质感
- 每页底部可加页码与主题小字标注

【大纲数据】
主题：{topic}
页面（JSON 数组，image 为该页图片相对路径，缺失表示无图页）：
{slides_json}

只输出 HTML 文档本身，不要任何解释文字、不要代码块围栏。"""

MAX_TOKENS = 16000

# 打印分页兜底：prompt 里要求过，但实测 LLM 常漏，靠代码补更可靠
PRINT_CSS = """
<style>
@media print {
  @page { size: 1280px 720px; margin: 0; }
  html, body { width: 1280px; }
  .slide { page-break-after: always; break-after: page; min-height: 720px; }
  .slide:last-child { page-break-after: auto; break-after: auto; }
}
</style>"""


def _ensure_print_css(html: str) -> str:
    """LLM 漏写打印分页规则时补上，保证浏览器能正确打印成 PDF（每页一张）。

    检测与定位均大小写不敏感（审计 P2-3）；@media print 存在但缺 @page 时仍注入
    （审计 P2-4：有 print 块没 @page 同样分不了页）。
    """
    lower = html.lower()
    if "@media print" in lower and "@page" in lower:
        return html
    idx = lower.rfind("</head>")
    if idx == -1:
        idx = lower.rfind("</body>")
    if idx == -1:
        return html
    return html[:idx] + PRINT_CSS + "\n" + html[idx:]


def _call_llm(prompt: str) -> str:
    # 设计任务输出量大（实测 129~200s），不复用 outline 的 60s 短超时客户端（审计 M1）
    api_key = os.getenv("ZHIPUAI_API_KEY")
    if not api_key or api_key == "your-key-here":
        raise RuntimeError("未配置 ZHIPUAI_API_KEY")
    client = outline.ZhipuAI(
        api_key=api_key,
        base_url=os.getenv("ZHIPUAI_BASE_URL") or None,
        timeout=300.0,
    )
    model = os.getenv("ZHIPUAI_DESIGN_MODEL") or os.getenv("ZHIPUAI_CHAT_MODEL") or "agnes-2.0-flash"
    resp = client.chat.completions.create(
        model=model,
        messages=[{"role": "user", "content": prompt}],
        max_tokens=MAX_TOKENS,
        temperature=0.7,
    )
    return resp.choices[0].message.content or ""


def _extract_html(text: str) -> str:
    """从回复中提取 <!DOCTYPE html> 到 </html> 的完整文档，截断时报错。"""
    lower = text.lower()
    start = lower.find("<!doctype html>")
    if start == -1:
        raise ValueError("回复中未找到 HTML 文档（缺 DOCTYPE）")
    end = text.find("</html>", start)
    if end == -1:
        raise ValueError("HTML 不完整（缺少 </html>，疑似截断）")
    return text[start : end + len("</html>")]


def _build_prompt(topic: str, slides: list[dict], image_map: dict[int, str], style: dict | None = None) -> str:
    deck_slides = []
    for i, s in enumerate(slides):
        item = {
            "type": s.get("type", "content"),
            "title": s.get("title", ""),
            "points": s.get("points", []),
        }
        if s.get("chart"):
            item["chart"] = s["chart"]
        if image_map.get(i):
            item["image"] = image_map[i]
        deck_slides.append(item)
    slides_json = json.dumps(deck_slides, ensure_ascii=False, indent=1)
    prompt = PROMPT_TEMPLATE.format(topic=topic, slides_json=slides_json)
    if style:
        import style as style_mod
        guidance = style_mod.style_guidance(style)
        if guidance:
            prompt = prompt.replace("【大纲数据】", guidance + "\n\n【大纲数据】", 1)
    return prompt


def generate_html_deck(topic: str, slides: list[dict], image_map: dict[int, str], style: dict | None = None) -> str:
    """生成完整 HTML 幻灯片文档；输出截断时自动重试一次。"""
    prompt = _build_prompt(topic, slides, image_map, style)
    last_err = None
    for _ in range(2):
        text = _call_llm(prompt)
        try:
            return _ensure_print_css(_extract_html(text))
        except ValueError as e:
            last_err = e
    raise RuntimeError(f"HTML 生成失败（已重试 1 次）: {last_err}")
