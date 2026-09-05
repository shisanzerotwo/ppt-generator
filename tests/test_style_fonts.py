"""字体多样化：六风格六种标题字体 + 令牌提示词注入 + 模板提取携带字体。"""

import html_gen
import style


def test_six_distinct_title_fonts():
    fonts = {v["font_title"].split(",")[0].strip('"') for v in style.STYLE_LIBRARY.values()}
    assert len(fonts) == len(style.STYLE_LIBRARY) == 6
    assert all(v["font_body"] for v in style.STYLE_LIBRARY.values())  # 正文字体栈齐全


def test_guidance_carries_font_pair():
    g = style.style_guidance(style.STYLE_LIBRARY["warm-editorial"])
    assert "STZhongsong" in g            # 衬线标题字体进入引导语
    assert "标题字体栈" in g and "正文字体栈" in g


def test_prompt_injects_font_tokens():
    prompt = html_gen._build_prompt("主题", [{"type": "content", "title": "t", "points": ["a：1"]}],
                                    {}, style.STYLE_LIBRARY["tech-deep"])
    assert "--font-title：\"SimHei" in prompt.replace('"', '"') or "SimHei" in prompt
    assert "--font-body" in prompt
    assert "font-family:var(--font-title)" in prompt  # 硬性要求第 6 条在 prompt 里
