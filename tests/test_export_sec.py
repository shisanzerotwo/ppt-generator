"""简版 HTML 导出的 img src 白名单（审计 A 回归：项目 JSON 注入 src）。"""

from app import _build_html_deck


def test_malicious_image_not_rendered(tmp_path):
    out = tmp_path / "x.html"
    slides = [{"type": "content", "title": "t", "points": ["p：1"],
               "image": '"><svg/onload=alert(2)>'}]
    _build_html_deck("主题", slides, "blue", str(out))
    text = out.read_text(encoding="utf-8")
    assert "<svg" not in text and "<img" not in text  # 非 images/ 路径整张图不放


def test_normal_image_rendered_escaped(tmp_path):
    out = tmp_path / "y.html"
    slides = [{"type": "content", "title": "t", "points": ["p：1"],
               "image": "/images/slide_0.png"}]
    _build_html_deck("主题", slides, "blue", str(out))
    text = out.read_text(encoding="utf-8")
    assert '<img class="deck-img" src="images/slide_0.png"' in text
