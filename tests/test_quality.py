"""quality.py 纯函数用例：重复页检测 / 内容过瘦（无 LLM、无 IO）。"""

from quality import check_deck, find_duplicate_pages, find_thin_pages


def test_duplicate_title_detected():
    slides = [
        {"title": "项目背景与目标", "points": ["a：1", "b：2"]},
        {"title": "项目背景与目标", "points": ["c：3", "d：4"]},
        {"title": "技术架构设计", "points": ["e：5", "f：6"]},
    ]
    dups = find_duplicate_pages(slides)
    assert any(d["a"] == 0 and d["b"] == 1 and d["kind"] == "title" for d in dups)


def test_duplicate_body_detected_when_titles_differ():
    slides = [
        {"title": "页面一", "points": ["统一部署：降低运维成本", "灰度发布：平滑升级"]},
        {"title": "完全不同的标题", "points": ["统一部署：降低运维成本", "灰度发布：平滑升级"]},
    ]
    dups = find_duplicate_pages(slides)
    assert any(d["kind"] == "body" for d in dups)


def test_no_false_positive_on_normal_deck():
    slides = [
        {"title": "封面页", "points": []},
        {"title": "目录", "points": ["背景", "架构"]},
        {"title": "技术架构", "points": ["前后端分离", "异步任务队列", "缓存设计"]},
    ]
    assert find_duplicate_pages(slides) == []


def test_thin_page_detection_with_exempt_types():
    slides = [
        {"type": "cover", "title": "封面", "points": []},          # 豁免
        {"type": "content", "title": "空页", "points": []},        # 过瘦
        {"type": "content", "title": "正常页", "points": ["要点一：内容足够长", "要点二：也很长"]},
    ]
    assert find_thin_pages(slides) == [1]


def test_check_deck_report_shape():
    rep = check_deck([{"type": "content", "title": "t", "points": []}])
    assert set(rep.keys()) == {"duplicates", "thin"}
    assert rep["thin"] == [0]
