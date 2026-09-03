"""critic 模块容错降级逻辑单元测试（离线，不调真实 vision API）。"""

from critic import review_image


def test_review_image_missing_file_returns_ok(tmp_path):
    """图片不存在时应降级为 ok=True，不抛异常。"""
    missing = str(tmp_path / "not_exist.png")
    result = review_image("标题", ["要点1"], missing)
    assert result == {"ok": True, "reason": "", "advice": ""}
