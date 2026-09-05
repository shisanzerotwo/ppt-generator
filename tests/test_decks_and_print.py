"""/api/decks 历史设计稿列表 + html_gen 打印分页兜底 的单元测试。"""

import os
import time

import pytest

import app as app_mod
import html_gen


@pytest.fixture
def client(tmp_path, monkeypatch):
    """把 DECKS_DIR 指到临时目录，避免读到真实 output。"""
    monkeypatch.setattr(app_mod, "DECKS_DIR", str(tmp_path))
    app_mod.app.config["TESTING"] = True
    return app_mod.app.test_client()


def _touch(dirpath, name, when=None):
    p = os.path.join(dirpath, name)
    with open(p, "w", encoding="utf-8") as f:
        f.write("<html></html>")
    if when is not None:
        os.utime(p, (when, when))
    return p


def test_decks_missing_dir_returns_empty(monkeypatch):
    monkeypatch.setattr(app_mod, "DECKS_DIR", r"D:\不存在的目录_xyz")
    app_mod.app.config["TESTING"] = True
    r = app_mod.app.test_client().get("/api/decks")
    assert r.status_code == 200
    assert r.get_json() == {"decks": []}


def test_decks_sorted_by_mtime_desc(client, tmp_path):
    now = time.time()
    _touch(tmp_path, "旧主题_20260901_100000.html", now - 3600)
    _touch(tmp_path, "新主题_20260904_200000.html", now)
    decks = client.get("/api/decks").get_json()["decks"]
    assert [d["title"] for d in decks] == ["旧主题", "新主题"][::-1]


def test_decks_filters_non_html(client, tmp_path):
    _touch(tmp_path, "正常_20260904_200000.html")
    _touch(tmp_path, "截图_full.png")
    decks = client.get("/api/decks").get_json()["decks"]
    assert len(decks) == 1
    assert decks[0]["title"] == "正常"


def test_decks_title_with_underscores(client, tmp_path):
    """主题名本身含下划线时，只剥掉末尾的日期与时间两段。"""
    _touch(tmp_path, "使用_Herdr_的工作方式_20260904_193552.html")
    decks = client.get("/api/decks").get_json()["decks"]
    assert decks[0]["title"] == "使用 Herdr 的工作方式"


def test_decks_url_is_encoded(client, tmp_path):
    _touch(tmp_path, "中文主题_20260904_200000.html")
    d = client.get("/api/decks").get_json()["decks"][0]
    assert d["url"].startswith("/decks/")
    assert "%" in d["url"]  # 中文被 quote 编码
    assert "中文" not in d["url"]


def test_decks_has_when_field(client, tmp_path):
    _touch(tmp_path, "主题_20260904_200000.html")
    d = client.get("/api/decks").get_json()["decks"][0]
    assert "when" in d and len(d["when"]) == 11  # MM-DD HH:MM
    assert "mtime" not in d  # 内部字段不外泄


def test_decks_caps_at_30(client, tmp_path):
    for i in range(35):
        _touch(tmp_path, f"主题{i}_20260904_2000{i:02d}.html", time.time() - i)
    decks = client.get("/api/decks").get_json()["decks"]
    assert len(decks) == 30


# ===== html_gen 打印分页兜底 =====

def test_print_css_injected_before_head_close():
    html = "<!DOCTYPE html><html><head><style>x</style></head><body></body></html>"
    out = html_gen._ensure_print_css(html)
    assert "@media print" in out
    assert out.index("@media print") < out.index("</head>")


def test_print_css_not_duplicated():
    html = "<!DOCTYPE html><html><head><style>@media print{@page{size:1280px 720px}.slide{}}</style></head><body></body></html>"
    assert html_gen._ensure_print_css(html).count("@media print") == 1


def test_print_css_injected_when_page_missing():
    # 有 @media print 但缺 @page 时仍注入（审计 P2-4）
    html = "<!DOCTYPE html><html><head><style>@media print{.slide{}}</style></head><body></body></html>"
    assert html_gen._ensure_print_css(html).count("@page") >= 1


def test_print_css_case_insensitive_tags():
    # 大写 </HEAD> 也能定位注入（审计 P2-3）
    html = "<!DOCTYPE html><html><HEAD></HEAD><body></body></html>"
    out = html_gen._ensure_print_css(html)
    assert "@media print" in out
    assert out.index("@media print") < out.index("</HEAD>")


def test_print_css_falls_back_to_body():
    html = "<html><body><section class='slide'>x</section></body></html>"
    out = html_gen._ensure_print_css(html)
    assert "@media print" in out
    assert out.index("@media print") < out.index("</body>")


def test_print_css_malformed_input_no_crash():
    assert isinstance(html_gen._ensure_print_css("<div>无标签</div>"), str)
