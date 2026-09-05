"""_design_and_save 的路径编码回归：含危险字符的主题不得原样进入 html_path（审计 C1）。"""

import urllib.parse

import app as app_module


def test_design_encodes_unsafe_topic(monkeypatch, tmp_path):
    monkeypatch.setattr(app_module, "DECKS_DIR", str(tmp_path))
    monkeypatch.setattr(app_module, "PROJECTS_DIR", str(tmp_path))
    monkeypatch.setattr(app_module.html_gen, "generate_html_deck", lambda t, s, m, st=None: "<html></html>")
    app_module.state["topic"] = "x');alert(1);(`y"
    app_module.state["slides"] = [{"type": "end", "title": "t", "points": [], "image": None}]
    try:
        assert app_module._design_and_save() is True
        hp = app_module.state["html_path"]
        assert hp.startswith("/decks/")
        # 单引号/反引号/括号必须被百分号编码，前端 onclick 拼接无法闭合
        assert "'" not in hp and "`" not in hp and "(" not in hp
        # 解码后应能还原为实际文件名
        name = urllib.parse.unquote(hp.removeprefix("/decks/"))
        assert name.endswith(".html")
    finally:
        # _design_and_save 会把 phase 置为 designing，必须一并恢复，
        # 否则后续走 /api/generate 的用例全部 409（测试间污染）
        app_module.state["topic"] = ""
        app_module.state["slides"] = []
        app_module.state["html_path"] = None
        app_module.state["phase"] = "idle"
