"""模型运行时切换：解析优先级 / API 收发与持久化 / 非法名拒收。"""

import json
import os

import pytest

import app as app_mod
import llm_util


@pytest.fixture
def runtime_cfg(tmp_path, monkeypatch):
    monkeypatch.setattr(llm_util, "RUNTIME_CONFIG_PATH", str(tmp_path / "runtime_config.json"))
    return tmp_path / "runtime_config.json"


@pytest.fixture
def client():
    app_mod.app.config["TESTING"] = True
    return app_mod.app.test_client()


def test_get_model_precedence(runtime_cfg, monkeypatch):
    # 无覆盖：.env > 内置默认
    monkeypatch.setenv("ZHIPUAI_CHAT_MODEL", "env-chat")
    assert llm_util.get_model("chat") == "env-chat"
    monkeypatch.delenv("ZHIPUAI_CHAT_MODEL")
    assert llm_util.get_model("chat") == llm_util._MODEL_DEFAULTS["chat"]
    # 界面覆盖最高
    llm_util.set_runtime_models({"chat": "ui-chat"})
    assert llm_util.get_model("chat") == "ui-chat"
    # design 未单独覆盖时跟随对话档覆盖
    assert llm_util.get_model("design") == "ui-chat"
    llm_util.set_runtime_models({"design": "ui-design"})
    assert llm_util.get_model("design") == "ui-design"


def test_set_empty_value_restores_env_default(runtime_cfg, monkeypatch):
    monkeypatch.setenv("ZHIPUAI_CHAT_MODEL", "env-chat")
    llm_util.set_runtime_models({"chat": "ui-chat"})
    assert llm_util.get_model("chat") == "ui-chat"
    llm_util.set_runtime_models({"chat": ""})  # 空值 = 恢复 .env 默认
    assert llm_util.get_model("chat") == "env-chat"


def test_models_api_roundtrip(runtime_cfg, client):
    resp = client.post("/api/models", json={"chat": "test-model-1"})
    assert resp.status_code == 200
    body = resp.get_json()
    assert body["models"]["chat"] == "test-model-1"
    # 持久化：直接读配置文件核对（下次调用即生效的前提）
    assert json.load(open(runtime_cfg, encoding="utf-8"))["chat"] == "test-model-1"
    # 空值恢复默认
    client.post("/api/models", json={"chat": ""})
    assert client.get("/api/models").get_json()["overrides"] == {}


def test_models_api_rejects_bad_name(runtime_cfg, client):
    resp = client.post("/api/models", json={"chat": "bad model; rm -rf"})
    assert resp.status_code == 400
    assert not os.path.exists(runtime_cfg)  # 拒收时不落任何配置


def test_models_api_ignores_unknown_keys(runtime_cfg, client):
    resp = client.post("/api/models", json={"chat": "ok-model", "hacker": "x"})
    assert resp.status_code == 200
    cfg = json.load(open(runtime_cfg, encoding="utf-8"))
    assert "hacker" not in cfg
