"""模型运行时切换 + 多渠道：解析优先级 / 渠道 CRUD / key 打码 / 切渠道 client 生效。"""

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


def test_images_enabled_defaults_true(runtime_cfg):
    """没配 images 键时默认开（保持既有行为）。"""
    assert llm_util.images_enabled() is True


def test_images_enabled_reads_runtime_flag(runtime_cfg):
    llm_util._save_runtime({"images": False})
    assert llm_util.images_enabled() is False


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


def test_decide_style_uses_runtime_chat_model(runtime_cfg, monkeypatch):
    """全部 LLM 调用方走 get_model：界面切对话档后风格判定立即跟随（style.py 曾漏接）。"""
    import style

    monkeypatch.delenv("ZHIPUAI_CHAT_MODEL", raising=False)
    llm_util.set_runtime_models({"chat": "ui-chat-model"})

    captured = {}

    class _Msg:
        content = '{"key": "fresh-light", "reason": "r"}'

    class _Resp:
        choices = [type("C", (), {"message": _Msg()})()]

    def fake_create(**kw):
        captured["model"] = kw.get("model")
        return _Resp()

    class _FakeClient:
        class chat:
            class completions:
                create = staticmethod(fake_create)

    monkeypatch.setattr(style, "llm_client", lambda timeout=60.0: _FakeClient())
    style.decide_style("测试主题", [{"type": "cover", "title": "t"}])
    assert captured["model"] == "ui-chat-model"


# ---------------------------------------------------------------- 多渠道

def test_default_channel_builtin_and_client_uses_env(runtime_cfg, monkeypatch):
    """无自定义渠道时：内置默认渠道走 .env，行为与旧版完全一致。"""
    monkeypatch.setenv("ZHIPUAI_API_KEY", "env-key")
    monkeypatch.setenv("ZHIPUAI_BASE_URL", "https://env.example.com/v1")

    made = {}

    class _FakeZhipuAI:
        def __init__(self, api_key=None, base_url=None, timeout=None):
            made["api_key"], made["base_url"] = api_key, base_url

    monkeypatch.setattr(llm_util, "ZhipuAI", _FakeZhipuAI)
    llm_client_obj = llm_util.llm_client(60.0)
    assert made == {"api_key": "env-key", "base_url": "https://env.example.com/v1"}
    assert llm_client_obj is not None


def test_channel_select_switches_client(runtime_cfg, client, monkeypatch):
    """选中自定义渠道后，llm_client 用该渠道的 base_url/key（对话/生图/设计全链路共用）。"""
    monkeypatch.setenv("ZHIPUAI_API_KEY", "env-key")
    monkeypatch.setenv("ZHIPUAI_BASE_URL", "https://env.example.com/v1")
    resp = client.post("/api/channels", json={
        "name": "智谱官方", "base_url": "https://open.bigmodel.cn/api/paas/v4",
        "api_key": "sk-ch-1234567890", "models": ["glm-4-flash"]})
    assert resp.status_code == 200
    cid = resp.get_json()["channel"]["id"]

    client.post("/api/channels/select", json={"id": cid})
    made = {}

    class _FakeZhipuAI:
        def __init__(self, api_key=None, base_url=None, timeout=None):
            made["api_key"], made["base_url"] = api_key, base_url

    monkeypatch.setattr(llm_util, "ZhipuAI", _FakeZhipuAI)
    llm_util.llm_client(60.0)
    assert made == {"api_key": "sk-ch-1234567890",
                    "base_url": "https://open.bigmodel.cn/api/paas/v4"}


def test_channel_list_masks_api_key(runtime_cfg, client):
    """GET /api/models 的渠道列表不回显完整 key（打码成尾 4 位）。"""
    client.post("/api/channels", json={
        "name": "c1", "base_url": "https://a.com/v1", "api_key": "sk-secret-987654321"})
    channels = client.get("/api/models").get_json()["channels"]
    custom = [c for c in channels if not c.get("builtin")]
    assert len(custom) == 1
    assert "sk-secret-987654321" not in json.dumps(channels)
    assert custom[0]["api_key"].endswith("4321")  # 打码只留尾 4 位


def test_channel_persist_keeps_full_key(runtime_cfg, client):
    """打码只影响 API 回显，落盘仍是完整 key（否则切换后 client 拿不到真 key）。"""
    client.post("/api/channels", json={
        "name": "c1", "base_url": "https://a.com/v1", "api_key": "sk-full-key-abcdef"})
    cfg = json.load(open(runtime_cfg, encoding="utf-8"))
    assert cfg["channels"][0]["api_key"] == "sk-full-key-abcdef"


def test_channel_validation(runtime_cfg, client):
    # base_url 必须 http(s)
    r = client.post("/api/channels", json={"name": "x", "base_url": "ftp://a", "api_key": "sk-1"})
    assert r.status_code == 400
    # 无 key 拒收
    r = client.post("/api/channels", json={"name": "x", "base_url": "https://a.com/v1"})
    assert r.status_code == 400
    # 无名称拒收
    r = client.post("/api/channels", json={"name": " ", "base_url": "https://a.com/v1", "api_key": "sk-1"})
    assert r.status_code == 400
    assert not os.path.exists(runtime_cfg)


def test_channel_delete_and_builtin_protected(runtime_cfg, client, monkeypatch):
    monkeypatch.setenv("ZHIPUAI_API_KEY", "env-key")
    client.post("/api/channels", json={
        "name": "c1", "base_url": "https://a.com/v1", "api_key": "sk-1", "models": ["m1"]})
    channels = client.get("/api/models").get_json()["channels"]
    builtin = next(c for c in channels if c.get("builtin"))
    custom = next(c for c in channels if not c.get("builtin"))
    # 内置默认渠道不可删
    assert client.post("/api/channels/delete", json={"id": builtin["id"]}).status_code == 403
    # 自定义可删
    assert client.post("/api/channels/delete", json={"id": custom["id"]}).get_json()["ok"]
    assert len(client.get("/api/models").get_json()["channels"]) == 1


def test_channel_select_back_to_default(runtime_cfg, client, monkeypatch):
    """切走再切回默认渠道，client 恢复 .env。"""
    monkeypatch.setenv("ZHIPUAI_API_KEY", "env-key")
    monkeypatch.setenv("ZHIPUAI_BASE_URL", "https://env.example.com/v1")
    r = client.post("/api/channels", json={
        "name": "c1", "base_url": "https://a.com/v1", "api_key": "sk-1"})
    cid = r.get_json()["channel"]["id"]
    client.post("/api/channels/select", json={"id": cid})
    client.post("/api/channels/select", json={"id": "default"})
    made = {}

    class _FakeZhipuAI:
        def __init__(self, api_key=None, base_url=None, timeout=None):
            made["api_key"], made["base_url"] = api_key, base_url

    monkeypatch.setattr(llm_util, "ZhipuAI", _FakeZhipuAI)
    llm_util.llm_client(60.0)
    assert made == {"api_key": "env-key", "base_url": "https://env.example.com/v1"}


def test_channel_model_list_add_remove(runtime_cfg, client):
    """往渠道添加/移除模型名；模型名过白名单校验。"""
    r = client.post("/api/channels", json={
        "name": "c1", "base_url": "https://a.com/v1", "api_key": "sk-1"})
    cid = r.get_json()["channel"]["id"]
    # 添加
    assert client.post("/api/channels/models",
                       json={"channel": cid, "model": "glm-4-plus"}).status_code == 200
    ch = next(c for c in client.get("/api/models").get_json()["channels"] if c["id"] == cid)
    assert "glm-4-plus" in ch["models"]
    # 非法模型名拒收
    r = client.post("/api/channels/models", json={"channel": cid, "model": "bad name;"})
    assert r.status_code == 400
    # 移除
    assert client.post("/api/channels/models",
                       json={"channel": cid, "model": "glm-4-plus", "remove": True}).status_code == 200
    ch = next(c for c in client.get("/api/models").get_json()["channels"] if c["id"] == cid)
    assert "glm-4-plus" not in ch["models"]


def test_legacy_outline_and_image_gen_use_llm_client(runtime_cfg, monkeypatch):
    """outline/image_gen 不再自建 client（否则切渠道这两处不跟随，用错 key）。"""
    import inspect

    import image_gen
    import outline

    assert "llm_client" in inspect.getsource(outline) or "llm_util" in inspect.getsource(outline)
    assert "llm_client" in inspect.getsource(image_gen)
    # 自建 ZhipuAI 构造必须从这两个模块消失（收敛到 llm_util 单点）
    assert "ZhipuAI(" not in inspect.getsource(outline)
    assert "ZhipuAI(" not in inspect.getsource(image_gen)
