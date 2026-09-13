"""LLM 调用公共件：client 构造、模型解析（运行时可切换）、多渠道管理、超时约定。

此前 critic.py 与 template.py 各自复制了一份 _client()，超时等参数已经漂移
（60s vs 120s）。收敛到单一来源，各调用方按需传超时——设计任务的 300s 长超时
与单轮问答的 60s 短超时都是合理差异，但实现只能有一份。

模型解析优先级（get_model，每次调用即取，界面切换立即生效）：
界面覆盖(runtime_config.json) > design 同族跟随 chat > .env > 内置默认。

渠道（多供应商）：运行时配置 channels 列表存「名称+base_url+key+模型清单」，
current_channel 指向当前渠道；无选中或选中内置 default 时走 .env（向后兼容）。
所有 LLM 调用方必须经 llm_client() 构造，切换渠道才全链路生效。
"""
import json
import os
import re
import uuid

from dotenv import load_dotenv
from zhipuai import ZhipuAI

load_dotenv()

# 运行时模型覆盖配置（界面"模型设置"写入）。模块常量便于测试 monkeypatch。
RUNTIME_CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "runtime_config.json")

_MODEL_ENV_KEYS = {"chat": "ZHIPUAI_CHAT_MODEL", "vision": "ZHIPUAI_VISION_MODEL",
                   "image": "ZHIPUAI_IMAGE_MODEL", "design": "ZHIPUAI_DESIGN_MODEL"}
_MODEL_DEFAULTS = {"chat": "agnes-2.0-flash", "vision": "agnes-2.5-pro",
                   "image": "agnes-image-2.5-flash", "design": "agnes-2.0-flash"}

DEFAULT_CHANNEL_ID = "default"
_CHANNEL_NAME_MAX = 30
_MODEL_NAME_RE = re.compile(r"^[A-Za-z0-9._\-/:]{1,80}$")
_URL_RE = re.compile(r"^https?://[^\s]+$")


def _load_runtime() -> dict:
    try:
        with open(RUNTIME_CONFIG_PATH, encoding="utf-8") as f:
            d = json.load(f)
        return d if isinstance(d, dict) else {}
    except (OSError, ValueError):
        return {}


def _save_runtime(cfg: dict):
    tmp = RUNTIME_CONFIG_PATH + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False)
    os.replace(tmp, RUNTIME_CONFIG_PATH)


def _channels(cfg: dict) -> list[dict]:
    """渠道列表；内置默认渠道永远在首位（走 .env，不落盘、不可删）。"""
    return [c for c in cfg.get("channels", []) if isinstance(c, dict)]


def get_current_channel() -> dict:
    """当前渠道 dict；无自定义渠道或未选中时返回 {}（= 内置默认，走 .env）。"""
    cfg = _load_runtime()
    cur = cfg.get("current_channel")
    for ch in _channels(cfg):
        if ch.get("id") == cur:
            return ch
    return {}


def add_channel(name: str, base_url: str, api_key: str, models=None) -> dict:
    """添加自定义渠道并落盘，返回渠道条目。校验失败抛 ValueError。"""
    name = str(name or "").strip()
    base_url = str(base_url or "").strip().rstrip("/")
    api_key = str(api_key or "").strip()
    if not name:
        raise ValueError("渠道名不能为空")
    if len(name) > _CHANNEL_NAME_MAX:
        raise ValueError(f"渠道名最长 {_CHANNEL_NAME_MAX} 字")
    if not _URL_RE.match(base_url):
        raise ValueError("接口地址需为 http(s):// 开头的完整 URL")
    if not api_key:
        raise ValueError("API Key 不能为空")
    clean_models = []
    for m in (models or []):
        m = str(m or "").strip()
        if m and _MODEL_NAME_RE.match(m) and m not in clean_models:
            clean_models.append(m)
    ch = {"id": uuid.uuid4().hex[:8], "name": name, "base_url": base_url,
          "api_key": api_key, "models": clean_models}
    cfg = _load_runtime()
    cfg.setdefault("channels", []).append(ch)
    _save_runtime(cfg)
    return ch


def delete_channel(cid: str) -> bool:
    """删除自定义渠道；若删的是当前渠道则回退默认。内置 default 不可删。"""
    if not cid or cid == DEFAULT_CHANNEL_ID:
        return False
    cfg = _load_runtime()
    before = len(cfg.get("channels", []))
    cfg["channels"] = [c for c in cfg.get("channels", [])
                       if isinstance(c, dict) and c.get("id") != cid]
    if len(cfg["channels"]) == before:
        return False
    if cfg.get("current_channel") == cid:
        cfg["current_channel"] = DEFAULT_CHANNEL_ID
    _save_runtime(cfg)
    return True


def select_channel(cid: str) -> dict:
    """设为当前渠道。只接受存在的渠道 id（含 default）。"""
    if cid != DEFAULT_CHANNEL_ID:
        if not any(c.get("id") == cid for c in _channels(_load_runtime())):
            raise ValueError("渠道不存在")
    cfg = _load_runtime()
    cfg["current_channel"] = cid
    _save_runtime(cfg)
    return {"current_channel": cid}


def channel_add_model(cid: str, model: str, remove: bool = False) -> dict:
    """往指定渠道添加/移除模型名（清单只是前端候选，不改变 get_model 逻辑）。"""
    model = str(model or "").strip()
    if not _MODEL_NAME_RE.match(model):
        raise ValueError("模型名只能含字母数字与 ._-/:（≤80 字符）")
    cfg = _load_runtime()
    for ch in cfg.get("channels", []):
        if isinstance(ch, dict) and ch.get("id") == cid:
            models = [m for m in ch.get("models", []) if m != model]
            if not remove:
                models.append(model)
            ch["models"] = models
            _save_runtime(cfg)
            return ch
    raise ValueError("渠道不存在")


def mask_key(api_key: str) -> str:
    """key 打码：只露尾 4 位（API 回显用，完整 key 永不出后端）。"""
    k = str(api_key or "")
    return ("*" * max(0, len(k) - 4) + k[-4:]) if len(k) > 4 else "****"


def channel_public(ch: dict) -> dict:
    """渠道条目的对外视图：key 打码，其余原样。"""
    out = dict(ch)
    out["api_key"] = mask_key(ch.get("api_key", ""))
    out["builtin"] = ch.get("id") == DEFAULT_CHANNEL_ID
    return out


def list_channels_public() -> list[dict]:
    """内置默认渠道（走 .env）+ 自定义渠道，key 均打码。"""
    default = {"id": DEFAULT_CHANNEL_ID, "name": "默认（.env）", "models": [],
               "api_key": "", "builtin": True}
    return [default] + [channel_public(c) for c in _channels(_load_runtime())]


def _load_runtime() -> dict:
    try:
        with open(RUNTIME_CONFIG_PATH, encoding="utf-8") as f:
            d = json.load(f)
        return d if isinstance(d, dict) else {}
    except (OSError, ValueError):
        return {}


def get_model(kind: str) -> str:
    """按档位解析当前应使用的模型名。kind ∈ chat / vision / image / design。"""
    cfg = _load_runtime()
    if cfg.get(kind):
        return str(cfg[kind])
    if kind == "design" and cfg.get("chat"):
        return str(cfg["chat"])  # 设计档未单独覆盖时跟随对话档
    env = os.getenv(_MODEL_ENV_KEYS.get(kind, ""))
    return env or _MODEL_DEFAULTS.get(kind, _MODEL_DEFAULTS["chat"])


def set_runtime_models(models: dict) -> dict:
    """写入界面覆盖（空值=该档恢复 .env 默认），原子落盘。返回覆盖后的配置。"""
    cfg = _load_runtime()
    for k, v in (models or {}).items():
        if k in _MODEL_ENV_KEYS:
            v = str(v or "").strip()
            if v:
                cfg[k] = v
            else:
                cfg.pop(k, None)
    _save_runtime(cfg)
    return cfg


def images_enabled() -> bool:
    """配图总开关（runtime_config.json 的 images 键，默认开）。

    生图上游不给力时（慢/质量差）一键关掉：全部走无图排版，其余流程不变。
    """
    return bool(_load_runtime().get("images", True))


def llm_client(timeout: float = 60.0) -> ZhipuAI:
    """构造 ZhipuAI 兼容客户端：当前自定义渠道优先，否则 .env。

    未配 key 直接抛错（快速失败优于静默降级）。
    """
    ch = get_current_channel()
    api_key = ch.get("api_key") or os.getenv("ZHIPUAI_API_KEY")
    if not api_key or api_key == "your-key-here":
        raise RuntimeError("未配置 ZHIPUAI_API_KEY")
    base_url = ch.get("base_url") or os.getenv("ZHIPUAI_BASE_URL") or None
    return ZhipuAI(api_key=api_key, base_url=base_url, timeout=timeout)
