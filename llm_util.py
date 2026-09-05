"""LLM 调用公共件：client 构造、模型解析（运行时可切换）、超时约定。

此前 critic.py 与 template.py 各自复制了一份 _client()，超时等参数已经漂移
（60s vs 120s）。收敛到单一来源，各调用方按需传超时——设计任务的 300s 长超时
与单轮问答的 60s 短超时都是合理差异，但实现只能有一份。

模型解析优先级（get_model，每次调用即取，界面切换立即生效）：
界面覆盖(runtime_config.json) > design 同族跟随 chat > .env > 内置默认。
"""
import json
import os

from dotenv import load_dotenv
from zhipuai import ZhipuAI

load_dotenv()

# 运行时模型覆盖配置（界面"模型设置"写入）。模块常量便于测试 monkeypatch。
RUNTIME_CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "runtime_config.json")

_MODEL_ENV_KEYS = {"chat": "ZHIPUAI_CHAT_MODEL", "vision": "ZHIPUAI_VISION_MODEL",
                   "image": "ZHIPUAI_IMAGE_MODEL", "design": "ZHIPUAI_DESIGN_MODEL"}
_MODEL_DEFAULTS = {"chat": "agnes-2.0-flash", "vision": "agnes-2.5-pro",
                   "image": "agnes-image-2.5-flash", "design": "agnes-2.0-flash"}


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
    tmp = RUNTIME_CONFIG_PATH + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False)
    os.replace(tmp, RUNTIME_CONFIG_PATH)
    return cfg


def llm_client(timeout: float = 60.0) -> ZhipuAI:
    """构造 ZhipuAI 兼容客户端；未配 key 直接抛错（快速失败优于静默降级）。"""
    api_key = os.getenv("ZHIPUAI_API_KEY")
    if not api_key or api_key == "your-key-here":
        raise RuntimeError("未配置 ZHIPUAI_API_KEY")
    return ZhipuAI(api_key=api_key, base_url=os.getenv("ZHIPUAI_BASE_URL") or None, timeout=timeout)
