"""pytest 共享配置：把项目根加入 sys.path，使测试可 import 顶层模块。"""

import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))


@pytest.fixture(autouse=True)
def _isolate_runtime_config(tmp_path, monkeypatch):
    """测试一律不读开发者本地的 runtime_config.json。

    渠道/模型档/配图开关都存在这个文件里，一旦本地把某个开关调成非默认值
    （例如 images=false），依赖该默认值的用例会莫名其妙地挂。隔离到临时目录。
    """
    import llm_util

    monkeypatch.setattr(llm_util, "RUNTIME_CONFIG_PATH", str(tmp_path / "runtime_config.json"))
