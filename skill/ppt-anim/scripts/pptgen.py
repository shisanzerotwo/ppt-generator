#!/usr/bin/env python
"""ppt-anim skill 的 CLI 薄封装：定位项目 venv 的 python，转发到 <project>/cli.py。

**不复制任何逻辑** —— 本体在 `cli.py`（契约 §8.4）。本脚本只做三件事：
① 找到项目根；② 找到 venv 解释器；③ 原样转发 argv，透传退出码与两条流。

用**任何** python 都能跑（只用标准库），所以宿主的 python 缺依赖也没关系。

定位顺序（可用环境变量覆盖）：
  项目根：$PPTGEN_HOME → 本脚本向上找含 cli.py 的目录 → 常见默认路径
  解释器：$PPTGEN_PYTHON → <项目根>/.venv/Scripts/python.exe（Windows）
          → <项目根>/.venv/bin/python（POSIX）→ 当前解释器
"""

import os
import subprocess
import sys

DEFAULT_HOME = r"D:\GitHub\xiangmu\ppt-generator"


def find_project_root() -> str | None:
    env = os.environ.get("PPTGEN_HOME")
    if env and os.path.isfile(os.path.join(env, "cli.py")):
        return os.path.abspath(env)
    here = os.path.dirname(os.path.abspath(__file__))
    for _ in range(6):                       # 从 skill/ppt-anim/scripts/ 往上找
        if os.path.isfile(os.path.join(here, "cli.py")):
            return here
        parent = os.path.dirname(here)
        if parent == here:
            break
        here = parent
    if os.path.isfile(os.path.join(DEFAULT_HOME, "cli.py")):
        return DEFAULT_HOME
    return None


def find_python(root: str) -> str:
    env = os.environ.get("PPTGEN_PYTHON")
    if env and os.path.isfile(env):
        return env
    for rel in (os.path.join(".venv", "Scripts", "python.exe"),
                os.path.join(".venv", "bin", "python")):
        cand = os.path.join(root, rel)
        if os.path.isfile(cand):
            return cand
    return sys.executable


def main() -> int:
    root = find_project_root()
    if root is None:
        sys.stderr.write(
            "找不到 ppt-generator 项目根（里面应有 cli.py）。\n"
            "请设置环境变量 PPTGEN_HOME 指向项目目录，例如：\n"
            f'  PPTGEN_HOME="{DEFAULT_HOME}"\n')
        return 5
    python = find_python(root)
    cli = os.path.join(root, "cli.py")
    proc = subprocess.run([python, cli] + sys.argv[1:], cwd=root)
    return proc.returncode


if __name__ == "__main__":
    raise SystemExit(main())
