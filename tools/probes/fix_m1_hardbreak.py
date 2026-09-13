"""M1 对照探针：`a:br`（段内软换行）在两套度量里算出的行数不一致。

背景（AUDIT_REPORT §2 M1）
--------------------------
契约 §2.3 让 `a:br` 读成 `"\\n"` 存进 `ParaInfo.text`；§4.3 又让"整段交给
`wrap_lines`"。但 `qa._tokenize` 把 `"\\n"` 归进 **cjk 分支**（既不是空格也不是
alnum），于是它被当成一个 1.0em 宽的**字形**，而不是硬换行 —— PowerPoint 却会在这里
断行。结果：

  · 真正产出几何的 `hl_layout._paragraph_lines` 先按 `\\n` 硬拆（渲染对，但与 qa 差 +1 行）
  · 契约 §4.2 那条"防漂移"等价断言只盖 `wrap_lines` —— **盖不住真正被调用的那个函数**

本探针纯内存，不启动任何外部程序。

判定：
  修复前 → _paragraph_lines=2 / wrap_lines=1 / qa=1（三者不一致）
  修复后 → 三者一致（硬换行在 qa 与 wrap_lines 里都是"断行"）

跑法：.venv/Scripts/python.exe tools/probes/fix_m1_hardbreak.py
"""

import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", ".."))
sys.path.insert(0, ROOT)

CASES = [
    ("两行", "第一行\n第二行", 18.0, 300.0),
    ("段尾换了行", "只有一行\n", 18.0, 300.0),
    ("连续两个硬换行", "甲\n\n乙", 18.0, 300.0),
    ("硬换行 + 软折行", "前缀字" * 20 + "\n后缀", 18.0, 300.0),
    ("行首带空格", "甲\n 乙", 18.0, 300.0),
]


def main() -> int:
    import hl_layout
    import qa

    print(f"{'用例':<16}{'_paragraph_lines':>18}{'wrap_lines':>12}{'qa':>6}   一致?")
    bad = 0
    for label, text, size, box in CASES:
        para = _para(text, size)
        p_lines = hl_layout._paragraph_lines(text, size, box, True)
        w_lines = hl_layout.wrap_lines(text, size, box)
        q_lines = qa.measure_text_lines(text, size, box)
        same = len(p_lines) == len(w_lines) == q_lines
        if not same:
            bad += 1
        print(f"{label:<16}{len(p_lines):>18}{len(w_lines):>12}{q_lines:>6}   "
              f"{'OK' if same else '**不一致**'}")

    print("\n" + "-" * 70)
    if bad == 0:
        print("判定：**FIXED** —— 三套度量对硬换行给出一致的行数")
    else:
        print(f"判定：**REPRODUCED** —— {bad}/{len(CASES)} 个用例上三套度量不一致（M1 成立）")
    print("-" * 70)

    print("\n附：契约 §4.2 要求的等价断言现在盖到哪一层")
    print("  wrap_lines ↔ qa            ：上面已对过")
    print("  _paragraph_lines ↔ qa      ：上面已对过（这是**真正产出几何**的那层）")
    return 0


def _para(text, size):
    from pptx_io import ParaInfo, RunInfo
    return ParaInfo(text=text, runs=[RunInfo(text=text, size_pt=size, bold=None,
                                             italic=None, font_name=None)])


if __name__ == "__main__":
    raise SystemExit(main())
