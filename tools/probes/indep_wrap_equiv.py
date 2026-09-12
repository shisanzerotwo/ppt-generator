"""独立验证 ⑤ · `hl_layout.wrap_lines` 与 `qa.measure_text_lines` 的等价性（自选语料）。

实现者的等价测试用了 8 类语料 + 300 条伪随机 fuzz。这里换**我自己的**语料，
专挑它大概率没覆盖的：制表符、CRLF、全角空格、零宽字符、组合音标、URL、
连续空格、emoji、韩文/日文假名、只有空格、换行前后缀、超长 ASCII 无空格串。

除了行数等价，还加了三条**独立不变量**（不依赖 qa，能抓"掉字/串行"这类真 bug）：
  I1 除空格外的字符一个都不能丢：去掉空格后，各行拼接 == 原文
  I2 任何一行都不以空格开头（契约：行首空格丢弃）
  I3 行宽 ≤ 行宽上限；例外只能是"单字符本身就超宽"的强拆行

一条命令：./.venv/Scripts/python.exe tools/probes/indep_wrap_equiv.py
"""

import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

import hl_layout  # noqa: E402
import qa  # noqa: E402

CORPUS = {
    "连续空格": "a  b   c     中文  结尾",
    "全角标点行首": "。，、！？；：「测试」全角标点",
    "emoji": "😀😀😀abc测试🎉",
    "制表符": "a\tb\tc\t中文",
    "CRLF": "第一行\r\n第二行\r\n第三行",
    "纯数字长串": "1234567890" * 8,
    "长英文词+中文": "supercalifragilisticexpialidocious中文混排测试",
    "全角空格U+3000": "中文　　全角空格　　结束",
    "零宽字符": "a​b‌c",
    "组合音标": "café näive résumé",
    "长破折号串": "——————" * 6,
    "ASCII标点串": "----_____++++====",
    "混合版本号": "版本 1.5.2 (beta) — 测试；含 3,000 项",
    "换行前缀": "\n\nabc",
    "换行后缀": "abc\n\n",
    "只有空格": "     ",
    "空串": "",
    "URL": "https://example.com/a/very/long/path?c=d&e=f#frag",
    "中英夹空格": "Python 与 Java 的对比 2026",
    "日韩": "テスト 한국어 テスト",
    "制表+换行混": "a\tb\r\nc\td",
    "单字超宽CJK": "疆",
    "中英无空格紧贴": "中文English混合without空格",
    "全角数字标点": "１２３４５６７８９０！？",
}

SIZES = [8.0, 12.0, 18.0, 24.0, 40.0, 72.0]
WIDTHS = [30.0, 60.0, 120.0, 200.0, 400.0, 900.0]

problems = []


def check(label, cond, detail=""):
    if not cond:
        problems.append(f"{label}: {detail}")


def main() -> int:
    n_eq = n_total = 0
    print("== 1. 行数等价（hl_layout.wrap_lines vs qa.measure_text_lines）==")
    for name, text in CORPUS.items():
        bad = []
        for size in SIZES:
            for w in WIDTHS:
                got = len(hl_layout.wrap_lines(text, size, w))
                exp = qa.measure_text_lines(text, size, w)
                n_total += 1
                if got == exp:
                    n_eq += 1
                else:
                    bad.append((size, w, got, exp))
        if bad:
            problems.append(f"[{name}] 行数不等价：{bad[:3]}")
            print(f"  ✗ {name:16} 不等价 {bad[:3]}")
        else:
            print(f"  OK {name:16} {len(SIZES) * len(WIDTHS)} 组全一致")

    print(f"\n  等价对：{n_eq}/{n_total}")

    print("\n== 2. 独立不变量（不依赖 qa）==")
    for name, text in CORPUS.items():
        for size in (12.0, 24.0):
            for w in (45.0, 150.0, 600.0):
                lines = hl_layout.wrap_lines(text, size, w)
                joined = "".join(ln.text for ln in lines)
                stripped_src = text.replace(" ", "")
                stripped_out = joined.replace(" ", "")
                check(f"[{name} s={size} w={w}] I1 掉字",
                      stripped_src == stripped_out,
                      f"原文={stripped_src!r} 输出={stripped_out!r}")
                for ln in lines:
                    check(f"[{name}] I2 行首空格", not ln.text.startswith(" "),
                          repr(ln.text[:20]))
                for ln in lines:
                    over = ln.width_pt > w + 1e-9
                    if over:
                        # 允许：整行只有一个字符（强拆行）
                        check(f"[{name} s={size} w={w}] I3 超宽行非强拆",
                              len(ln.text) == 1,
                              f"宽={ln.width_pt:.1f} 上限={w} 行={ln.text[:16]!r}")
                if text.strip(" "):
                    check(f"[{name} s={size} w={w}] I4 出现空行",
                          all(ln.text for ln in lines),
                          [ln.text for ln in lines if not ln.text])
    print(f"  不变量检查完成，累计问题 {len(problems)} 项")

    print("\n== 3. `_paragraph_lines`（含手动换行 \n）与逐段 wrap_lines 的加和关系 ==")
    for name in ("CRLF", "换行前缀", "换行后缀", "制表+换行混"):
        text = CORPUS[name]
        for size in (12.0, 24.0):
            for w in (60.0, 240.0):
                got = len(hl_layout._paragraph_lines(text, size, w, True))
                exp = sum(len(hl_layout.wrap_lines(part, size, w)) for part in text.split("\n"))
                check(f"[{name}] 手动换行加和", got == exp, f"{got} != {exp}")
    print("  （不等则说明 _paragraph_lines 的 a:br 拆分与逐段 wrap 不一致）")

    print("\n== 4. 与 qa 在 word_wrap=False 下的**有意分歧**（契约规定的例外）==")
    long_text = "很长很长的一句话，重复很多遍。" * 10
    print(f"  文本长度 {len(long_text)} 字；wrap_lines(折行) = "
          f"{len(hl_layout.wrap_lines(long_text, 18.0, 100.0))} 行，"
          f"qa = {qa.measure_text_lines(long_text, 18.0, 100.0)} 行")
    print("  → `_paragraph_lines(..., word_wrap=False)` 只出 1 行（契约 §4.3 规定），"
          "此时**不**参与等价，属有意分歧，不是漂移。")

    print("\n== 判定 ==")
    if problems:
        for x in problems[:20]:
            print(f"  ✗ {x}")
        print(f"  共 {len(problems)} 项")
        return 1
    print("  行数等价 + 三条不变量全部通过")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
