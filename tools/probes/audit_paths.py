"""审计复现 · 底图路径的编码链与 F5（..%2f 绕过）是否真的可利用。

用法：
    PYTHONIOENCODING=utf-8 ./.venv/Scripts/python.exe tools/probes/audit_paths.py

只做字符串运算，不读不写文件、不起浏览器。
"""
import os
import sys
from urllib.parse import quote

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

import hl_anim  # noqa: E402


def main():
    print("== _check_bg_path 放行/拦截 + _web_path 编码链 ==")
    cases = [
        "bg/..%2f..%2fetc.png",
        "bg/%2e%2e/x.png",
        "bg/1.png",
        "..%2f..%2fetc.png",
        "%2e%2e%2fx.png",
        "bg/sub/../../../x.png",
        "bg/./1.png",
        ".",
    ]
    for p in cases:
        try:
            hl_anim._check_bg_path(p)
            verdict = "放行"
        except Exception as exc:                               # noqa: BLE001
            verdict = f"拦截({type(exc).__name__}/{getattr(exc, 'code', '')})"
        web = hl_anim._web_path(p)
        one_decode = quote(p.replace("\\", "/"), safe="/")
        print(f"  {p!r:28s} {verdict:22s} _web_path -> {web!r}")

    print("\n== 关键推理：%2f 被双重编码后不再是路径分隔符 ==")
    p = "bg/..%2f..%2fetc.png"
    web = hl_anim._web_path(p)
    print(f"  原始 deck.json 里的 bg 值 : {p!r}")
    print(f"  进 src 的字符串           : {web!r}")
    print("  浏览器对 URL 路径只做一次百分号解码 → 解码结果是字面文件名 "
          "'..%2f..%2fetc.png'（%2f 不是分隔符）")
    print("  且 % 已被 _web_path 先编码成 %25，所以连这一次解码也解不回 '/'")
    print("  → F5 担心的『绕出 out_dir』在 _web_path 之后不可达；"
          "但 _check_bg_path 本身确实不做百分号解码（若将来别处直接拿它拼 URL 就会踩到）")

    print("\n== 对照组：确实会被拦截的形态 ==")
    for p in ["../x.png", "..\\x.png", "C:/x.png", "//srv/x.png"]:
        try:
            hl_anim._check_bg_path(p)
            print(f"  {p!r:16s} 放行（意外）")
        except Exception as exc:                               # noqa: BLE001
            print(f"  {p!r:16s} 拦截 {getattr(exc, 'code', type(exc).__name__)}")


if __name__ == "__main__":
    main()
