"""独立验证 ⑥ · M1 验收数字口径：不靠 python-pptx API，直接用**裸 XML** 数形状。

实现者报（D9）：「原始 41 形状 = 40 文本框 + 1 图表；40 个文本框里 18 个文本为空，
过滤后保留 23（22 text + 1 chart）」。本探针解压 pptx、逐个 `slideN.xml` 数
`p:sp`/`p:graphicFrame`/`p:pic`/`p:grpSp` 元素，再与 `read_pages` 的产出对账。

一条命令：./.venv/Scripts/python.exe tools/probes/indep_m1_counts.py
"""

import os
import re
import sys
import zipfile
from collections import Counter

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

import pptx_io  # noqa: E402

SRC = os.path.join(ROOT, "output", "b_multislide.pptx")
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"


def raw_counts(path):
    """裸 XML：逐页数顶层元素与文本情况。"""
    from lxml import etree
    per_page = []
    with zipfile.ZipFile(path) as z:
        names = [n for n in z.namelist() if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)]
        names.sort(key=lambda n: int(re.search(r"(\d+)", n.split("/")[-1]).group(1)))
        for n in names:
            root = etree.fromstring(z.read(n))
            tree = root.find(f"{P}cSld/{P}spTree")
            kinds = Counter()
            textful = 0
            empty = 0
            for el in tree:
                tag = etree.QName(el).localname
                if tag == "sp":
                    kinds["sp"] += 1
                    ts = [t.text or "" for t in el.iter(f"{A}t")]
                    if any(t.strip() for t in ts):
                        textful += 1
                    else:
                        empty += 1
                elif tag == "graphicFrame":
                    if el.find(f".//{C}chart") is not None:
                        kinds["chart"] += 1
                    elif el.find(f".//{A}tbl") is not None:
                        kinds["table"] += 1
                    else:
                        kinds["graphicFrame?"] += 1
                elif tag in ("pic", "grpSp", "cxnSp", "contentPart"):
                    kinds[tag] += 1
            per_page.append({"file": n, "kinds": kinds, "textful": textful, "empty": empty})
    return per_page


def main() -> int:
    per_page = raw_counts(SRC)
    tot = Counter()
    raw_shapes = 0
    for p in per_page:
        tot.update(p["kinds"])
        raw_shapes += sum(p["kinds"].values())

    print(f"素材：{os.path.basename(SRC)}（{os.path.getsize(SRC)} 字节）")
    print("\n== 裸 XML 逐页 ==")
    for i, p in enumerate(per_page, 1):
        print(f"  第{i:2}页 顶层元素 {sum(p['kinds'].values()):2}  "
              f"{dict(p['kinds'])}  有文本sp={p['textful']} 空sp={p['empty']}")
    print(f"\n  合计：顶层元素 {raw_shapes} 个 → {dict(tot)}")

    sp = tot.get("sp", 0)
    print(f"\n== 与实现者声明对账 ==")
    print(f"  文本框（p:sp）  = {sp}    （实现者报 40）"
          f"  {'一致' if sp == 40 else '**不一致**'}")
    print(f"  图表（chart）   = {tot.get('chart', 0)}    （实现者报 1）"
          f"  {'一致' if tot.get('chart', 0) == 1 else '**不一致**'}")
    print(f"  原始形状总数     = {raw_shapes}   （实现者报 41）"
          f"  {'一致' if raw_shapes == 41 else '**不一致**'}")

    # read_pages 侧
    deck, skipped = pptx_io.read_pages(SRC)
    kept = sum(len(pg["shapes"]) for pg in deck.pages)
    kinds = Counter()
    for pg in deck.pages:
        for s in pg["shapes"]:
            kinds[s.kind] += 1
    reasons = Counter(s["reason"] for s in skipped)
    print(f"\n== read_pages 产出 ==")
    print(f"  页数 = {len(deck.pages)}（裸 XML 页数 {len(per_page)}）"
          f"  {'一致' if len(deck.pages) == len(per_page) else '**不一致**'}")
    print(f"  保留形状 = {kept} → {dict(kinds)}   （实现者报 23 / 22 text + 1 chart）")
    print(f"  skipped = {len(skipped)} 条 → {dict(reasons)}   （实现者报 18 empty_text）")
    print(f"  画布 EMU = {deck.width_emu} x {deck.height_emu}")
    print(f"  对账：保留 {kept} + 过滤 {len(skipped)} = {kept + len(skipped)}"
          f"  vs 裸 XML 顶层 {raw_shapes}"
          f"  {'一致' if kept + len(skipped) == raw_shapes else '**不一致**'}")

    ok = (sp == 40 and tot.get("chart", 0) == 1 and raw_shapes == 41
          and kept == 23 and kinds.get("text") == 22 and kinds.get("chart") == 1
          and len(skipped) == 18 and dict(reasons) == {"empty_text": 18})
    print(f"\n== 判定：{'口径与实现者完全一致（独立核实通过）' if ok else '存在口径差异，见上'} ==")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
