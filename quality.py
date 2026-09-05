"""deck 级内容质量纯函数：重复页检测 / 内容过瘦（全部 warning 级，不阻断导出）。

思路来自参考项目 ai-ppt 的 quality.py：归一化 bigram Jaccard 判重复。
刻意保持无 LLM、无第三方依赖——质检要能每次导出前零成本现算；
warning 只提示用户核对，不让 AI 自动"砍重复页"（误伤风险大于收益）。
"""


def _bigrams(text: str) -> set[str]:
    """去空白后取相邻字符对集合；单字文本退化为单字集合。"""
    t = "".join(text.split())
    if len(t) < 2:
        return {t} if t else set()
    return {t[i:i + 2] for i in range(len(t) - 1)}


def _jaccard(a: set, b: set) -> float:
    if not a or not b:
        return 0.0
    inter = len(a & b)
    return inter / (len(a) + len(b) - inter)


def find_duplicate_pages(slides: list[dict], title_th: float = 0.85, body_th: float = 0.8) -> list[dict]:
    """页对页两两比较，返回疑似重复组 [{a, b, kind, score}]，kind ∈ title/body。"""
    out = []
    n = len(slides)
    for i in range(n):
        for j in range(i + 1, n):
            t_sim = _jaccard(_bigrams(slides[i].get("title", "")), _bigrams(slides[j].get("title", "")))
            if t_sim >= title_th:
                out.append({"a": i, "b": j, "kind": "title", "score": round(t_sim, 2)})
                continue
            b_sim = _jaccard(
                _bigrams(" ".join(slides[i].get("points", []) or [])),
                _bigrams(" ".join(slides[j].get("points", []) or [])),
            )
            if b_sim >= body_th:
                out.append({"a": i, "b": j, "kind": "body", "score": round(b_sim, 2)})
    return out


def find_thin_pages(slides: list[dict], min_points: int = 2, min_chars: int = 15) -> list[int]:
    """内容页要点过少或过短的下标列表；封面/目录/章节/结尾/数据页豁免。"""
    thin = []
    for i, s in enumerate(slides):
        if s.get("type") not in (None, "", "content"):
            continue
        pts = [p for p in (s.get("points") or []) if str(p).strip()]
        if len(pts) < min_points or sum(len(str(p)) for p in pts) < min_chars:
            thin.append(i)
    return thin


def check_deck(slides: list[dict]) -> dict:
    """汇总 deck 级质量报告（现算不落库，供 /api/quality 与 ready 日志用）。"""
    return {
        "duplicates": find_duplicate_pages(slides),
        "thin": find_thin_pages(slides),
    }
