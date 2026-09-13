#!/usr/bin/env python
"""全功能端到端验证：驱动工作台 HTTP API 走通所有功能，产出带证据的验证报告。

前置：工作台已启动（`python app.py`，默认 http://127.0.0.1:5000）；路线① 需模型网关可用。

用法（分阶段跑，避免单次太久）：
    .venv/Scripts/python.exe tools/e2e_full_test.py gen      # 阶段1 分步确认 + 生成
    .venv/Scripts/python.exe tools/e2e_full_test.py edit     # 阶段2 修改类功能
    .venv/Scripts/python.exe tools/e2e_full_test.py export   # 阶段3 六种导出
    .venv/Scripts/python.exe tools/e2e_full_test.py pptx     # 阶段4 pptx 导入（路线②）
    .venv/Scripts/python.exe tools/e2e_full_test.py misc     # 阶段5 模板/品牌/换色/质检/渠道
    .venv/Scripts/python.exe tools/e2e_full_test.py all
"""

import json
import os
import sys
import time

# 本机 Windows 默认 stdout 是 GBK；WSL→Windows 的 PYTHONIOENCODING 不会自动转发，
# 因此脚本内自己拉 UTF-8（否则带中文/emoji 的输出会 UnicodeEncodeError 中断）
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception:      # noqa: BLE001
        pass

import httpx

BASE = os.environ.get("PPTGEN_BASE", "http://127.0.0.1:5000")
TOPIC = os.environ.get("PPTGEN_TOPIC", "时间管理：让每一天更有效率")
REPORT = []


# --------------------------------------------------------------------------- 工具

def mark(name, ok, detail=""):
    REPORT.append((name, bool(ok), detail))
    print(f"  {'[OK]' if ok else '[X ]'} {name}" + (f"  -- {detail}" if detail else ""), flush=True)


def get(path, timeout=180):
    r = httpx.get(BASE + path, timeout=timeout)
    try:
        return r.status_code, r.json()
    except Exception:
        return r.status_code, {"_raw": r.text[:200]}


def post(path, data=None, timeout=900):
    r = httpx.post(BASE + path, json=data, timeout=timeout)
    try:
        return r.status_code, r.json()
    except Exception:
        return r.status_code, {"_raw": r.text[:200]}


def delete(path, timeout=180):
    r = httpx.delete(BASE + path, timeout=timeout)
    try:
        return r.status_code, r.json()
    except Exception:
        return r.status_code, {}


def wait_phase(target=("ready",), timeout=1200, resume_review=True, label=""):
    """轮询 /api/status 到目标 phase；review 态自动放行（记录次数）。"""
    t0 = time.time()
    last = None
    resumes = 0
    while time.time() - t0 < timeout:
        sc, d = get("/api/status")
        ph = d.get("phase")
        if ph != last:
            extra = f" await={d.get('await_step')}" if d.get("await_step") else ""
            print(f"    · phase={ph}{extra}  ({time.time()-t0:.0f}s)", flush=True)
            last = ph
        if ph in target:
            d["_resumes"] = resumes
            return d
        if ph == "review" and resume_review:
            post("/api/continue", {})
            resumes += 1
        time.sleep(3)
    print(f"    ! 超时未到 {target}（{label}）", flush=True)
    return None


def artifacts():
    sc, d = get("/api/artifacts")
    return d.get("items", d.get("artifacts", [])) if isinstance(d, dict) else []


# --------------------------------------------------------------------------- 阶段 1

def stage_gen():
    print("\n【阶段1】分步确认 + 全自动生成", flush=True)

    sc, d = post("/api/stepwise", {"enabled": True})
    mark("POST /api/stepwise（开分步确认）", sc == 200 and d.get("stepwise") is True, f"HTTP {sc}")

    density = os.environ.get("PPTGEN_DENSITY", "balanced")
    sc, d = post("/api/generate", {"topic": TOPIC, "density": density})
    mark("POST /api/generate", sc == 200 and d.get("ok"), f"HTTP {sc}，density={density}")

    t0 = time.time()
    st = wait_phase(("ready",), label="生成")
    ok = st is not None
    mark("生成到 ready", ok, f"{time.time()-t0:.0f}s，途中放行 {st.get('_resumes') if st else '?'} 次")
    if not ok:
        return None

    slides = st.get("slides") or []
    mark("大纲页数合理（6~12）", 6 <= len(slides) <= 12, f"{len(slides)} 页")
    mark("已选风格", bool(st.get("style_name") or st.get("style")), str(st.get("style_name")))
    mark("设计稿 html_path 存在", bool(st.get("html_path")), str(st.get("html_path")))
    mark("每页有标题", all((s.get("title") or "").strip() for s in slides))

    post("/api/stepwise", {"enabled": False})
    return st


# --------------------------------------------------------------------------- 阶段 2

def stage_edit():
    print("\n【阶段2】修改类功能", flush=True)
    sc, st = get("/api/status")
    if st.get("phase") != "ready":
        mark("阶段2 前置（需 ready）", False, f"phase={st.get('phase')}")
        return

    n0 = len(st.get("slides") or [])

    # 整篇修改
    t0 = time.time()
    sc, d = post("/api/refine", {"instruction": "整篇语言更简洁，每条要点不超过 20 字"})
    ok = sc == 200
    if ok:
        wait_phase(("ready",), label="整篇 refine")
    sc2, st2 = get("/api/status")
    mark("POST /api/refine（整篇）", ok and st2.get("phase") == "ready",
         f"HTTP {sc}，{time.time()-t0:.0f}s")

    # 单页文字编辑
    sc, d = post("/api/slide/0/text", {"title": "（测试）时间管理导论", "points": ["要点A", "要点B"]})
    sc2, st2 = get("/api/status")
    t_ok = (st2.get("slides") or [{}])[0].get("title") == "（测试）时间管理导论"
    mark("POST /api/slide/<i>/text（改单页）", sc == 200 and t_ok, f"HTTP {sc}")

    # 排序
    order = list(range(n0))[::-1]
    sc, d = post("/api/slide/reorder", {"order": order})
    sc2, st2 = get("/api/status")
    titles = [s.get("title") for s in (st2.get("slides") or [])]
    ok = sc == 200 and len(titles) == n0
    mark("POST /api/slide/reorder（倒序）", ok, f"HTTP {sc}，首行现在=「{titles[0] if titles else ''}」")
    post("/api/slide/reorder", {"order": list(range(n0))})   # 还原

    # 加页
    sc, d = post("/api/slide/add", {"after": 0})
    sc2, st2 = get("/api/status")
    n1 = len(st2.get("slides") or [])
    mark("POST /api/slide/add（插入一页）", sc == 200 and n1 == n0 + 1, f"{n0} -> {n1}")

    # 删页
    sc, d = delete("/api/slide/1/delete")   # 注意路由带 /delete 后缀，方法为 DELETE
    sc2, st2 = get("/api/status")
    n2 = len(st2.get("slides") or [])
    mark("DELETE /api/slide/<i>（删页）", sc in (200, 204) and n2 == n1 - 1, f"{n1} -> {n2}")

    # 定点修改（带 target）
    sc, d = post("/api/refine", {"instruction": "这条要点改得更口语化",
                                "target": {"slide": 0}})
    ok = sc == 200
    if ok:
        wait_phase(("ready",), label="定点 refine")
    sc2, st2 = get("/api/status")
    mark("POST /api/refine（定点 target）", ok and st2.get("phase") == "ready", f"HTTP {sc}")

    # 重新设计（重生成 HTML）
    t0 = time.time()
    sc, d = post("/api/redesign", {})
    ok = sc == 200
    if ok:
        wait_phase(("ready",), label="redesign")
    sc2, st2 = get("/api/status")
    mark("POST /api/redesign（重新出设计稿）", ok and bool(st2.get("html_path")),
         f"HTTP {sc}，{time.time()-t0:.0f}s")


# --------------------------------------------------------------------------- 阶段 3

def stage_export():
    print("\n【阶段3】六种导出", flush=True)
    sc, st = get("/api/status")
    if st.get("phase") != "ready":
        mark("阶段3 前置（需 ready）", False, f"phase={st.get('phase')}")
        return

    before = {a.get("path") or a.get("url") or a.get("name") for a in artifacts()}

    for path, label, wait in [
        ("/api/export", "pptx（图片版）", None),
        ("/api/export_html", "HTML 演示版", None),
        ("/api/export_txt", "大纲 txt", None),
        ("/api/export_pdf", "PDF（走 PowerPoint COM）", None),
        ("/api/export_animation", "教学动画（逐元素揭示）", None),
        ("/api/export_video", "视频 MP4（后台合成）", 420),
    ]:
        t0 = time.time()
        sc, d = post(path, {}, timeout=600)
        ok = sc == 200
        extra = f"HTTP {sc}"
        if ok and wait:
            time.sleep(wait * 0.35)          # 先等一段，再确认产物
        if ok:
            got = d.get("path") or d.get("file") or d.get("name") or ""
            extra += f"，{time.time()-t0:.0f}s" + (f"，{got}" if got else "")
        else:
            extra += f"，{str(d)[:90]}"
        mark(f"POST {path}（{label}）", ok, extra)

    time.sleep(60)                            # 给视频合成留时间
    after = artifacts()
    new = [a for a in after if (a.get("path") or a.get("url") or a.get("name")) not in before]
    mark("导出产物被收录进 /api/artifacts", len(new) >= 4, f"新增 {len(new)} 项")


# --------------------------------------------------------------------------- 阶段 4

def stage_pptx():
    print("\n【阶段4】pptx 导入 → 高亮播放器（路线②）", flush=True)
    src = "output/b_multislide.pptx"
    if not os.path.isfile(src):
        cands = [f for f in os.listdir("output") if f.endswith(".pptx")]
        if not cands:
            mark("阶段4 前置（需一份 pptx）", False, "output/ 下没有 pptx")
            return
        src = os.path.join("output", sorted(cands)[0])
    print(f"    素材：{src}", flush=True)

    t0 = time.time()
    with open(src, "rb") as f:
        r = httpx.post(BASE + "/api/pptx/import",
                       files={"file": (os.path.basename(src), f,
                                       "application/vnd.openxmlformats-officedocument.presentationml.presentation")},
                       timeout=600)
    mark("POST /api/pptx/import（上传）", r.status_code == 200, f"HTTP {r.status_code}")

    last = None
    for _ in range(120):
        sc, d = get("/api/status")
        p = (d.get("pptx") or {})
        s = p.get("status")
        if s != last:
            print(f"    · pptx.status={s} pages={p.get('pages')} units={p.get('units')}", flush=True)
            last = s
        if s in ("ready", "error"):
            break
        time.sleep(3)
    sc, d = get("/api/status")
    p = d.get("pptx") or {}
    mark("pptx 导入到 ready", p.get("status") == "ready",
         f"{time.time()-t0:.0f}s，{p.get('pages')} 页 / {p.get('units')} 单元 / 错误={p.get('error')}")
    mark("产出播放器路径", bool(p.get("player")), str(p.get("player")))


# --------------------------------------------------------------------------- 阶段 5

def stage_misc():
    print("\n【阶段5】模板 / 品牌 / 换色 / 质检 / 渠道 / 历史", flush=True)

    sc, d = get("/api/templates")
    tpls = d.get("templates") or d.get("items") or []
    mark("GET /api/templates", sc == 200 and len(tpls) >= 1, f"{len(tpls)} 套模板")

    if tpls:
        key = tpls[0].get("key") or tpls[0].get("id")
        sc, d = post("/api/template/select", {"key": key})
        mark("POST /api/template/select", sc == 200, f"选中 {tpls[0].get('name')}")
        post("/api/template/select", {"key": ""})         # 清除，回到 AI 自选

    sc, d = post("/api/brand", {"name": "功能测试品牌", "color": "#1F6FEB"})
    mark("POST /api/brand（品牌色）", sc == 200, f"HTTP {sc}")
    post("/api/brand", {"name": "", "color": ""})          # 清除

    sc, d = post("/api/theme", {"accent": "#E8590C"})
    mark("POST /api/theme（一键换色）", sc in (200, 409),
         f"HTTP {sc}" + ("" if sc == 200 else "（无 :root 变量的旧稿会 409，属预期）"))

    sc, d = get("/api/quality")
    mark("GET /api/quality（重复/过瘦检测）", sc == 200, json.dumps(d, ensure_ascii=False)[:80])

    sc, d = get("/api/models")
    mark("GET /api/models", sc == 200, json.dumps(d, ensure_ascii=False)[:80])

    for path in ("/api/projects", "/api/decks", "/api/artifacts"):
        sc, d = get(path)
        n = len(d.get("items") or d.get("projects") or d.get("decks") or d.get("artifacts") or [])
        mark(f"GET {path}", sc == 200, f"{n} 项")


# --------------------------------------------------------------------------- 主流程

STAGES = {"gen": stage_gen, "edit": stage_edit, "export": stage_export,
          "pptx": stage_pptx, "misc": stage_misc}


def main():
    which = sys.argv[1] if len(sys.argv) > 1 else "all"
    try:
        sc, d = get("/", timeout=10)
        mark("工作台可达", sc == 200, BASE)
    except Exception as exc:
        print(f"无法连接工作台 {BASE}：{exc}")
        return 2

    names = list(STAGES) if which == "all" else [which]
    for nm in names:
        fn = STAGES.get(nm)
        if fn is None:
            print(f"未知阶段：{nm}（可选：{'/'.join(STAGES)}/all）")
            return 2
        try:
            fn()
        except Exception as exc:                                    # noqa: BLE001
            mark(f"阶段 {nm} 异常", False, f"{type(exc).__name__}: {exc}")

    ok = sum(1 for _, o, _ in REPORT if o)
    bad = [r for r in REPORT if not r[1]]
    print("\n" + "=" * 72)
    print(f"结果：{ok}/{len(REPORT)} 通过")
    if bad:
        print("未通过：")
        for n, _, dt in bad:
            print(f"  - {n}  {dt}")
    print("=" * 72)
    return 0 if not bad else 1


if __name__ == "__main__":
    sys.exit(main())
