"""`pptgen` —— pptx 双向链路的命令行入口（方向 B：给 agent 调）。

进程契约（契约 §8.1）
--------------------
- **stdout 恒为一行 JSON**（`ensure_ascii=False`）；人类可读进度/警告一律走 stderr。
  理由：调用方是 bash 里的 agent，`json.loads(stdout)` 必须永远成立，混进一行日志就解析失败。
- **永不上抛裸堆栈**：所有异常收敛成 `{"ok": false, "error": {code, message, hint}}`
  + 非 0 退出码。堆栈只写 stderr，供人 debug。
- 退出码：`0` 成功｜`2` 参数错误（argparse 原生）｜`3` 输入问题｜`4` 缺外部依赖｜`5` 内部错误。

子命令：`import` / `animate` / `video` / `export` / `deck`。

规避已知问题（本轮不修，见 KNOWLEDGE.md 已知风险表）
----------------------------------------------------
- **F1**（`export_pages` 会拆掉调用线程的 COM apartment）：`import` 在 COM 导出之后
  **不再触碰任何 COM**，所以本 CLI 不受影响；若将来要在导出后复用 COM 代理，
  必须先重新 `pythoncom.CoInitialize()`。
- **L4**（播放器不校验底图存在）：`animate` 落盘前逐个 `os.path.isfile` 校验底图，
  缺图即 `IR_MISMATCH`，不让黑帧流到 `video`。
"""

import argparse
import dataclasses
import json
import os
import re
import sys
import traceback

import hl_anim
import hl_layout
import pptx_io
import qa
import video

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
OUTPUT_ROOT = os.path.join(PROJECT_ROOT, "output")

# 契约 §9 的错误码 → 退出码（实现期不得自行扩写）
_EXIT_CODE = {
    "BAD_ARGS": 2,
    "PPTX_NOT_FOUND": 3, "PPTX_UNREADABLE": 3, "PPTX_ENCRYPTED": 3, "PPTX_EMPTY": 3,
    "IR_MISMATCH": 3, "TEMPLATE_INVALID": 3,
    "NO_POWERPOINT": 4, "NO_FFMPEG": 4, "NO_BROWSER": 4,
    "COM_EXPORT_FAILED": 5, "INTERNAL": 5,
}

_NO_POWERPOINT_HINT = ("安装 Microsoft Office（含 PowerPoint）；或改用 --mode redesign")
_UNSAFE_NAME = re.compile(r'[\\/:*?"<>|]')


def _log(msg: str) -> None:
    print(msg, file=sys.stderr, flush=True)


def _write_json(path: str, obj) -> None:
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=1)


def _emit(payload: dict) -> None:
    """stdout 只此一行。"""
    sys.stdout.write(json.dumps(payload, ensure_ascii=False) + "\n")
    sys.stdout.flush()


def _emit_error(cmd: str, code: str, message: str, hint: str = "") -> int:
    _emit({"ok": False, "cmd": cmd, "error": {"code": code, "message": message, "hint": hint}})
    return _EXIT_CODE.get(code, 5)


def _safe_stem(path: str) -> str:
    stem = os.path.splitext(os.path.basename(path))[0]
    return _UNSAFE_NAME.sub("_", stem).strip() or "deck"


def _default_import_out(pptx_path: str) -> str:
    return os.path.join(OUTPUT_ROOT, "pptx_src", _safe_stem(pptx_path))


def _canvas_height_px(deck) -> int:
    if deck.width_emu <= 0:
        raise pptx_io.PptxError("deck.json 的画布宽度非法（为 0）", "IR_MISMATCH",
                                "重新执行 pptgen import 生成 deck.json")
    return round(deck.export_width_px * deck.height_emu / deck.width_emu)


def _deck_dir(args) -> str:
    return os.path.abspath(getattr(args, "dir", None) or ".")


def _load_deck_from_dir(dir_path: str):
    deck_path = os.path.join(dir_path, "deck.json")
    if not os.path.isfile(deck_path):
        raise pptx_io.PptxError(f"目录里没有 deck.json：{dir_path}", "IR_MISMATCH",
                                "先执行 pptgen import 生成 deck.json")
    return pptx_io.load_deck(deck_path)


def _require_bg(deck, dir_path: str) -> list:
    """取出底图相对路径并**逐个校验文件存在**（规避审计 L4：缺图会静默出黑帧）。"""
    bgs = [p.get("bg") for p in deck.pages]
    if not bgs or any(b is None for b in bgs):
        raise pptx_io.PptxError(
            "deck.json 里没有底图（bg 为空）", "IR_MISMATCH",
            "faithful 模式需要 PowerPoint 底图；请去掉 --no-com 重新 import")
    missing = [b for b in bgs if not os.path.isfile(os.path.join(dir_path, b))]
    if missing:
        raise pptx_io.PptxError(
            f"底图文件缺失（{len(missing)}/{len(bgs)}）：{missing[0]}", "IR_MISMATCH",
            "重新执行 pptgen import 生成底图")
    return bgs


def _pages_units(deck, skipped: list | None = None):
    return [hl_layout.build_units(hl_layout.page_shapes(p, deck), skipped=skipped)
            for p in deck.pages]


# ---------------------------------------------------------------- 设计流水线

def _design_pipeline(out_dir: str, topic: str, slides: list, image_map: dict | None = None):
    """复用现有 outline/style/html_gen 流水线产出一份设计稿 HTML（方向 A 的产物）。"""
    import html_gen
    import style as style_mod

    style = style_mod.decide_style(topic, slides)
    doc = html_gen.generate_html_deck(topic, slides, image_map or {}, style)
    path = os.path.join(out_dir, "deck.html")
    with open(path, "w", encoding="utf-8") as f:
        f.write(doc)
    return path


def _extract_text(deck) -> str:
    """把导入稿的讲解单元文字拼成一段文本，喂给 redesign 流水线。"""
    lines = []
    for page in deck.pages:
        texts = [u.text for u in hl_layout.build_units(hl_layout.page_shapes(page, deck))
                 if u.text.strip()]
        if texts:
            lines.append("｜".join(texts))
    return "\n".join(lines)


# ---------------------------------------------------------------- import

def _cmd_import(args) -> dict:
    if args.no_com and args.mode == "faithful":
        raise pptx_io.PptxError(
            "--no-com 与 --mode faithful 不能同时使用", "BAD_ARGS",
            "faithful 模式需要 PowerPoint 导出底图；请去掉 --no-com，或用 --mode redesign")
    if args.pages is not None and args.pages <= 0:
        raise pptx_io.PptxError(f"--pages 需为正整数：{args.pages}", "BAD_ARGS",
                                "例如 --pages 10")

    deck, skipped = pptx_io.read_pages(args.pptx, mode=args.mode,
                                       export_width_px=args.width)
    kept = sum(len(p["shapes"]) for p in deck.pages)
    _log(f"读入 {len(deck.pages)} 页 / 保留 {kept} 个形状 / 过滤 {len(skipped)} 个")

    if args.pages and args.pages < len(deck.pages):
        deck.pages = deck.pages[:args.pages]
        _log(f"--pages {args.pages}：只保留前 {len(deck.pages)} 页")

    out = os.path.abspath(args.out or _default_import_out(args.pptx))
    os.makedirs(out, exist_ok=True)
    _log(f"输出目录：{out}")

    bg_count = 0
    if not args.no_com:
        if not pptx_io.powerpoint_available():
            raise pptx_io.PptxError("未检测到 PowerPoint，无法导出保真底图",
                                    "NO_POWERPOINT", _NO_POWERPOINT_HINT)
        _log(f"用 PowerPoint 无窗口导出底图（{args.width}px 宽）…")
        # F1：export_pages 会拆掉本线程的 COM apartment；此处之后不再使用 COM
        paths = pptx_io.export_pages(args.pptx, os.path.join(out, "bg"), width=args.width)
        bg_count = len(paths)
        for i, page in enumerate(deck.pages):
            page["bg"] = f"bg/slide_{i + 1}.png"
        _log(f"底图 {bg_count} 张已写入 {os.path.join(out, 'bg')}")
        # 契约 §8.2「--pages N 时 deck.json 的 pages 与 bg/ 同步截断」：
        # export_pages 没有页范围参数，会整本导出；这里把本次多导的尾部清掉，
        # 避免 bg/ 里躺着 deck.json 不引用的文件（下游按文件数核对会误判）。
        surplus = [p for p in paths[len(deck.pages):] if os.path.isfile(p)]
        for p in surplus:
            os.remove(p)
        if surplus:
            _log(f"--pages {args.pages}：清掉本次多导的 {len(surplus)} 张底图")
        bg_count = len(paths) - len(surplus)

    # 布局期丢弃的形状也要有出口（审计 M3）：契约 §4.3 要求 inner_w_pt<=0 记 warning，
    # 不接收集器的话用户只会看到"有些文字没有高亮"，排查时毫无线索。
    layout_skipped: list = []
    units_by_page = _pages_units(deck, layout_skipped)
    total_units = sum(len(u) for u in units_by_page)
    for item in layout_skipped:
        _log(f"[警告] 第{item['page_index'] + 1}页 形状'{item['shape_name']}' "
             f"没有产出高亮（{item['reason']}）")
    unit_warns = _unit_warning_summary(units_by_page)
    if unit_warns:
        _log(f"[警告] 讲解单元的精度提示：{'、'.join(unit_warns)}")

    _write_json(os.path.join(out, "deck.json"), pptx_io.deck_to_dict(deck))
    _write_json(os.path.join(out, "units.json"), [
        {"page_index": page["index"],
         "units": [dataclasses.asdict(u) for u in units]}
        for page, units in zip(deck.pages, units_by_page)])

    data = {
        "out": out, "mode": deck.mode, "width": deck.export_width_px,
        "pages": len(deck.pages), "bg": bg_count, "units": total_units,
        "skipped": len(skipped) + len(layout_skipped),
        "warnings": (_skip_summary(skipped) + _skip_summary(layout_skipped)
                     + unit_warns),
        "deck_json": os.path.join(out, "deck.json"),
    }
    if deck.mode == "redesign":
        import outline
        text = _extract_text(deck)
        _log(f"redesign：用提取到的 {len(text)} 字文本重跑设计流水线…")
        slides = outline.generate_outline_from_text(text)
        html_path = _design_pipeline(out, _safe_stem(args.pptx), slides)
        data["deck_html"] = html_path
        _log(f"设计稿：{html_path}")
    return data


def _skip_summary(skipped: list) -> list:
    counts: dict[str, int] = {}
    for item in skipped:
        counts[item["reason"]] = counts.get(item["reason"], 0) + 1
    return [f"{k}×{v}" for k, v in sorted(counts.items())]


def _unit_warning_summary(units_by_page: list) -> list:
    """把 Unit 级警告汇总成 `名字×条数`，与 `_skip_summary` 同格式便于并排读。"""
    counts: dict[str, int] = {}
    for units in units_by_page:
        for u in units:
            for name in u.warnings:
                counts[name] = counts.get(name, 0) + 1
    return [f"{k}×{v}" for k, v in sorted(counts.items())]


# ---------------------------------------------------------------- animate

def _cmd_animate(args) -> dict:
    dir_path = _deck_dir(args)
    deck = _load_deck_from_dir(dir_path)
    bgs = _require_bg(deck, dir_path)
    pages_units = _pages_units(deck)

    title = _safe_stem(deck.source_pptx)
    player_dir = os.path.join(dir_path, "player")
    player = hl_anim.build_player(
        player_dir, bgs, pages_units, title=title,
        dim=args.dim, auto_step_ms=args.auto_ms,
        canvas_width_px=deck.export_width_px, canvas_height_px=_canvas_height_px(deck),
        bg_base_dir=dir_path)  # deck.json 里的 bg 相对 <dir>，播放器在 <dir>/player/
    _log(f"播放器：{player}")
    return {"player": player, "pages": len(bgs),
            "units": sum(len(u) for u in pages_units), "title": title}


# ---------------------------------------------------------------- video

def _cmd_video(args) -> dict:
    if video.ffmpeg_path() is None:
        raise pptx_io.PptxError("未找到 ffmpeg", "NO_FFMPEG",
                                "winget install ffmpeg")
    dir_path = _deck_dir(args)
    deck = _load_deck_from_dir(dir_path)

    if args.mode == "step":
        player = os.path.join(dir_path, "player", "index.html")
        if not os.path.isfile(player):
            raise pptx_io.PptxError(f"还没有播放器：{player}", "IR_MISMATCH",
                                    "先执行 pptgen animate")
        pages_units = _pages_units(deck)
        steps = [(pi, si) for pi, units in enumerate(pages_units)
                 for si in range(1, len(units) + 1)]
        _log(f"step 模式：截 {len(steps)} 帧（Σ units）…")
        try:
            shots = hl_anim.shot_player(player, os.path.join(dir_path, "frames"), steps)
        except RuntimeError as exc:
            if "Chrome/Edge" in str(exc):
                raise pptx_io.PptxError("未找到可用的 Chrome/Edge", "NO_BROWSER",
                                        "安装 Google Chrome，或 `playwright install chromium`") from exc
            raise
    else:
        bgs = _require_bg(deck, dir_path)
        shots = [os.path.join(dir_path, b) for b in bgs]
        _log(f"page 模式：用 {len(shots)} 张底图直出…")

    out = os.path.abspath(args.out or os.path.join(dir_path, "out.mp4"))
    _log(f"ffmpeg 合成 {len(shots)} 段 → {out}")
    video.synthesize(shots, out, seconds=args.sec, fade=args.fade, fps=args.fps,
                     progress_cb=lambda i, n: _log(f"  片段 {i}/{n}"))
    size = os.path.getsize(out) if os.path.isfile(out) else 0
    return {"video": out, "clips": len(shots), "seconds": args.sec, "fps": args.fps,
            "fade": args.fade, "bytes": size, "mode": args.mode}


# ---------------------------------------------------------------- export

def _cmd_export(args) -> dict:
    deck_path = os.path.abspath(args.deck_json)
    if not os.path.isfile(deck_path):
        raise pptx_io.PptxError(f"找不到文件：{deck_path}", "PPTX_NOT_FOUND",
                                "检查路径是否正确（建议用绝对路径）")
    deck = pptx_io.load_deck(deck_path)

    template = None
    if args.template:
        template = os.path.abspath(args.template)
        if not os.path.isfile(template) or not template.lower().endswith(".pptx"):
            raise pptx_io.PptxError(f"模板不可用：{args.template}", "TEMPLATE_INVALID",
                                    "模板需为标准 .pptx 文件")

    out = os.path.abspath(args.out or os.path.join(
        os.path.dirname(deck_path), _safe_stem(deck_path) + ".pptx"))

    import pptx_out  # 步 7 落地；懒导入避免 animate/video 被它的依赖牵连
    pptx_out.build_deck_pptx(deck, out, template=template, no_com=args.no_com)
    _log(f"已导出：{out}")

    check = qa.check_pptx(out, len(deck.pages))
    for w in check["warnings"]:
        _log(f"[警告] {w}")
    for e in check["errors"]:
        _log(f"[错误] {e}")
    if check["errors"]:
        raise pptx_io.PptxError(
            f"导出的 pptx 自检未通过（{len(check['errors'])} 项）：{check['errors'][0]}",
            "INTERNAL", "附命令与 deck.json 反馈；这是导出器的问题")
    return {"pptx": out, "pages": len(deck.pages), "template": bool(template),
            "warnings": check["warnings"], "used_estimate": check["used_estimate"],
            "errors": []}


# ---------------------------------------------------------------- deck

def _cmd_deck(args) -> dict:
    import outline

    out = os.path.abspath(args.out or os.path.join(OUTPUT_ROOT, "pptx_src",
                                                   _safe_stem(args.topic)))
    os.makedirs(out, exist_ok=True)
    _log(f"生成大纲：{args.topic}")
    slides = outline.generate_outline(args.topic)
    _log(f"共 {len(slides)} 页，产出设计稿…")
    html_path = _design_pipeline(out, args.topic, slides)
    _log(f"设计稿：{html_path}")
    return {"out": out, "deck_html": html_path, "pages": len(slides), "topic": args.topic}


# ---------------------------------------------------------------- 入口

def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="pptgen", description="pptx 双向链路：导入成会动的演示 / 导出可编辑 pptx")
    sub = p.add_subparsers(dest="cmd")

    imp = sub.add_parser("import", help="pptx → deck.json + 底图 + 讲解单元")
    imp.add_argument("pptx", help="源 .pptx 路径")
    imp.add_argument("--out", help="输出目录（默认 output/pptx_src/<稿名>/）")
    imp.add_argument("--mode", choices=["faithful", "redesign"], default="faithful")
    imp.add_argument("--width", type=int, choices=[1280, 1920, 2560], default=1920)
    imp.add_argument("--pages", type=int, help="只导前 N 页")
    imp.add_argument("--no-com", action="store_true", help="不导出 PowerPoint 底图")
    imp.set_defaults(func=_cmd_import)

    ani = sub.add_parser("animate", help="生成高亮播放器")
    ani.add_argument("dir", help="import 的输出目录")
    ani.add_argument("--dim", type=float, default=0.25, help="降暗比例，0<dim<0.6")
    ani.add_argument("--auto-ms", type=int, default=2000, help="自动讲解步进毫秒")
    ani.add_argument("--highlight", action="store_true", help="显式声明要高亮（默认即开）")
    ani.set_defaults(func=_cmd_animate)

    vid = sub.add_parser("video", help="导出 MP4")
    vid.add_argument("dir", help="import 的输出目录")
    vid.add_argument("--sec", type=float, default=4.0, help="每页时长（秒）")
    vid.add_argument("--fps", type=int, default=25)
    vid.add_argument("--fade", type=float, default=0.8, help="转场时长（秒），需 0<fade<sec")
    vid.add_argument("--mode", choices=["page", "step"], default="page")
    vid.add_argument("-o", "--out", help="输出 mp4（默认 <dir>/out.mp4）")
    vid.set_defaults(func=_cmd_video)

    exp = sub.add_parser("export", help="deck.json → 可编辑 pptx")
    exp.add_argument("deck_json", help="deck.json 路径")
    exp.add_argument("-o", "--out", help="输出 pptx（默认同目录 <deck.json 主名>.pptx）")
    exp.add_argument("--template", help="自定义母版 .pptx")
    exp.add_argument("--no-com", action="store_true", help="不支持的形状跳过而非报错")
    exp.set_defaults(func=_cmd_export)

    dk = sub.add_parser("deck", help="主题 → AI 设计稿（复用现有流水线）")
    dk.add_argument("topic", help="主题文字")
    dk.add_argument("--out", help="输出目录")
    dk.set_defaults(func=_cmd_deck)
    return p


def main(argv=None) -> int:
    parser = _build_parser()
    args = parser.parse_args(argv)
    if getattr(args, "func", None) is None:
        parser.print_usage(sys.stderr)
        _emit_error("", "BAD_ARGS", "没有指定子命令", "可用：import / animate / video / export / deck")
        return 2

    cmd = args.cmd
    try:
        data = args.func(args)
    except pptx_io.PptxError as exc:
        return _emit_error(cmd, exc.code, exc.message, exc.hint)
    except ValueError as exc:
        return _emit_error(cmd, "BAD_ARGS", str(exc), "检查参数取值（可用 --help 查看）")
    except RuntimeError as exc:
        return _emit_error(cmd, "INTERNAL", str(exc),
                           "检查 .env 里的模型渠道配置；或附命令与 deck.json 反馈")
    except KeyboardInterrupt:
        return _emit_error(cmd, "INTERNAL", "已被用户中断", "重跑该命令")
    except Exception as exc:  # noqa: BLE001  堆栈只进 stderr，stdout 仍是合法 JSON
        traceback.print_exc(file=sys.stderr)
        return _emit_error(cmd, "INTERNAL", f"内部错误：{type(exc).__name__}: {exc}",
                           "这属于 bug，请附命令与 deck.json")
    _emit({"ok": True, "cmd": cmd, "data": data})
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
