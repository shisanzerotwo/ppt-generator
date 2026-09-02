"""生成可视化测试报告（HTML），展示测试过程与每页成品效果。"""

import html
import os
import webbrowser


def _step_badge(ok: bool) -> str:
    return '<span class="badge ok">通过</span>' if ok else '<span class="badge fail">失败</span>'


def _slide_card(idx: int, slide: dict, image_path: str | None, ok: bool, report_dir: str) -> str:
    title = html.escape(slide.get("title", ""))
    points = slide.get("points", [])
    is_cover = idx == 0 and not points
    if is_cover:
        inner = f'<div class="cover-title">{title}</div>'
    else:
        lis = "".join(f"<li>{html.escape(p)}</li>" for p in points)
        if image_path and os.path.exists(image_path):
            rel = os.path.relpath(image_path, report_dir).replace("\\", "/")
            img_html = f'<img src="{rel}" alt="slide{idx + 1}">'
        else:
            img_html = '<div class="img-placeholder">（图片生成失败）</div>'
        inner = (
            f'<div class="layout-text"><h3>{title}</h3><ul>{lis}</ul></div>'
            f'<div class="layout-img">{img_html}</div>'
        )
    status = _step_badge(ok)
    body_cls = "cover" if is_cover else "content"
    return (
        f'<div class="slide-card"><div class="slide-head">页 {idx + 1} · {title} {status}</div>'
        f'<div class="slide-body {body_cls}">{inner}</div></div>'
    )


def build_report(topic: str, slides: list[dict], image_results: list[dict],
                 steps: list[dict], out_pptx: str, report_path: str) -> str:
    """image_results: [{path: str|None, ok: bool, seconds: float}]；
    steps: [{name, ok, seconds, detail}]。"""
    report_dir = os.path.dirname(os.path.abspath(report_path))
    total_sec = sum(s["seconds"] for s in steps)
    all_ok = all(s["ok"] for s in steps)

    step_rows = "".join(
        f"<tr><td>{html.escape(s['name'])}</td><td>{_step_badge(s['ok'])}</td>"
        f"<td>{s['seconds']:.1f}s</td><td>{html.escape(s.get('detail', ''))}</td></tr>"
        for s in steps
    )

    cards = "".join(
        _slide_card(i, slides[i],
                    image_results[i]["path"] if i < len(image_results) else None,
                    image_results[i]["ok"] if i < len(image_results) else False,
                    report_dir)
        for i in range(len(slides))
    )

    page = f"""<!DOCTYPE html>
<html lang="zh">
<head><meta charset="utf-8"><title>PPT 生成测试报告 · {html.escape(topic)}</title>
<style>
  body {{ font-family: "Microsoft YaHei", sans-serif; background:#f5f6f8; margin:0; padding:32px; color:#333; }}
  .container {{ max-width: 960px; margin: 0 auto; }}
  h1 {{ font-size: 24px; }}
  .summary {{ background:#fff; border-radius:10px; padding:20px 24px; margin-bottom:24px;
              box-shadow:0 1px 4px rgba(0,0,0,.08); display:flex; gap:32px; align-items:center; }}
  .verdict {{ font-size:20px; font-weight:bold; }}
  .verdict.ok {{ color:#1a9a55; }} .verdict.fail {{ color:#d03050; }}
  .summary .meta {{ color:#777; font-size:14px; }}
  table {{ width:100%; border-collapse:collapse; background:#fff; border-radius:10px; overflow:hidden;
           box-shadow:0 1px 4px rgba(0,0,0,.08); margin-bottom:24px; }}
  th, td {{ padding:10px 16px; text-align:left; border-bottom:1px solid #eee; font-size:14px; }}
  th {{ background:#fafafa; }}
  .badge {{ padding:2px 10px; border-radius:10px; font-size:12px; }}
  .badge.ok {{ background:#e6f7ef; color:#1a9a55; }} .badge.fail {{ background:#fdecec; color:#d03050; }}
  .slide-card {{ background:#fff; border-radius:10px; margin-bottom:20px; overflow:hidden;
                 box-shadow:0 1px 4px rgba(0,0,0,.08); }}
  .slide-head {{ padding:10px 20px; font-size:14px; color:#555; border-bottom:1px solid #f0f0f0; }}
  .slide-body {{ padding:24px; }}
  .slide-body.cover {{ text-align:center; padding:56px 24px; }}
  .cover-title {{ font-size:32px; font-weight:bold; }}
  .slide-body.content {{ display:flex; gap:24px; }}
  .layout-text {{ flex:1; }} .layout-img {{ flex:1; display:flex; align-items:center; justify-content:center; }}
  .layout-text h3 {{ margin:0 0 14px; }}
  .layout-text li {{ margin-bottom:8px; color:#555; }}
  .layout-img img {{ max-width:100%; max-height:340px; border-radius:6px; box-shadow:0 1px 6px rgba(0,0,0,.12); }}
  .img-placeholder {{ width:100%; height:240px; background:#e0e0e0; border-radius:6px;
                      display:flex; align-items:center; justify-content:center; color:#888; }}
  .pptx-path {{ background:#fff; border-radius:10px; padding:16px 20px; font-size:14px;
                box-shadow:0 1px 4px rgba(0,0,0,.08); word-break:break-all; }}
</style></head>
<body><div class="container">
  <h1>PPT 自动生成 · 测试报告</h1>
  <div class="summary">
    <div class="verdict {'ok' if all_ok else 'fail'}">{'✔ 测试全部通过' if all_ok else '✘ 存在失败步骤'}</div>
    <div class="meta">主题：{html.escape(topic)}<br>总耗时 {total_sec:.1f}s · 共 {len(slides)} 页</div>
  </div>
  <table>
    <tr><th>测试步骤</th><th>状态</th><th>耗时</th><th>明细</th></tr>
    {step_rows}
  </table>
  {cards}
  <div class="pptx-path">📄 成品文件：{html.escape(os.path.abspath(out_pptx)) if out_pptx else '未生成'}</div>
</div></body></html>"""

    with open(report_path, "w", encoding="utf-8") as f:
        f.write(page)
    return report_path


def open_report(report_path: str):
    webbrowser.open(f"file:///{os.path.abspath(report_path).replace(chr(92), '/')}")
