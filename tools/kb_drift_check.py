"""知识库漂移核对（零第三方依赖，供 Windows 任务计划程序每周自动跑）。

机械核对：路由表 / 模块清单 / 测试条数 / iframe 安全表述 / 关键文件存在。
有漂移 → 写报告到 KB 目录，并尝试弹 Windows 通知（失败静默）。
用法：python kb_drift_check.py
"""
import os
import re
import subprocess
import sys
from datetime import datetime

REPO = r"D:\GitHub\xiangmu\ppt-generator"
KB = r"D:\ObsidianChanku\ai学习\research\AI Agent学习\PPT制作台"
PY = os.path.join(REPO, ".venv", "Scripts", "python.exe")
REPORT = os.path.join(KB, "_drift_report.md")

KB_FILES = ["README.md", "DECISIONS.md", "KNOWLEDGE.md"]


def _read(p):
    try:
        with open(p, encoding="utf-8") as f:
            return f.read()
    except OSError:
        return ""


def _norm(s):
    """动态段归一：/api/slide/<int:i>/text 与 /api/slide/<i>/text 都 → /api/slide/X/text。"""
    return re.sub(r"<[^>]+>", "X", s)


def check_routes(doc):
    app = _read(os.path.join(REPO, "app.py"))
    routes = set(re.findall(r'@app\.route\("([^"]+)"', app))
    ndoc = _norm(doc)
    issues = []
    for r in sorted(routes):
        if _norm(r) not in ndoc:
            issues.append(f"路由 `{r}` 未出现在 KNOWLEDGE.md 路由清单")
    return issues


def check_modules(doc):
    issues = []
    for fn in os.listdir(REPO):
        if fn.endswith(".py") and fn not in doc:
            issues.append(f"模块 `{fn}` 未在文档模块表登记")
    return issues


def check_test_count(doc):
    if not os.path.isfile(PY):
        return ["找不到 venv python，跳过测试数核对"]
    try:
        out = subprocess.run([PY, "-m", "pytest", "tests/", "-q"],
                             cwd=REPO, capture_output=True, text=True, timeout=600).stdout
    except Exception as e:
        return [f"跑测试失败：{e}"]
    m = re.search(r"(\d+) passed", out)
    if not m:
        return ["未能从 pytest 输出解析通过数"]
    actual = int(m.group(1))
    doc_nums = [int(x) for x in re.findall(r"(\d+)\s*条", doc)]
    if doc_nums and actual not in doc_nums:
        return [f"测试数不一致：实际 {actual} 条，文档写 {doc_nums}"]
    return []


def check_iframe_security(doc):
    """防复现历史错误：文档绝不能建议给 allow-scripts 再加 allow-same-origin。"""
    bad = re.search(r"(移除限制|加回|显式)\s*allow-same-origin", doc)
    if bad:
        return ["KN""OWLEDGE.md 出现危险建议：给 allow-scripts 追加 allow-same-origin（会破坏同源防护）"]
    if "allow-same-origin" not in doc:
        return ["KN""OWLEDGE.md 缺少 iframe sandbox 安全说明（应写明禁 allow-same-origin 是有意加固）"]
    return []


def check_links():
    return [f"知识库文件缺失：{f}" for f in KB_FILES if not os.path.isfile(os.path.join(KB, f))]


def notify(title, msg):
    ps = (
        "try{"
        "[Windows.UI.Notifications.ToastNotificationManager,Windows.UI.Notifications,ContentType=WindowsRuntime]|Out-Null;"
        "$t=[Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent("
        "[Windows.UI.Notifications.ToastTemplateType]::ToastText02);"
        "$x=$t.GetElementsByTagName('text');"
        f"$x.Item(0).AppendChild($t.CreateTextNode('{title}'))|Out-Null;"
        f"$x.Item(1).AppendChild($t.CreateTextNode('{msg}'))|Out-Null;"
        "[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier('Claude KB').Show("
        "[Windows.UI.Notifications.ToastNotification]::new($t))"
        "}catch{}"
    )
    try:
        subprocess.run(["powershell", "-NoProfile", "-Command", ps],
                       capture_output=True, timeout=30)
    except Exception:
        pass


def main():
    doc = "\n".join(_read(os.path.join(KB, f)) for f in KB_FILES)
    issues = (check_links() + check_modules(doc) + check_routes(doc)
              + check_test_count(doc) + check_iframe_security(doc))
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M")
    if issues:
        body = f"# 知识库漂移报告 {stamp}\n\n发现 {len(issues)} 项：\n" + \
               "\n".join(f"- {i}" for i in issues)
        with open(REPORT, "w", encoding="utf-8") as f:
            f.write(body)
        print(body)
        notify("知识库需要更新", f"发现 {len(issues)} 项漂移，详见 _drift_report.md")
        return 1
    if os.path.isfile(REPORT):
        os.remove(REPORT)  # 漂移已修好，清掉旧报告
    print(f"✓ 无漂移（{stamp}）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
