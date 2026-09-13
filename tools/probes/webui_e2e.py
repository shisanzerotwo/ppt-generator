"""WebUI 端到端冒烟：真实点击「生成」→ 跟踪阶段 → 等 ready → 截图。

用法：python tools/probes/webui_e2e.py [base_url] [topic] [out_dir]
验证口径：退出码非 0 = JS 报错 / 生成失败 / 超时。
"""
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

import json
import urllib.request

from playwright.sync_api import sync_playwright

import shot as shot_mod

BASE = sys.argv[1] if len(sys.argv) > 1 else 'http://127.0.0.1:5001'
TOPIC = sys.argv[2] if len(sys.argv) > 2 else '量子计算入门：从薛定谔的猫到量子霸权'
OUT = sys.argv[3] if len(sys.argv) > 3 else 'output/webui_e2e'
os.makedirs(OUT, exist_ok=True)


def status():
    with urllib.request.urlopen(BASE + '/api/status', timeout=10) as r:
        return json.load(r)


def main():
    with sync_playwright() as p:
        browser = shot_mod._launch_browser(p)
        page = browser.new_page(viewport={'width': 1600, 'height': 1000})
        errors = []
        page.on('pageerror', lambda e: errors.append(str(e)))
        page.goto(BASE + '/', wait_until='load')
        page.wait_for_timeout(800)

        page.fill('#topic', TOPIC)
        page.click('#btn-gen')
        print('clicked generate:', TOPIC)

        # 跟踪阶段：生成中截一张，然后等 ready（上限 8 分钟）
        mid_shot_done = False
        t0 = time.time()
        last_phase = ''
        while time.time() - t0 < 480:
            s = status()
            ph = s.get('phase')
            if ph != last_phase:
                print(f'[{time.time()-t0:5.1f}s] phase={ph} slides={len(s.get("slides", []))}')
                last_phase = ph
            if not mid_shot_done and ph in ('outline', 'images', 'designing', 'review'):
                page.wait_for_timeout(1200)
                page.screenshot(path=f'{OUT}/e2e_working.png')
                mid_shot_done = True
            if ph == 'ready':
                break
            if ph == 'idle' and t0 + 5 < time.time():
                print('FAILED: back to idle, log tail:',
                      [l['msg'] for l in s.get('log', [])[-3:]])
                return 1
            page.wait_for_timeout(2000)
        else:
            print('TIMEOUT waiting for ready')
            return 1

        s = status()
        page.wait_for_timeout(1500)   # 等 refresh() 渲染卡片与预览
        page.screenshot(path=f'{OUT}/e2e_ready.png', full_page=False)
        page.screenshot(path=f'{OUT}/e2e_ready_full.png', full_page=True)

        # 打开设计稿查看器（新 UI 的 dialog）截一张
        try:
            page.click('text=查看 / 放映', timeout=5000)
            page.wait_for_timeout(1500)
            page.screenshot(path=f'{OUT}/e2e_viewer.png')
            page.keyboard.press('Escape')
        except Exception as e:
            print('viewer skip:', e)

        browser.close()

    print('--- summary ---')
    print('style:', s.get('style_name'), '| slides:', len(s.get('slides', [])),
          '| html:', s.get('html_path'))
    print('log tail:', [l['msg'] for l in s.get('log', [])[-5:]])
    if errors:
        print('PAGEERROR:', errors, file=sys.stderr)
        return 1
    print('E2E OK')
    return 0


if __name__ == '__main__':
    sys.exit(main())
