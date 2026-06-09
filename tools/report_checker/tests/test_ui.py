"""
查重分析 — UI 交互测试
用 Playwright 模拟浏览器，验证：
  1. 页面加载正常
  2. 点击「加载示例数据」后自动分析并展示结果

运行： python -m tools.report_checker.tests.test_ui
"""
import sys, os, subprocess, time, urllib.request

sys.stdout.reconfigure(encoding='utf-8')

PROJECT = os.path.dirname(
    os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
)
sys.path.insert(0, PROJECT)

from playwright.sync_api import sync_playwright

PASS = FAIL = 0


def check(cond, msg):
    global PASS, FAIL
    if cond: PASS += 1; print(f'  ✓ {msg}')
    else:    FAIL += 1; print(f'  ✗ {msg}')


def section(name):
    print(f'\n{"="*60}\n  {name}\n{"="*60}')


PORT = 8594
URL = f'http://localhost:{PORT}'

section('启动 Streamlit 服务')

proc = subprocess.Popen(
    [sys.executable, '-m', 'streamlit', 'run', 'app.py',
     '--server.headless', 'true',
     '--server.port', str(PORT),
     '--server.fileWatcherType', 'none'],
    cwd=PROJECT,
    stdout=subprocess.DEVNULL,
    stderr=subprocess.DEVNULL,
)

started = False
for _ in range(40):
    try:
        urllib.request.urlopen(URL, timeout=2)
        started = True
        break
    except Exception:
        time.sleep(1)

if not started:
    proc.kill()
    print('✗ Streamlit 启动失败')
    sys.exit(1)

print(f'  Streamlit 已启动: {URL}')


try:
    # ════════════════════════════════════════════
    section('Test 1: 页面加载与示例数据分析')
    # ════════════════════════════════════════════

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.set_viewport_size({'width': 1280, 'height': 720})

        page.goto(URL)
        page.wait_for_timeout(1500)

        sidebar = page.locator('section[data-testid="stSidebar"]')
        sidebar.get_by_text('头歌图文实验图片查重').click()
        page.wait_for_timeout(1500)

        heading = page.get_by_text('头歌图文实验图片查重', exact=True).first
        check(heading.is_visible(), '查重分析页面已加载')

        # 点击加载示例数据
        page.get_by_text('加载示例数据看看效果').first.click()
        page.wait_for_timeout(3000)

        # 轮询等待「总报告数」出现，最长 45 秒
        found = False
        for _ in range(45):
            try:
                el = page.get_by_text('总报告数').first
                if el.is_visible():
                    found = True
                    break
            except Exception:
                pass
            time.sleep(1)

        check(found, '示例数据分析后「总报告数」出现')

        browser.close()

finally:
    proc.kill()
    try:
        proc.wait(timeout=5)
    except Exception:
        pass


section('SUMMARY')
print(f'  report_checker UI: {PASS} 通过, {FAIL} 失败')
if FAIL:
    sys.exit(1)
