"""
考勤分析 — UI 交互测试
用 Playwright 模拟浏览器，验证：
  1. 页面加载和工具切换工作正常
  2. 点击「加载示例数据」后自动解析并展示结果
  3. 单文件 / 合并模式切换正常

运行： python -m tools.attendance.tests.test_ui
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


PORT = 8593
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
    section('Test 1: 页面加载与工具切换')
    # ════════════════════════════════════════════

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.set_viewport_size({'width': 1280, 'height': 720})

        page.goto(URL)
        page.wait_for_timeout(1500)

        sidebar = page.locator('section[data-testid="stSidebar"]')
        sidebar.get_by_text('雨课堂课堂数据分析').click()
        page.wait_for_timeout(1500)

        heading = page.get_by_text('雨课堂课堂数据分析', exact=True).first
        check(heading.is_visible(), '考勤分析页面已加载')

        sidebar.get_by_text('setDiff工具').click()
        page.wait_for_timeout(1000)
        sidebar.get_by_text('雨课堂课堂数据分析').click()
        page.wait_for_timeout(1000)
        check(heading.is_visible(), '切换后回到考勤分析页面')

        browser.close()

    # ════════════════════════════════════════════
    section('Test 2: 加载示例数据 → 自动解析展示结果')
    # ════════════════════════════════════════════

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.set_viewport_size({'width': 1280, 'height': 720})

        page.goto(URL)
        page.wait_for_timeout(1500)
        sidebar = page.locator('section[data-testid="stSidebar"]')
        sidebar.get_by_text('雨课堂课堂数据分析').click()
        page.wait_for_timeout(1500)

        # 点击加载示例数据
        page.get_by_text('加载示例数据看看效果').first.click()
        page.wait_for_timeout(2000)

        # 轮询等待「考勤概况」出现，最长 40 秒
        found = False
        for _ in range(40):
            try:
                el = page.get_by_text('考勤概况').first
                if el.is_visible():
                    found = True
                    break
            except Exception:
                pass
            time.sleep(1)

        check(found, '示例数据分析后「考勤概况」出现')

        browser.close()

    # ════════════════════════════════════════════
    section('Test 3: 合并模式切换')
    # ════════════════════════════════════════════

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.set_viewport_size({'width': 1280, 'height': 720})

        page.goto(URL)
        page.wait_for_timeout(1500)
        sidebar = page.locator('section[data-testid="stSidebar"]')
        sidebar.get_by_text('雨课堂课堂数据分析').click()
        page.wait_for_timeout(1500)

        # 找到第二个 segmented_control 按钮并点击
        seg_btns = page.locator('button[kind="segmented_control"], button[kind="segmented_controlActive"]')
        if seg_btns.count() >= 2:
            seg_btns.nth(1).click()
            page.wait_for_timeout(3000)

        # 检查合并模式的标志性文本
        el = page.get_by_text('选择文件一')
        check(el.count() == 1, '合并模式显示「选择文件一」')

        browser.close()

finally:
    proc.kill()
    try:
        proc.wait(timeout=5)
    except Exception:
        pass


section('SUMMARY')
print(f'  attendance UI: {PASS} 通过, {FAIL} 失败')
if FAIL:
    sys.exit(1)
