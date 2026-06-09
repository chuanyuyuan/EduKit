"""
名单比对 — UI 交互测试
用 Playwright 模拟浏览器，验证：
  1. 点击文本域和按钮后页面不跳转
  2. 点击「开始比对」后正确展示结果
  3. 清除按钮工作正常

运行： python -m tools.roster_diff.tests.test_ui
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


PORT = 8592
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
    section('Test 1: 点击文本域和按钮后页面不跳转')
    # ════════════════════════════════════════════

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.set_viewport_size({'width': 1280, 'height': 720})

        page.goto(URL)
        page.wait_for_timeout(1500)

        # 通过侧边栏切换到 setDiff 工具
        sidebar = page.locator('section[data-testid="stSidebar"]')
        sidebar.get_by_text('setDiff工具').click()
        page.wait_for_timeout(1500)

        tool_heading = page.get_by_text('setDiff工具', exact=True).first
        check(tool_heading.is_visible(), 'setDiff 工具已加载')

        # 在文本域 A 输入
        textareas = page.locator('textarea')
        textareas.nth(0).click()
        textareas.nth(0).fill('张三\n李四\n王五')
        page.wait_for_timeout(500)

        check(tool_heading.is_visible(), '输入 A 后页面未跳转')

        # 点击文本域 B（不输入）
        textareas.nth(1).click()
        page.wait_for_timeout(500)

        check(tool_heading.is_visible(), '点击文本域 B 后页面未跳转')

        # 在文本域 B 输入
        textareas.nth(1).fill('张三\n王五\n赵六')
        page.wait_for_timeout(500)

        check(tool_heading.is_visible(), '输入 B 后页面未跳转')
        check(page.get_by_role('button', name='开始比对').is_visible(),
              '开始比对按钮可见')

        # 点击开始比对
        page.get_by_role('button', name='开始比对').click()
        page.wait_for_timeout(1500)

        check(tool_heading.is_visible(), '点击比对后页面未跳转')

        # 确认有统计指标
        metric = page.locator('[data-testid="stMetricLabel"]').first
        check(metric.is_visible(), '统计指标出现')

        # 再次点击文本域 B
        textareas.nth(1).click()
        page.wait_for_timeout(500)

        check(tool_heading.is_visible(), '再次点击文本域 B 后页面未跳转')
        check(metric.is_visible(), '结果仍然可见')

        browser.close()

    # ════════════════════════════════════════════
    section('Test 2: 比对结果正确')
    # ════════════════════════════════════════════

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.set_viewport_size({'width': 1280, 'height': 720})

        page.goto(URL)
        page.wait_for_timeout(1500)
        page.locator('section[data-testid="stSidebar"]').get_by_text('setDiff工具').click()
        page.wait_for_timeout(1500)

        textareas = page.locator('textarea')
        textareas.nth(0).fill('04230001\n04230002\n04230003\n04230004')
        textareas.nth(1).fill('04230001\n04230003\n04230005')
        page.get_by_role('button', name='开始比对').click()
        page.wait_for_timeout(1500)

        # 统计指标可见
        metric = page.locator('[data-testid="stMetricLabel"]').first
        check(metric.is_visible(), '比对后统计指标可见')

        # 点击"仅列表 A" tab
        page.get_by_text('仅列表 A').first.click()
        page.wait_for_timeout(500)

        # 04230002 和 04230004 应在仅 A 中
        code_blocks = page.locator('code')
        code_text = code_blocks.all_text_contents()
        full_text = '\n'.join(code_text)
        check('04230002' in full_text, '仅 A 包含 04230002')
        check('04230004' in full_text, '仅 A 包含 04230004')

        browser.close()

    # ════════════════════════════════════════════
    section('Test 3: 清除按钮')
    # ════════════════════════════════════════════

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.set_viewport_size({'width': 1280, 'height': 720})

        page.goto(URL)
        page.wait_for_timeout(1500)
        page.locator('section[data-testid="stSidebar"]').get_by_text('setDiff工具').click()
        page.wait_for_timeout(1500)

        textareas = page.locator('textarea')
        textareas.nth(0).fill('04230001\n04230002')
        textareas.nth(1).fill('04230001\n04230003')
        page.get_by_role('button', name='开始比对').click()
        page.wait_for_timeout(1500)

        # 确认结果出现
        metric = page.locator('[data-testid="stMetricLabel"]').first
        check(metric.is_visible(), '结果已显示')

        # 清除按钮可见
        check(page.get_by_text('清除').is_visible(), '清除按钮可见')

        # 点击清除
        page.get_by_text('清除').click()
        page.wait_for_timeout(1500)

        # 确认 textarea 被清空
        textareas = page.locator('textarea')
        val_a = textareas.nth(0).input_value()
        val_b = textareas.nth(1).input_value()
        check(val_a == '' and val_b == '', '文本域已被清空')

        browser.close()

finally:
    proc.kill()
    try:
        proc.wait(timeout=5)
    except Exception:
        pass


section('SUMMARY')
print(f'  roster_diff UI: {PASS} 通过, {FAIL} 失败')
if FAIL:
    sys.exit(1)
