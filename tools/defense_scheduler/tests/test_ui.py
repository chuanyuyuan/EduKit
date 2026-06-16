"""
答辩顺序生成器 — UI 交互测试
验证 SortableJS 拖拽组件渲染、随机生成流程、Excel 下载。

运行: python -m tools.defense_scheduler.tests.test_ui
"""
import sys, os, subprocess, time, urllib.request

sys.stdout.reconfigure(encoding='utf-8')

PROJECT = os.path.dirname(
    os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
)
sys.path.insert(0, PROJECT)

from playwright.sync_api import sync_playwright, expect

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

print('  ✓ Streamlit 已启动')


def nav_to_defense(page):
    """通过侧边栏导航到答辩顺序生成器"""
    page.locator('section[data-testid="stSidebar"]').get_by_text('答辩顺序生成器').click()
    page.wait_for_load_state('networkidle')
    page.wait_for_timeout(1500)


try:
    # ════════════════════════════════════════════
    section('Test 1: 页面渲染')
    # ════════════════════════════════════════════

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.set_viewport_size({'width': 1280, 'height': 720})

        page.goto(URL)
        page.wait_for_load_state('networkidle')
        nav_to_defense(page)

        # 标题
        title = page.locator('h1')
        expect(title).to_have_text('🎤 答辩顺序生成器')
        check(True, '标题显示正确')

        # 姓名输入框
        names_input = page.locator('textarea')
        expect(names_input).to_be_visible()
        check(True, '姓名输入框可见')

        # 开始/结束时间输入
        check(page.locator('input[value="08:00"]').is_visible(), '开始时间 08:00')
        check(page.locator('input[value="09:30"]').is_visible(), '结束时间 09:30')

        # 随机生成按钮
        gen_btn = page.locator('button:has-text("随机生成")')
        expect(gen_btn).to_be_visible()
        check(True, '随机生成按钮可见')

        browser.close()

    # ════════════════════════════════════════════
    section('Test 2: 随机生成 + 数据展示')
    # ════════════════════════════════════════════

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.set_viewport_size({'width': 1280, 'height': 720})

        page.goto(URL)
        page.wait_for_load_state('networkidle')
        nav_to_defense(page)

        # 填入自定义姓名
        textareas = page.locator('textarea')
        textareas.fill('甲\n乙\n丙\n丁\n戊')
        page.wait_for_load_state('networkidle')

        # 点击随机生成
        page.get_by_role('button', name='随机生成').click()
        page.wait_for_load_state('networkidle')
        page.wait_for_timeout(3000)  # 等待组件加载

        # 检查拖拽提示
        check(page.locator('text=拖拽调整顺序').is_visible(), '"拖拽调整顺序"提示可见')

        # 检查成功消息（含时长信息）
        success = page.locator('.stAlert')
        expect(success).to_be_visible()
        check(True, '成功提示可见')

        # 检查 dataframe 显示
        df = page.locator('[data-testid="stDataFrame"]')
        expect(df).to_be_visible()
        check(True, '时间表 dataframe 可见')

        # 检查 5 个学生
        rows = page.locator('[data-testid="stDataFrame"] tbody tr')
        check(rows.count() == 5, f'时间表显示 5 名学生（实际 {rows.count()}）')

        # 检查每行有姓名
        cells = page.locator('[data-testid="stDataFrame"] tbody td').all_text_contents()
        cell_text = ' '.join(cells)
        check('甲' in cell_text, '表格含学生甲')
        check('乙' in cell_text, '表格含学生乙')

        # 检查时间格式
        check('08:00' in cell_text, '表格含开始时间 08:00')

        browser.close()

    # ════════════════════════════════════════════
    section('Test 3: 重新随机 → 列表刷新')
    # ════════════════════════════════════════════

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.set_viewport_size({'width': 1280, 'height': 720})

        page.goto(URL)
        page.wait_for_load_state('networkidle')
        nav_to_defense(page)

        textareas = page.locator('textarea')
        textareas.fill('甲\n乙\n丙\n丁\n戊')
        page.wait_for_load_state('networkidle')

        gen_btn = page.get_by_role('button', name='随机生成')
        gen_btn.click()
        page.wait_for_timeout(3000)

        check(page.locator('[data-testid="stDataFrame"]').is_visible(), '首次生成成功')

        # 再次点击随机生成
        gen_btn.click()
        page.wait_for_timeout(3000)

        check(page.locator('[data-testid="stDataFrame"]').is_visible(), '重新随机后表格仍在')
        check(page.locator('.stAlert').is_visible(), '成功提示仍然可见')

        rows2 = page.locator('[data-testid="stDataFrame"] tbody tr')
        check(rows2.count() == 5, '重新随机后仍有 5 名学生')

        browser.close()

    # ════════════════════════════════════════════
    section('Test 4: Excel 下载按钮')
    # ════════════════════════════════════════════

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.set_viewport_size({'width': 1280, 'height': 720})

        page.goto(URL)
        page.wait_for_load_state('networkidle')
        nav_to_defense(page)

        textareas = page.locator('textarea')
        textareas.fill('甲\n乙\n丙')
        page.get_by_role('button', name='随机生成').click()
        page.wait_for_timeout(3000)

        dl_btn = page.locator('button:has-text("下载 Excel")')
        expect(dl_btn).to_be_visible()
        check(dl_btn.is_enabled(), '下载按钮可见且可点击')

        browser.close()

    # ════════════════════════════════════════════
    section('Test 5: 空输入校验')
    # ════════════════════════════════════════════

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.set_viewport_size({'width': 1280, 'height': 720})

        page.goto(URL)
        page.wait_for_load_state('networkidle')
        nav_to_defense(page)

        textareas = page.locator('textarea')
        textareas.fill('')
        page.get_by_role('button', name='随机生成').click()
        page.wait_for_timeout(1000)

        check(page.locator('text=请至少输入一个姓名').is_visible(), '空输入显示错误提示')

        browser.close()

finally:
    proc.kill()
    try:
        proc.wait(timeout=5)
    except Exception:
        pass

section('汇总')
print(f'  defense_scheduler UI: {PASS} 通过, {FAIL} 失败')
if FAIL:
    sys.exit(1)
else:
    print('  ✅ 全部通过')
