"""
名单比对 — 核心逻辑测试
"""
import sys, os
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..', '..'))

from tools.roster_diff.core import diff_sets, _parse_lines

PASS = FAIL = 0


def check(cond, msg):
    global PASS, FAIL
    if cond: PASS += 1; print(f'  ✓ {msg}')
    else: FAIL += 1; print(f'  ✗ {msg}')


def check_eq(a, b, msg):
    global PASS, FAIL
    if a == b: PASS += 1; print(f'  ✓ {msg} ({a})')
    else: FAIL += 1; print(f'  ✗ {msg}: got {a}, expected {b}')


def section(name):
    print(f'\n{"="*60}\n  {name}\n{"="*60}')


# ════════════════════════════════════════════
# Test 1: _parse_lines 基础功能
# ════════════════════════════════════════════
section('Test 1: _parse_lines 基础功能')

check_eq(_parse_lines(''), [], '空字符串返回空列表')
check_eq(_parse_lines('  \n  \n'), [], '纯空白行被忽略')
check_eq(_parse_lines('abc'), ['abc'], '单行')
check_eq(_parse_lines('abc\ndef'), ['abc', 'def'], '两行')
check_eq(_parse_lines('  abc  \n  def  '), ['abc', 'def'], 'trim 前后空白')
check_eq(_parse_lines('abc\nabc'), ['abc'], '重复行去重只保留一个')


# ════════════════════════════════════════════
# Test 2: diff_sets 空值
# ════════════════════════════════════════════
section('Test 2: diff_sets 空值')

r = diff_sets('', '')
check_eq(r['stats']['a_count'], 0, '两边为空 → A 计数 0')
check_eq(r['stats']['b_count'], 0, '两边为空 → B 计数 0')
check_eq(len(r['only_a']), 0, '两边为空 → only_a 空')
check_eq(len(r['only_b']), 0, '两边为空 → only_b 空')
check_eq(len(r['both']), 0, '两边为空 → both 空')

r = diff_sets('张三\n李四', '')
check_eq(r['stats']['a_count'], 2, '仅 A 有内容')
check_eq(r['stats']['b_count'], 0, '仅 A 有内容 → B 计数 0')
check_eq(r['stats']['only_a_count'], 2, '仅 A 有内容 → only_a 为 2')
check_eq(r['stats']['only_b_count'], 0, '仅 A 有内容 → only_b 为 0')
check_eq(r['stats']['both_count'], 0, '仅 A 有内容 → both 为 0')
check_eq(r['only_a'], ['张三', '李四'], '仅 A 有内容 → only_a 列表正确')

r = diff_sets('', '张三\n李四')
check_eq(r['stats']['only_b_count'], 2, '仅 B 有内容 → only_b 为 2')
check_eq(r['only_b'], ['张三', '李四'], '仅 B 有内容 → only_b 列表正确')


# ════════════════════════════════════════════
# Test 3: diff_sets 基本交集与差集
# ════════════════════════════════════════════
section('Test 3: diff_sets 基本交集与差集')

r = diff_sets('张三\n李四\n王五', '张三\n王五\n赵六')
check_eq(r['stats']['a_count'], 3, 'A 去重后 3 条')
check_eq(r['stats']['b_count'], 3, 'B 去重后 3 条')
check_eq(r['stats']['both_count'], 2, '交集 2 条（张三、王五）')
check_eq(r['stats']['only_a_count'], 1, '仅 A 1 条（李四）')
check_eq(r['stats']['only_b_count'], 1, '仅 B 1 条（赵六）')
check_eq(r['only_a'], ['李四'], '仅 A 为李四')
check_eq(r['only_b'], ['赵六'], '仅 B 为赵六')
check_eq(set(r['both']), {'张三', '王五'}, '交集包含张三、王五')


# ════════════════════════════════════════════
# Test 4: diff_sets 完全一致
# ════════════════════════════════════════════
section('Test 4: diff_sets 完全一致')

r = diff_sets('张三\n李四\n王五', '张三\n李四\n王五')
check_eq(r['stats']['both_count'], 3, '完全一致 → both=3')
check_eq(r['stats']['only_a_count'], 0, '完全一致 → only_a=0')
check_eq(r['stats']['only_b_count'], 0, '完全一致 → only_b=0')


# ════════════════════════════════════════════
# Test 5: diff_sets 无交集
# ════════════════════════════════════════════
section('Test 5: diff_sets 无交集')

r = diff_sets('张三\n李四', '王五\n赵六')
check_eq(r['stats']['both_count'], 0, '无交集 → both=0')
check_eq(r['stats']['only_a_count'], 2, '无交集 → only_a=2')
check_eq(r['stats']['only_b_count'], 2, '无交集 → only_b=2')


# ════════════════════════════════════════════
# Test 6: diff_sets 去重
# ════════════════════════════════════════════
section('Test 6: diff_sets 去重')

r = diff_sets('张三\n张三\n李四', '李四\n李四\n王五')
check_eq(r['stats']['a_count'], 2, 'A 重复行去重后 2 条')
check_eq(r['stats']['b_count'], 2, 'B 重复行去重后 2 条')
check_eq(r['stats']['both_count'], 1, '去重后交集 1 条（李四）')
check_eq(r['stats']['only_a_count'], 1, '去重后仅 A 1 条（张三）')
check_eq(r['stats']['only_b_count'], 1, '去重后仅 B 1 条（王五）')


# ════════════════════════════════════════════
# Test 7: diff_sets trim 空白
# ════════════════════════════════════════════
section('Test 7: diff_sets trim 空白')

r = diff_sets('  张三  \n 李四 ', '张三\n李四')
check_eq(r['stats']['both_count'], 2, 'trim 后交集匹配')
check_eq(r['stats']['only_a_count'], 0, 'trim 后无差集')


# ════════════════════════════════════════════
# Test 8: 学生场景 — 查未交作业
# ════════════════════════════════════════════
section('Test 8: 学生场景 — 查未交作业')

roster = '04230001\n04230002\n04230003\n04230004'
submitted = '04230001\n04230003\n04230005'
r = diff_sets(roster, submitted)
check_eq(r['stats']['only_a_count'], 2, '未交作业 2 人')
check_eq(r['only_a'], ['04230002', '04230004'], '未交学号正确')
check_eq(r['stats']['only_b_count'], 1, '多余提交 1 人（04230005）')
check_eq(r['only_b'], ['04230005'], '多余学号正确')
check_eq(r['stats']['both_count'], 2, '已交作业 2 人')


# ════════════════════════════════════════════
# Test 9: 学生场景 — 查缺勤
# ════════════════════════════════════════════
section('Test 9: 学生场景 — 查缺勤')

roster = '张三\n李四\n王五\n赵六'
signed = '张三\n王五\n赵六'
r = diff_sets(roster, signed)
check_eq(r['stats']['only_a_count'], 1, '缺勤 1 人（李四）')
check_eq(r['only_a'], ['李四'], '缺勤学生正确')


# ════════════════════════════════════════════
# Test 10: case_sensitive=False 忽略大小写
# ════════════════════════════════════════════
section('Test 10: case_sensitive=False 忽略大小写')

r = diff_sets('张三\n李四', '张三\n李四', case_sensitive=False)
check_eq(r['stats']['both_count'], 2, '完全相同 → both=2')

r = diff_sets('Zhang San\nLi Si', 'zhang san\nli si', case_sensitive=False)
check_eq(r['stats']['both_count'], 2, '大小写不同但匹配 → both=2')
check_eq(r['stats']['only_a_count'], 0, '大小写不同 → only_a=0')

r = diff_sets('张三\n李四', '张三\n王五', case_sensitive=False)
check_eq(r['stats']['both_count'], 1, '普通交集 still=1')

# 验证原始大小写保留
r = diff_sets('Zhang San', 'zhang san', case_sensitive=False)
check_eq(r['both'], ['Zhang San'], '交集保留 A 的原始大小写')

r = diff_sets('zhang san', 'Zhang San\nLi Si', case_sensitive=False)
check_eq(r['only_b'], ['Li Si'], '差集保留 B 的原始大小写')

# 默认 case_sensitive=True
r = diff_sets('Zhang San', 'zhang san')
check_eq(r['stats']['both_count'], 0, '默认大小写敏感 → 不匹配')


# ════════════════════════════════════════════
# Test 11: 重复行 + 空白行组合
# ════════════════════════════════════════════
section('Test 11: 重复行 + 空白行组合')

r = diff_sets('\n\n张三\n\n李四\n\n', '\n张三\n\n王五\n')
check_eq(r['stats']['a_count'], 2, 'A 空白行被忽略')
check_eq(r['stats']['b_count'], 2, 'B 空白行被忽略')
check_eq(r['stats']['both_count'], 1, '交集 1 条（张三）')
check_eq(r['only_a'], ['李四'], '仅 A 为李四')
check_eq(r['only_b'], ['王五'], '仅 B 为王五')


# ════════════════════════════════════════════
section(f'SUMMARY\n  roster_diff: {PASS} 通过, {FAIL} 失败')
if FAIL:
    sys.exit(1)
