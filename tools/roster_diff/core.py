"""
名单比对 — 集合运算核心逻辑
参照 setdiff.com：两个文本列表求交集和差集。
"""


def _parse_lines(text: str) -> list[str]:
    """Split text into lines, trim whitespace, remove empties."""
    seen = set()
    result = []
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        if line not in seen:
            seen.add(line)
            result.append(line)
    return result


def diff_sets(text_a: str, text_b: str, *, case_sensitive: bool = True):
    """
    比较两份名单，返回 { only_a, only_b, both, stats }。

    参数：
        text_a: 列表 A 的原始文本（每行一条）
        text_b: 列表 B 的原始文本（每行一条）
        case_sensitive: 是否大小写敏感

    返回：
        {
            'only_a': [...],   # 仅 A 有的条目
            'only_b': [...],   # 仅 B 有的条目
            'both': [...],     # 交集
            'stats': {
                'a_count': int,     # A 去重后总数
                'b_count': int,     # B 去重后总数
                'both_count': int,  # 交集大小
                'only_a_count': int,
                'only_b_count': int,
            }
        }
    """
    lines_a = _parse_lines(text_a)
    lines_b = _parse_lines(text_b)

    if case_sensitive:
        set_a = set(lines_a)
        set_b = set(lines_b)
    else:
        set_a = {l.lower() for l in lines_a}
        set_b = {l.lower() for l in lines_b}

    both_set = set_a & set_b
    only_a_set = set_a - set_b
    only_b_set = set_b - set_a

    # 按原始顺序返回（仅对差集保留原始大小写形式）
    if case_sensitive:
        only_a = [l for l in lines_a if l in only_a_set]
        only_b = [l for l in lines_b if l in only_b_set]
        both = [l for l in lines_a if l in both_set]
    else:
        a_lower_map = {l.lower(): l for l in lines_a}
        b_lower_map = {l.lower(): l for l in lines_b}
        only_a = [a_lower_map[l] for l in sorted(only_a_set)]
        only_b = [b_lower_map[l] for l in sorted(only_b_set)]
        both = [a_lower_map[l] for l in sorted(both_set)]

    return {
        'only_a': only_a,
        'only_b': only_b,
        'both': both,
        'stats': {
            'a_count': len(set_a),
            'b_count': len(set_b),
            'both_count': len(both_set),
            'only_a_count': len(only_a_set),
            'only_b_count': len(only_b_set),
        },
    }
