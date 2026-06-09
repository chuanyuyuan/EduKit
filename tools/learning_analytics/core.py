"""
学情分析 — 核心逻辑
指标统计 + DeepSeek API 调用 + 数据准备
"""
import json

from openai import OpenAI

from .config import (
    DEEPSEEK_API_KEY, DEEPSEEK_API_BASE, DEEPSEEK_MODEL,
    DEEPSEEK_TEMPERATURE, DEEPSEEK_MAX_TOKENS, DEEPSEEK_TIMEOUT,
)
from .prompts import (
    CLASS_ANALYSIS_SYSTEM, CLASS_ANALYSIS_PROMPT,
    STUDENT_ANALYSIS_SYSTEM, STUDENT_ANALYSIS_PROMPT,
)

# 考勤状态常量（与 tools.attendance.core 一致）
_PRESENT = {'扫二维码', '“正在上课”提示', '教师添加', '课堂暗号'}
_LEAVE = {'病假', '事假'}
_ABSENT = '未上课'


# ════════════════════════════════════════════════════════════
# 统计指标
# ════════════════════════════════════════════════════════════

def _attendance_status(raw, has_leave=False):
    """原始签到状态 → 上课/缺勤/请假。"""
    if raw in _PRESENT:
        return "上课"
    if raw == _ABSENT:
        return "请假" if has_leave else "缺勤"
    if raw in _LEAVE:
        return "请假"
    return raw


def _compute_session_score_info(students, session_keys):
    """计算每次课的最高分、平均分和得分人数。"""
    info = {}
    for sk in session_keys:
        scores = []
        for s in students:
            v = s.get("scores", {}).get(sk)
            if v is not None:
                try:
                    scores.append(float(v))
                except (ValueError, TypeError):
                    pass
        info[sk] = {
            "max": max(scores) if scores else 0,
            "avg": sum(scores) / len(scores) if scores else 0,
            "count": len(scores),
        }
    return info


def compute_class_stats(students, session_keys):
    """计算班级整体统计指标。

    Returns:
        dict: {
            "avg_attendance_rate": float,
            "avg_score": float,
            "session_stats": [{session, attendance_rate, avg_score}],
            "attention_list": [(name, reason_str)],
        }
    """
    n = len(students)
    if n == 0 or not session_keys:
        return {
            "avg_attendance_rate": 0,
            "avg_score": 0,
            "session_stats": [],
            "attention_list": [],
        }

    # 每次课统计
    session_stats = []
    total_rate = 0
    total_score = 0
    score_count = 0

    for sk in session_keys:
        attended = 0
        score_sum = 0
        sc = 0
        for s in students:
            att = s.get("attendance", {}).get(sk, "")
            if _attendance_status(att, sk in s.get("leave_sessions", set())) == "上课":
                attended += 1
            score_val = s.get("scores", {}).get(sk)
            if score_val is not None:
                try:
                    score_sum += float(score_val)
                    sc += 1
                except (ValueError, TypeError):
                    pass
        rate = attended / n * 100
        total_rate += rate
        avg = score_sum / sc if sc > 0 else 0
        total_score += avg
        score_count += 1
        session_stats.append({
            "session": sk,
            "attendance_rate": round(rate, 1),
            "avg_score": round(avg, 1),
        })

    # 需关注学生（出勤率低 + 得分低）
    attention_list = []
    for s in students:
        a = sum(1 for sk in session_keys
                if _attendance_status(s.get("attendance", {}).get(sk, ""),
                                      sk in s.get("leave_sessions", set())) == "上课")
        ar = a / len(session_keys) * 100 if session_keys else 0
        sv = [s.get("scores", {}).get(sk) for sk in session_keys]
        sv = [float(v) for v in sv if v is not None]
        avg_s = sum(sv) / len(sv) if sv else 0
        reasons = []
        if ar < 70:
            reasons.append(f"出勤率{ar:.0f}%")
        if avg_s < 60 and sv:
            reasons.append(f"平均分{avg_s:.0f}")
        if reasons:
            attention_list.append((s.get("name", ""), "；".join(reasons)))

    return {
        "avg_attendance_rate": round(total_rate / score_count, 1) if score_count else 0,
        "avg_score": round(total_score / score_count, 1) if score_count else 0,
        "session_stats": session_stats,
        "attention_list": attention_list,
    }


def compute_student_stats(student, session_keys):
    """计算单个学生统计指标。

    Returns:
        dict: {name, student_id, session_count,
               attended_count, absent_count, leave_count,
               attendance_rate, scores, attendance_detail}
    """
    attended = absent = leave = 0
    scores = {}
    detail = {}

    for sk in session_keys:
        raw = student.get("attendance", {}).get(sk, "")
        has_lv = sk in student.get("leave_sessions", set())
        status = _attendance_status(raw, has_lv)
        detail[sk] = status
        if status == "上课":
            attended += 1
        elif status == "请假":
            leave += 1
        elif status == "缺勤":
            absent += 1

        score_val = student.get("scores", {}).get(sk)
        if score_val is not None:
            try:
                scores[sk] = round(float(score_val), 1)
            except (ValueError, TypeError):
                pass

    total = len(session_keys)
    return {
        "name": student.get("name", ""),
        "student_id": student.get("id", ""),
        "session_count": total,
        "attended_count": attended,
        "absent_count": absent,
        "leave_count": leave,
        "attendance_rate": round(attended / total * 100, 1) if total else 0,
        "scores": scores,
        "attendance_detail": detail,
    }


# ════════════════════════════════════════════════════════════
# DeepSeek API
# ════════════════════════════════════════════════════════════

def _get_client(api_key=None):
    return OpenAI(
        api_key=api_key or DEEPSEEK_API_KEY,
        base_url=DEEPSEEK_API_BASE,
        timeout=DEEPSEEK_TIMEOUT,
    )


def call_deepseek(system_prompt, user_prompt, api_key=None):
    """调用 DeepSeek API，返回解析后的 dict。"""
    key = api_key or DEEPSEEK_API_KEY
    if not key:
        return {"error": "未配置 DEEPSEEK_API_KEY"}

    try:
        client = _get_client(api_key)
        resp = client.chat.completions.create(
            model=DEEPSEEK_MODEL,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            temperature=DEEPSEEK_TEMPERATURE,
            max_tokens=DEEPSEEK_MAX_TOKENS,
        )
        text = resp.choices[0].message.content
        # 清理可能的 markdown 代码块标记
        text = text.strip()
        if text.startswith("```"):
            text = text.split("\n", 1)[-1]
            text = text.rsplit("```", 1)[0]
        return json.loads(text.strip())
    except Exception as e:
        return {"error": f"AI 分析调用失败：{e}"}


# ════════════════════════════════════════════════════════════
# 数据准备（拼接 prompt）
# ════════════════════════════════════════════════════════════

def prepare_class_data(class_name, students, session_keys):
    """拼装班级分析的 user prompt。"""
    session_info = _compute_session_score_info(students, session_keys)

    # 课次信息（含得分范围和百分比）
    sessions_data_lines = []
    for i, sk in enumerate(session_keys):
        si = session_info[sk]
        has_quiz = si["max"] > 0
        if has_quiz:
            pct = si["avg"] / si["max"] * 100 if si["max"] > 0 else 0
            sessions_data_lines.append(
                f"  第{i + 1}次课({sk}): 全班最高{si['max']:.0f}分, "
                f"班级平均{si['avg']:.1f}分（{pct:.0f}%）"
            )
        else:
            sessions_data_lines.append(
                f"  第{i + 1}次课({sk}): 未出题（全班最高0分, 平均0分）"
            )
    sessions_data = "\n".join(sessions_data_lines)

    # 学生概要（得分附百分比）
    lines = []
    for s in students:
        name = s.get("name", "")
        score_parts = []
        for sk in session_keys:
            v = s.get("scores", {}).get(sk, "-")
            si = session_info[sk]
            if si["max"] > 0 and v != "-":
                try:
                    pct = float(v) / si["max"] * 100
                    score_parts.append(f"{sk}:{v}/{si['max']:.0f}分（{pct:.0f}%）")
                except (ValueError, TypeError):
                    score_parts.append(f"{sk}:{v}分")
            else:
                score_parts.append(f"{sk}:未出题" if not si["max"] else f"{sk}:{v}分")
        scores = " ".join(score_parts)
        atts = " ".join(
            f"{sk}:{_attendance_status(s.get('attendance', {}).get(sk, ''), sk in s.get('leave_sessions', set()))}"
            for sk in session_keys
        )
        lines.append(f"  {name} | 得分: {scores} | 出勤: {atts}")

    return CLASS_ANALYSIS_PROMPT.format(
        class_name=class_name,
        student_count=len(students),
        session_count=len(session_keys),
        sessions_data=sessions_data,
        students_data="\n".join(lines),
    )


def prepare_student_data(student, session_keys, session_score_info=None):
    """拼装个人画像的 user prompt。"""
    stats = compute_student_stats(student, session_keys)

    scores_lines_lines = []
    for i, sk in enumerate(session_keys):
        score_val = stats['scores'].get(sk)
        score_str = f"{score_val}分" if score_val is not None else "未出题"
        if session_score_info and sk in session_score_info:
            si = session_score_info[sk]
            if si["max"] > 0 and score_val is not None:
                pct = score_val / si["max"] * 100
                score_str += f"（全班最高{si['max']:.0f}分, 平均{si['avg']:.1f}分, 得分率{pct:.0f}%）"
            else:
                score_str += f"（全班最高{si['max']:.0f}分, 平均{si['avg']:.1f}分）"
        scores_lines_lines.append(f"  第{i + 1}次课({sk}): {score_str}")
    scores_lines = "\n".join(scores_lines_lines)

    detail_lines = "\n".join(
        f"  第{i + 1}次课({sk}): {stats['attendance_detail'].get(sk, '')}"
        for i, sk in enumerate(session_keys)
    )

    return STUDENT_ANALYSIS_PROMPT.format(
        name=stats["name"],
        student_id=stats["student_id"],
        session_count=stats["session_count"],
        attended_count=stats["attended_count"],
        absent_count=stats["absent_count"],
        leave_count=stats["leave_count"],
        attendance_rate=stats["attendance_rate"],
        scores=scores_lines,
        attendance_detail=detail_lines,
    )


# ════════════════════════════════════════════════════════════
# 分析入口（统计 + AI 一站式）
# ════════════════════════════════════════════════════════════

def analyze_class(class_name, students, session_keys, api_key=None):
    """班级学情分析：统计指标 + AI 评语。

    Returns:
        dict: {stats: {...}, ai: {...}}
    """
    stats = compute_class_stats(students, session_keys)
    prompt = prepare_class_data(class_name, students, session_keys)
    ai_result = call_deepseek(CLASS_ANALYSIS_SYSTEM, prompt, api_key=api_key)
    return {"stats": stats, "ai": ai_result}


def analyze_student(student, session_keys, session_score_info=None, api_key=None):
    """个人学情画像：统计指标 + AI 评语。

    Args:
        session_score_info: 由 _compute_session_score_info 生成的每次课得分上下文，可选
        api_key: DeepSeek API key，可选，默认使用 config.DEEPSEEK_API_KEY

    Returns:
        dict: {stats: {...}, ai: {...}}
    """
    stats = compute_student_stats(student, session_keys)
    prompt = prepare_student_data(student, session_keys, session_score_info)
    ai_result = call_deepseek(STUDENT_ANALYSIS_SYSTEM, prompt, api_key=api_key)
    return {"stats": stats, "ai": ai_result}
