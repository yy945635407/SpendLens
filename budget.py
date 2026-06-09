"""预算管理、规则引擎、财务评分、What-if 模拟 — 轻量模块。"""
import os
import json
import math
import uuid
import re
from datetime import datetime

# ---- 文件路径 ----
BASE = os.path.dirname(os.path.abspath(__file__))
BUDGET_FILE = os.path.join(BASE, 'budget_config.json')
RULES_FILE = os.path.join(BASE, 'auto_rules.json')
SHARES_DIR = os.path.join(BASE, 'shares')
DATA_DIR = os.path.join(BASE, 'data')

os.makedirs(SHARES_DIR, exist_ok=True)
os.makedirs(DATA_DIR, exist_ok=True)


# ================================================================
#  1. 财务健康评分
# ================================================================

def health_score(data):
    """基于多维度打分，返回 {score, grade, color, suggestions, details}"""
    details = []
    score = 0

    # 维度1：储蓄率（满分35）
    sr = data.get('savings_rate', 0)
    if sr >= 60:
        score += 35
        details.append({'dim': '储蓄率', 'score': 35, 'max': 35, 'comment': f'储蓄率 {sr}%，非常优秀 ✨'})
    elif sr >= 40:
        score += 25
        details.append({'dim': '储蓄率', 'score': 25, 'max': 35, 'comment': f'储蓄率 {sr}%，良好，还有提升空间'})
    elif sr >= 20:
        score += 15
        details.append({'dim': '储蓄率', 'score': 15, 'max': 35, 'comment': f'储蓄率 {sr}%，偏低，建议控制支出'})
    else:
        score += 5
        details.append({'dim': '储蓄率', 'score': 5, 'max': 35, 'comment': f'储蓄率 {sr}%，需关注财务状况'})

    # 维度2：支出结构合理性（满分25）
    cat1 = data.get('cat1_list', [])
    total_exp = data.get('total_expense', 1) or 1
    top1_pct = (cat1[0][1] / total_exp * 100) if cat1 else 100

    if top1_pct < 25:
        score += 25
        details.append({'dim': '支出结构', 'score': 25, 'max': 25, 'comment': f'最大类别占比 {top1_pct:.0f}%，支出分散健康'})
    elif top1_pct < 45:
        score += 17
        details.append({'dim': '支出结构', 'score': 17, 'max': 25, 'comment': f'最大类别占比 {top1_pct:.0f}%，尚可但有一定集中风险'})
    else:
        score += 8
        details.append({'dim': '支出结构', 'score': 8, 'max': 25, 'comment': f'最大类别占比 {top1_pct:.0f}%，支出过于集中'})

    # 维度3：收入多样性（满分10）
    income_list = data.get('income_list', [])
    income_count = len(income_list)
    if income_count >= 3:
        score += 10
        details.append({'dim': '收入来源', 'score': 10, 'max': 10, 'comment': f'{income_count} 个收入来源，多元化良好'})
    elif income_count >= 2:
        score += 7
        details.append({'dim': '收入来源', 'score': 7, 'max': 10, 'comment': f'{income_count} 个收入来源，基本多元'})
    else:
        score += 3
        details.append({'dim': '收入来源', 'score': 3, 'max': 10, 'comment': '收入来源单一，建议开拓副业或投资'})

    # 维度4：交易稳定性（满分15）
    daily_avg = data.get('daily_avg', 0)
    daily_list = data.get('daily_list', [])
    if daily_list and len(daily_list) > 5:
        amounts = [d[1] for d in daily_list if d[1] > 0]
        if len(amounts) > 3:
            mean_amt = sum(amounts) / len(amounts)
            variance = sum((a - mean_amt) ** 2 for a in amounts) / len(amounts)
            std = math.sqrt(variance)
            cv = (std / mean_amt * 100) if mean_amt > 0 else 999
            if cv < 50:
                score += 15
                details.append({'dim': '消费稳定', 'score': 15, 'max': 15, 'comment': f'日消费波动率 {cv:.0f}%，支出习惯稳定'})
            elif cv < 100:
                score += 10
                details.append({'dim': '消费稳定', 'score': 10, 'max': 15, 'comment': f'日消费波动率 {cv:.0f}%，有一定波动'})
            else:
                score += 4
                details.append({'dim': '消费稳定', 'score': 4, 'max': 15, 'comment': f'日消费波动率 {cv:.0f}%，支出不太规律'})
        else:
            score += 8
            details.append({'dim': '消费稳定', 'score': 8, 'max': 15, 'comment': '数据不足，无法全面评估'})
    else:
        score += 8
        details.append({'dim': '消费稳定', 'score': 8, 'max': 15, 'comment': '日数据不足'})

    # 维度5：无异常大额（满分15）
    top_expenses = data.get('top_expenses', [])
    if top_expenses and total_exp > 0:
        largest_pct = top_expenses[0]['amount'] / total_exp * 100
        if largest_pct < 10:
            score += 15
            details.append({'dim': '支出均匀', 'score': 15, 'max': 15, 'comment': '无异常大额支出'})
        elif largest_pct < 25:
            score += 10
            details.append({'dim': '支出均匀', 'score': 10, 'max': 15, 'comment': f'最大单笔占 {largest_pct:.0f}%，尚属合理'})
        else:
            score += 4
            details.append({'dim': '支出均匀', 'score': 4, 'max': 15, 'comment': f'最大单笔占 {largest_pct:.0f}%，存在大额异常支出'})
    else:
        score += 10
        details.append({'dim': '支出均匀', 'score': 10, 'max': 15, 'comment': '数据不足'})

    # 评级
    if score >= 90:
        grade, color, grade_text = 'A+', '#7BC8A4', '财务自由'
    elif score >= 75:
        grade, color, grade_text = 'A', '#7BC8A4', '非常健康'
    elif score >= 60:
        grade, color, grade_text = 'B', '#FFD4B8', '健康良好'
    elif score >= 40:
        grade, color, grade_text = 'C', '#FFB3C6', '需要关注'
    else:
        grade, color, grade_text = 'D', '#FF6B8A', '需及时调整'

    # 建议
    suggestions = []
    if sr < 40:
        suggestions.append('💡 储蓄率偏低，建议设定月度存款目标')
    if top1_pct > 45:
        suggestions.append(f'📊 {cat1[0][0]}支出占比偏高，可以关注是否有优化空间')
    if income_count < 2:
        suggestions.append('💰 收入来源单一，可考虑投资理财增加被动收入')
    if '餐饮' in [c[0] for c in cat1]:
        food_amt = [c[1] for c in cat1 if c[0] == '餐饮'][0]
        food_pct = food_amt / total_exp * 100
        if food_pct > 30:
            suggestions.append('🍳 餐饮占比较高，多做饭可有效降低开支')

    return {
        'score': score,
        'grade': grade,
        'grade_text': grade_text,
        'color': color,
        'details': details,
        'suggestions': suggestions,
    }


# ================================================================
#  2. 预算管理
# ================================================================

DEFAULT_BUDGET = {
    "monthly_total": 8000,
    "categories": {
        "餐饮": 2000,
        "交通": 500,
        "购物": 800,
        "住房": 1000,
        "娱乐": 500,
        "通讯": 200,
    }
}


def load_budget():
    """加载预算配置，文件不存在则返回默认值。"""
    if os.path.exists(BUDGET_FILE):
        try:
            with open(BUDGET_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError):
            pass
    return dict(DEFAULT_BUDGET)


def save_budget(config):
    """保存预算配置。"""
    with open(BUDGET_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=2)
    return True


def compare_budget(data, budget_config=None):
    """对比实际支出 vs 预算，返回各分类的对比结果。"""
    if budget_config is None:
        budget_config = load_budget()

    cat1_list = data.get('cat1_list', [])
    cat_budgets = budget_config.get('categories', {})
    total_budget = budget_config.get('monthly_total', 0)
    total_expense = data.get('total_expense', 0)

    comparisons = []
    for cat_name, actual_amt in cat1_list:
        budget_amt = cat_budgets.get(cat_name, 0)
        if budget_amt > 0:
            pct_used = round(actual_amt / budget_amt * 100, 1)
            diff = round(actual_amt - budget_amt, 2)
        else:
            pct_used = None
            diff = None
        comparisons.append({
            'category': cat_name,
            'actual': actual_amt,
            'budget': budget_amt,
            'pct_used': pct_used,
            'diff': diff,
            'status': 'over' if (pct_used and pct_used > 100) else ('ok' if pct_used else 'no_budget'),
        })

    total_pct = round(total_expense / total_budget * 100, 1) if total_budget > 0 else None
    return {
        'total_budget': total_budget,
        'total_actual': total_expense,
        'total_pct': total_pct,
        'categories': comparisons,
    }


# ================================================================
#  3. 自定义规则引擎
# ================================================================

DEFAULT_RULES = [
    {"id": 1, "field": "note", "op": "contains", "value": "外卖", "set_tag": "外卖"},
    {"id": 2, "field": "cat1", "op": "eq", "value": "餐饮",
     "condition": {"field": "amount", "op": "lt", "value": 50}, "set_tag": "小食"},
]


def load_rules():
    """加载规则配置。"""
    if os.path.exists(RULES_FILE):
        try:
            with open(RULES_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError):
            pass
    return list(DEFAULT_RULES)


def save_rules(rules):
    """保存规则配置。"""
    with open(RULES_FILE, 'w', encoding='utf-8') as f:
        json.dump(rules, f, ensure_ascii=False, indent=2)
    return True


def _match_condition(record, condition):
    """检查单条记录是否满足条件。"""
    field = condition['field']
    op = condition['op']
    value = condition['value']

    rec_val = record.get(field, '')
    if isinstance(value, (int, float)) and isinstance(rec_val, str):
        try:
            rec_val = float(rec_val)
        except (ValueError, TypeError):
            return False

    if op == 'contains':
        return str(value).lower() in str(rec_val).lower()
    elif op == 'eq':
        return rec_val == value
    elif op == 'lt':
        return rec_val < value
    elif op == 'gt':
        return rec_val > value
    elif op == 'lte':
        return rec_val <= value
    elif op == 'gte':
        return rec_val >= value
    elif op == 'regex':
        try:
            return bool(re.search(str(value), str(rec_val)))
        except re.error:
            return False
    return False


def apply_rules(rows, rules=None):
    """对交易记录应用规则引擎，自动打标签。返回修改后的rows和标签统计。"""
    if rules is None:
        rules = load_rules()

    tags_applied = {rule['id']: 0 for rule in rules}

    for record in rows:
        for rule in rules:
            # 有附加条件时先检查条件
            if 'condition' in rule:
                if not _match_condition(record, rule['condition']):
                    continue
            # 检查主条件
            if not rule.get('condition') or _match_condition(record, rule):
                if _match_condition(record, rule):
                    existing = record.get('tag', '') or ''
                    new_tag = rule['set_tag']
                    if new_tag not in existing:
                        record['tag'] = f"{existing},{new_tag}".strip(',')
                        tags_applied[rule['id']] += 1

    return rows, tags_applied


# ================================================================
#  4. What-if 场景模拟
# ================================================================

def simulate(data, scenario):
    """运行 What-if 模拟，返回结果。

    scenario 类型：
    - {"type": "reduce_category", "category": "餐饮", "by_pct": 50}
    - {"type": "increase_income", "by_pct": 10}
    - {"type": "target_savings_rate", "target": 70}
    """
    total_income = data.get('total_income', 0)
    total_expense = data.get('total_expense', 0)
    cat1_list = data.get('cat1_list', [])

    if total_income <= 0:
        return {'error': '需要有效收入数据'}

    sim_type = scenario.get('type')

    if sim_type == 'reduce_category':
        cat_name = scenario.get('category', '')
        by_pct = scenario.get('by_pct', 0) / 100.0
        cat_amount = 0
        for c, amt in cat1_list:
            if c == cat_name:
                cat_amount = amt
                break
        saved = round(cat_amount * by_pct, 2)
        new_expense = round(total_expense - saved, 2)
        new_savings_rate = round((total_income - new_expense) / total_income * 100, 1)
        return {
            'type': 'reduce_category',
            'description': f'如果将 {cat_name} 支出减少 {scenario["by_pct"]}%',
            'saved_monthly': saved,
            'saved_yearly': round(saved * 12, 2),
            'new_expense': new_expense,
            'new_balance': round(total_income - new_expense, 2),
            'new_savings_rate': new_savings_rate,
            'old_savings_rate': data.get('savings_rate', 0),
        }

    elif sim_type == 'increase_income':
        by_pct = scenario.get('by_pct', 0) / 100.0
        new_income = round(total_income * (1 + by_pct), 2)
        new_savings_rate = round((new_income - total_expense) / new_income * 100, 1)
        return {
            'type': 'increase_income',
            'description': f'如果收入增加 {scenario["by_pct"]}%',
            'new_income': new_income,
            'income_increase': round(new_income - total_income, 2),
            'new_balance': round(new_income - total_expense, 2),
            'new_savings_rate': new_savings_rate,
            'old_savings_rate': data.get('savings_rate', 0),
        }

    elif sim_type == 'target_savings_rate':
        target = scenario.get('target', 0) / 100.0
        target_expense = round(total_income * (1 - target), 2)
        cut_needed = round(total_expense - target_expense, 2) if total_expense > target_expense else 0
        pct_cut = round(cut_needed / total_expense * 100, 1) if total_expense > 0 else 0
        return {
            'type': 'target_savings_rate',
            'description': f'要达到 {scenario["target"]}% 储蓄率',
            'target_expense': target_expense,
            'cut_needed': cut_needed,
            'pct_cut_needed': pct_cut,
            'new_balance': round(total_income - target_expense, 2),
            'achievable': cut_needed <= total_expense * 0.3,
        }

    return {'error': f'未知的模拟类型: {sim_type}'}


# ================================================================
#  5. 分享功能
# ================================================================

def create_share(data, charts=None):
    """创建分享快照，返回 share_id。"""
    share_id = uuid.uuid4().hex[:10]
    snapshot = {
        'share_id': share_id,
        'created_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'data': data,
    }
    filepath = os.path.join(SHARES_DIR, f'{share_id}.json')
    # 序列化：datetime 等不可序列化对象已在 data 中处理为字符串
    with open(filepath, 'w', encoding='utf-8') as f:
        json.dump(snapshot, f, ensure_ascii=False, indent=2, default=str)
    return share_id


def load_share(share_id):
    """加载分享快照。"""
    filepath = os.path.join(SHARES_DIR, f'{share_id}.json')
    if not os.path.exists(filepath):
        return None
    with open(filepath, 'r', encoding='utf-8') as f:
        return json.load(f)


# ================================================================
#  6. 数据持久化（用于环比对比）
# ================================================================

def save_month_data(data):
    """保存本月分析结果，用于下月环比。"""
    month = data.get('month', 'unknown')
    filename = f"{month.replace('年','-').replace('月','')}.json"
    filepath = os.path.join(DATA_DIR, filename)
    with open(filepath, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2, default=str)
    return filename


def load_prev_month(current_month):
    """尝试加载上个月的数据。"""
    try:
        parts = current_month.replace('年', ' ').replace('月', '').split()
        if len(parts) == 2:
            y, m = int(parts[0]), int(parts[1])
            if m == 1:
                y, m = y - 1, 12
            else:
                m -= 1
            prev_month = f'{y}年{m:02d}月'
            filename = f"{prev_month.replace('年','-').replace('月','')}.json"
            filepath = os.path.join(DATA_DIR, filename)
            if os.path.exists(filepath):
                with open(filepath, 'r', encoding='utf-8') as f:
                    return json.load(f)
    except (ValueError, IndexError):
        pass
    return None
