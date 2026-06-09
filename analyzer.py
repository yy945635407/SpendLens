"""iCost 账单数据分析模块"""

from collections import defaultdict
from datetime import datetime, timedelta
import openpyxl
from budget import apply_rules, load_rules


def analyze(filepath_or_stream, apply_auto_tags=True):
    """解析 iCost Excel 文件，返回完整分析结果。

    Args:
        filepath_or_stream: Excel 文件路径或 BytesIO 流

    Returns:
        dict: 分析结果
    """
    wb = openpyxl.load_workbook(filepath_or_stream)
    ws = wb['收支账单']

    rows = []
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
        if row[0] is None:
            continue
        rows.append({
            'date': row[0],
            'type': row[1],
            'amount': float(row[2]),
            'cat1': row[3] if row[3] else '未分类',
            'cat2': row[4] if row[4] else '',
            'account': row[5] if row[5] else '未知账户',
            'note': row[7] if row[7] else '',
            'tag': row[9] if row[9] else '',
        })

    income_rows = [r for r in rows if r['type'] == '收入']
    expense_rows = [r for r in rows if r['type'] == '支出']

    total_income = sum(r['amount'] for r in income_rows)
    total_expense = sum(r['amount'] for r in expense_rows)
    balance = total_income + total_expense  # expense is negative

    # 月份信息
    if rows:
        first_date = rows[0]['date']
        if isinstance(first_date, str):
            dt = datetime.strptime(first_date[:10], '%Y/%m/%d')
        else:
            dt = first_date
        month_str = dt.strftime('%Y年%m月')

    # ---- 一级分类占比 ----
    expense_by_cat1 = defaultdict(float)
    for r in expense_rows:
        expense_by_cat1[r['cat1']] += abs(r['amount'])
    cat1_list = sorted(expense_by_cat1.items(), key=lambda x: x[1], reverse=True)

    # ---- 二级分类 ----
    expense_by_cat2 = defaultdict(float)
    for r in expense_rows:
        key = f"{r['cat1']} > {r['cat2']}" if r['cat2'] else r['cat1']
        expense_by_cat2[key] += abs(r['amount'])
    cat2_list = sorted(expense_by_cat2.items(), key=lambda x: x[1], reverse=True)

    # ---- 餐饮细分 ----
    food_cats = defaultdict(float)
    for r in expense_rows:
        if r['cat1'] == '餐饮':
            sub = r['cat2'] if r['cat2'] else '其他'
            food_cats[sub] += abs(r['amount'])
    food_list = sorted(food_cats.items(), key=lambda x: x[1], reverse=True)
    daily_food = sum(v for _, v in food_list) / 31

    # ---- 每日统计 ----
    daily = defaultdict(float)
    date_range = set()
    for r in rows:
        d = r['date']
        if isinstance(d, str):
            dt = datetime.strptime(d[:10], '%Y/%m/%d')
        else:
            dt = d
        day_key = dt.strftime('%m-%d')
        daily[day_key] += abs(r['amount'])
        date_range.add(dt.date())
    # 生成完整日期序列（填充无支出日为0）
    daily_list = []
    if date_range:
        min_date = min(date_range)
        max_date = max(date_range)
        cur = min_date
        while cur <= max_date:
            key = cur.strftime('%m-%d')
            daily_list.append((key, round(daily.get(key, 0), 2)))
            cur += timedelta(days=1)
    else:
        daily_list = sorted(daily.items())

    # ---- 每周趋势 ----
    weekly = defaultdict(float)
    for r in expense_rows:
        d = r['date']
        if isinstance(d, str):
            dt = datetime.strptime(d[:10], '%Y/%m/%d')
        else:
            dt = d
        week_label = f"W{dt.strftime('%U')}"
        weekly[week_label] += abs(r['amount'])
    weekly_list = sorted(weekly.items())

    # ---- 账户分布 ----
    expense_by_acct = defaultdict(float)
    for r in expense_rows:
        expense_by_acct[r['account']] += abs(r['amount'])
    acct_list = sorted(expense_by_acct.items(), key=lambda x: x[1], reverse=True)

    # ---- 收入来源 ----
    income_by_acct = defaultdict(float)
    for r in income_rows:
        income_by_acct[r['account']] += r['amount']
    income_list = sorted(income_by_acct.items(), key=lambda x: x[1], reverse=True)

    # ---- Top 10 支出 ----
    top_expenses = sorted(expense_rows, key=lambda x: abs(x['amount']), reverse=True)[:10]

    # ---- 规则引擎自动打标签 ----
    if apply_auto_tags:
        try:
            rules = load_rules()
            rows, tags_applied = apply_rules(rows, rules)
        except Exception:
            tags_applied = {}

    # ---- 标签统计 ----
    tag_stats = defaultdict(lambda: {'count': 0, 'amount': 0})
    for r in rows:
        if r['tag']:
            for t in r['tag'].split(','):
                t = t.strip()
                if t:
                    tag_stats[t]['count'] += 1
                    tag_stats[t]['amount'] += abs(r['amount'])
    tag_list = sorted(tag_stats.items(), key=lambda x: x[1]['amount'], reverse=True)

    return {
        'month': month_str,
        'total_income': round(total_income, 2),
        'total_expense': round(abs(total_expense), 2),
        'balance': round(balance, 2),
        'savings_rate': round(balance / total_income * 100, 1) if total_income > 0 else 0,
        'transaction_count': len(rows),
        'income_count': len(income_rows),
        'expense_count': len(expense_rows),
        'daily_avg': round(abs(total_expense) / 31, 2),
        'cat1_list': [(cat, round(amt, 2)) for cat, amt in cat1_list],
        'cat2_list': [(cat, round(amt, 2)) for cat, amt in cat2_list[:15]],
        'food_list': [(cat, round(amt, 2)) for cat, amt in food_list],
        'daily_food': round(daily_food, 2),
        'weekly_list': [(w, round(amt, 2)) for w, amt in weekly_list],
        'account_list': [(acct, round(amt, 2)) for acct, amt in acct_list],
        'income_list': [(acct, round(amt, 2)) for acct, amt in income_list],
        'top_expenses': [{
            'date': str(r['date']),
            'amount': round(abs(r['amount']), 2),
            'cat1': r['cat1'],
            'cat2': r['cat2'],
            'note': r['note'],
        } for r in top_expenses],
        'tag_list': [(tag, data['count'], round(data['amount'], 2)) for tag, data in tag_list],
        'daily_list': daily_list,
    }


if __name__ == '__main__':
    import json
    result = analyze('/Users/ylyt_bot/Desktop/iCost_202605.xlsx')
    print(json.dumps(result, ensure_ascii=False, indent=2))


def compare(prev_data, curr_data):
    """对比两个月的数据，返回环比差异。

    Args:
        prev_data: analyze() 输出的上月数据 dict
        curr_data: analyze() 输出的本月数据 dict

    Returns:
        dict: 对比结果
    """
    def _chg(old, new):
        """计算变化率。"""
        if old and old > 0:
            return round((new - old) / old * 100, 1)
        return None

    result = {
        'prev_month': prev_data.get('month', '上月'),
        'curr_month': curr_data.get('month', '本月'),
        'income': {
            'prev': prev_data.get('total_income', 0),
            'curr': curr_data.get('total_income', 0),
            'change_pct': _chg(prev_data.get('total_income', 0), curr_data.get('total_income', 0)),
        },
        'expense': {
            'prev': prev_data.get('total_expense', 0),
            'curr': curr_data.get('total_expense', 0),
            'change_pct': _chg(prev_data.get('total_expense', 0), curr_data.get('total_expense', 0)),
        },
        'balance': {
            'prev': prev_data.get('balance', 0),
            'curr': curr_data.get('balance', 0),
            'change_pct': _chg(prev_data.get('balance', 0), curr_data.get('balance', 0)),
        },
        'savings_rate': {
            'prev': prev_data.get('savings_rate', 0),
            'curr': curr_data.get('savings_rate', 0),
            'change': round(curr_data.get('savings_rate', 0) - prev_data.get('savings_rate', 0), 1),
        },
        'daily_avg': {
            'prev': prev_data.get('daily_avg', 0),
            'curr': curr_data.get('daily_avg', 0),
            'change_pct': _chg(prev_data.get('daily_avg', 0), curr_data.get('daily_avg', 0)),
        },
        'categories': [],
    }

    # 分类环比
    prev_cats = {c[0]: c[1] for c in prev_data.get('cat1_list', [])}
    curr_cats = {c[0]: c[1] for c in curr_data.get('cat1_list', [])}
    all_cats = sorted(set(list(prev_cats.keys()) + list(curr_cats.keys())))

    for cat in all_cats:
        p_amt = prev_cats.get(cat, 0)
        c_amt = curr_cats.get(cat, 0)
        result['categories'].append({
            'name': cat,
            'prev': p_amt,
            'curr': c_amt,
            'change_pct': _chg(p_amt, c_amt),
            'change_amt': round(c_amt - p_amt, 2),
            'is_new': p_amt == 0 and c_amt > 0,
            'is_gone': p_amt > 0 and c_amt == 0,
        })

    return result
