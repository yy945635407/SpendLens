"""粉色卡通风 PDF 报告生成模块。使用 fpdf2，横向布局参考 PPT。"""
from io import BytesIO
from fpdf import FPDF
from fpdf.enums import XPos, YPos, Align
import os
import glob


def _find_pdf_font():
    """跨平台查找中文字体，返回 ttf 路径。"""
    base = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        # 项目目录下自带字体（最优先，保证跨平台一致性）
        os.path.join(base, 'fonts', 'wqy-microhei.ttc'),
        os.path.join(base, 'fonts', 'NotoSansSC-Regular.ttf'),
    ] + glob.glob(os.path.join(base, 'fonts', '*.ttc')) \
      + glob.glob(os.path.join(base, 'fonts', '*.ttf')) \
      + [
        # macOS
        '/Library/Fonts/Arial Unicode.ttf',
        '/System/Library/Fonts/PingFang.ttc',
        '/System/Library/Fonts/STHeiti Light.ttc',
        # Linux (Railway / Ubuntu) — Noto CJK
        '/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc',
        '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc',
        '/usr/share/fonts/noto-cjk/NotoSansCJK-Regular.ttc',
        # Linux — WenQuanYi (更轻量可靠)
        '/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc',
        '/usr/share/fonts/truetype/wqy/wqy-microhei.ttc',
    ] + glob.glob('/usr/share/fonts/truetype/noto/NotoSans*') \
      + glob.glob('/usr/share/fonts/opentype/noto/NotoSans*') \
      + glob.glob('/usr/share/fonts/truetype/wqy/*.ttc') \
      + glob.glob('/usr/share/fonts/truetype/wqy/*.ttf') \
      + glob.glob('/usr/share/fonts/**/*.ttc') \
      + glob.glob('/usr/share/fonts/**/*.ttf')
    for fp in candidates:
        if os.path.isfile(fp):
            return fp
    raise FileNotFoundError('未找到中文字体，请安装 fonts-wqy-zenhei 或放置字体到 fonts/ 目录')


PDF_FONT_PATH = _find_pdf_font()


# ---- 粉色卡通风色板 ----
C = {
    'darkBg':   (0x5C, 0x2D, 0x3E),
    'primary':  (0xFF, 0x6B, 0x8A),
    'pink2':    (0xFF, 0x85, 0xA2),
    'pink3':    (0xFF, 0x9E, 0xBB),
    'pink4':    (0xFF, 0xB3, 0xC6),
    'secondary':(0x7B, 0xC8, 0xA4),
    'gold':     (0xFF, 0xD4, 0xB8),
    'text':     (0x3D, 0x1E, 0x2A),
    'muted':    (0xC4, 0x90, 0x9E),
    'white':    (0xFF, 0xFF, 0xFF),
    'bg':       (0xFF, 0xF0, 0xF5),
    'positive': (0x7B, 0xC8, 0xA4),
    'rose':     (0xFF, 0x7E, 0xB3),
}


class BillPDF(FPDF):
    """定制 PDF 类，带页眉/页脚。"""

    def header(self):
        if self.page_no() == 1:
            return
        self.set_font('CN', '', 8)
        self.set_text_color(*C['muted'])
        self.cell(0, 10, 'SpendLens  ·  账单分析报告', align='C')
        self.ln(8)
        self.set_draw_color(*C['pink4'])
        self.line(self.l_margin, self.get_y(), self.w - self.r_margin, self.get_y())
        self.ln(6)

    def footer(self):
        if self.page_no() == 1:
            return
        self.set_y(-15)
        self.set_font('CN', '', 8)
        self.set_text_color(*C['muted'])
        self.cell(0, 10, f'Page {self.page_no() - 1}', align='C')


def _rgb_tuple(val):
    """Convert hex string '#RGB' or tuple (R,G,B) to integer tuple for fpdf2."""
    if isinstance(val, str) and val.startswith('#'):
        h = val.lstrip('#')
        return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    if isinstance(val, tuple):
        return tuple(int(v) for v in val)
    return (0, 0, 0)


def build_pdf(data, charts):
    """生成 PDF 报告，返回 BytesIO。

    Args:
        data: 分析数据字典
        charts: 图表 BytesIO 字典

    Returns:
        BytesIO: PDF 文件
    """
    pdf = BillPDF(orientation='L')
    pdf.set_auto_page_break(auto=True, margin=20)
    pdf.add_font('CN', '', PDF_FONT_PATH)
    pdf.add_font('CN', 'B', PDF_FONT_PATH)

    month = data.get('month', '未知月份')
    PW, PH = pdf.w, pdf.h  # landscape: 297 x 210

    # ================ 封面 ================
    pdf.add_page()
    # 暗色背景
    pdf.set_fill_color(*C['darkBg'])
    pdf.rect(0, 0, PW, PH, 'F')

    # 装饰圆
    pdf.set_fill_color(*C['primary'])
    pdf.circle(40, 50, 25, 'F')
    pdf.set_fill_color(*C['pink3'])
    pdf.circle(250, 150, 30, 'F')

    pdf.set_y(55)
    pdf.set_font('CN', 'B', 32)
    pdf.set_text_color(*C['white'])
    pdf.cell(0, 15, f'{month}  SpendLens', align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font('CN', '', 16)
    pdf.set_text_color(*C['pink4'])
    pdf.cell(0, 10, 'iCost 智能记账 · 让每一笔都清晰可见', align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(5)
    pdf.set_font('CN', '', 10)
    pdf.set_text_color(*C['muted'])
    pdf.cell(0, 8, 'YLYT FAMILY  ·  CONFIDENTIAL', align='C')

    # ================ 概览 ================
    pdf.add_page()
    pdf.set_fill_color(*C['bg'])
    pdf.rect(0, 0, PW, PH, 'F')

    pdf.set_font('CN', 'B', 22)
    pdf.set_text_color(*C['text'])
    pdf.cell(0, 12, '总体概况', align='L', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font('CN', '', 12)
    pdf.set_text_color(*C['muted'])
    pdf.cell(0, 8, f'{month} 收支总览', align='L', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(8)

    # 6 卡片网格 (2行 × 3列) — 横向加宽
    income_count = data.get('income_count', 0)
    expense_count = data.get('expense_count', 0)
    transaction_count = data.get('transaction_count', income_count + expense_count)
    stats = [
        ('总收入', 'Total Income', f'¥{data["total_income"]:,.0f}', f'{income_count}笔', C['positive']),
        ('总支出', 'Total Expense', f'¥{data["total_expense"]:,.0f}', f'{expense_count}笔', C['primary']),
        ('结余', 'Balance', f'¥{data["balance"]:,.0f}', f'储蓄率{data["savings_rate"]}%', C['pink2']),
        ('日均支出', 'Daily Avg', f'¥{data["daily_avg"]:,.0f}', f'{data["month"]}', C['pink3']),
        ('最大支出', 'Top Category', data['cat1_list'][0][0] if data['cat1_list'] else '-', f'¥{data["cat1_list"][0][1]:,.0f}' if data['cat1_list'] else '', C['rose']),
        ('交易笔数', 'Transactions', f'{transaction_count}笔', f'{income_count}收·{expense_count}支', C['gold']),
    ]

    col_w = 82
    row_h = 36
    gap_x = 8
    start_x = 18
    for i, (label, en, value, sub, accent) in enumerate(stats):
        x = start_x + (i % 3) * (col_w + gap_x)
        y = 55 + (i // 3) * (row_h + 8)

        pdf.set_xy(x, y)
        pdf.set_fill_color(*C['white'])
        pdf.set_draw_color(*accent)
        pdf.set_line_width(0.6)

        # 左侧 accent bar
        pdf.set_fill_color(*accent)
        pdf.rect(x, y, 3, row_h, 'F')

        # 卡片主体
        pdf.set_fill_color(*C['white'])
        pdf.rect(x + 3, y, col_w - 3, row_h, 'D')

        pdf.set_xy(x + 6, y + 4)
        pdf.set_font('CN', 'B', 15)
        pdf.set_text_color(*C['text'])
        pdf.cell(col_w - 12, 7, value, align='L')

        pdf.set_xy(x + 6, y + 14)
        pdf.set_font('CN', '', 9)
        pdf.set_text_color(*C['text'])
        pdf.cell(col_w - 12, 5, label, align='L')

        pdf.set_xy(x + 6, y + 19)
        pdf.set_font('CN', '', 7)
        pdf.set_text_color(*C['muted'])
        pdf.cell(col_w - 12, 5, en, align='L')

        pdf.set_xy(x + 6, y + 25)
        pdf.set_font('CN', '', 8)
        pdf.set_text_color(*accent)
        pdf.cell(col_w - 12, 5, sub, align='L')

    # 健康评分
    if data.get('health'):
        h = data['health']
        pdf.set_xy(start_x, 150)
        pdf.set_font('CN', 'B', 14)
        pdf.set_text_color(*C['text'])
        pdf.cell(0, 10, '🌸 财务健康评分', align='L', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.set_xy(start_x, pdf.get_y())
        pdf.set_font('CN', 'B', 28)
        pdf.set_text_color(*_rgb_tuple(h['color']))
        pdf.cell(30, 12, f"{h['score']}分  {h['grade']}", align='L')
        pdf.set_font('CN', '', 11)
        pdf.set_text_color(*C['muted'])
        pdf.ln(12)
        pdf.cell(0, 6, h['grade_text'], align='L', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        # Health details
        if h.get('details'):
            health_lines = []
            for detail in h['details'][:3]:
                health_lines.append(f"• {detail['dim']}: {detail['comment']}")
            pdf.set_font('CN', '', 9)
            pdf.set_text_color(*C['text'])
            for line in health_lines:
                pdf.cell(0, 6, line, align='L', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        # Suggestions
        if h.get('suggestions'):
            pdf.ln(2)
            pdf.set_font('CN', '', 9)
            pdf.set_text_color(*C['primary'])
            pdf.cell(0, 6, '💡 ' + ' · '.join(h['suggestions'][:2]), align='L', new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    # ================ 支出结构 ================
    pdf.add_page()
    pdf.set_font('CN', 'B', 22)
    pdf.set_text_color(*C['text'])
    pdf.cell(0, 12, '支出结构分析', align='L', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(5)

    if charts.get('pie_spending'):
        chart_buf = charts['pie_spending']
        chart_buf.seek(0)
        pdf.image(chart_buf, x=10, y=pdf.get_y(), w=130)

    pdf.set_xy(155, 45)
    pdf.set_font('CN', 'B', 13)
    pdf.set_text_color(*C['text'])
    pdf.cell(0, 8, '支出分类明细', align='L', new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    for i, (cat, amt) in enumerate(data.get('cat1_list', [])[:8]):
        pct = amt / data['total_expense'] * 100
        pdf.set_xy(155, 55 + i * 10)
        pdf.set_font('CN', '', 10)
        pdf.set_text_color(*C['text'])
        pdf.cell(55, 7, cat[:8], align='L')
        pdf.set_font('CN', '', 10)
        pdf.cell(65, 7, f'¥{amt:,.0f}   ({pct:.1f}%)', align='R')

    # Spending structure analysis
    top1 = data['cat1_list'][0] if data.get('cat1_list') else ('无', 0)
    top1_pct = top1[1] / data['total_expense'] * 100 if data.get('total_expense', 1) > 0 else 0
    pdf.set_xy(155, 145)
    pdf.set_font('CN', 'B', 10)
    pdf.set_text_color(*C['text'])
    if top1_pct > 40:
        pdf.multi_cell(120, 6, f'📊 最大类别「{top1[0]}」占{top1_pct:.0f}%，支出较为集中，可关注优化空间', align='L')
    elif top1_pct > 25:
        pdf.multi_cell(120, 6, f'📊 最大类别「{top1[0]}」占{top1_pct:.0f}%，支出结构基本合理', align='L')
    else:
        pdf.multi_cell(120, 6, f'📊 最大类别仅占{top1_pct:.0f}%，支出分散健康', align='L')

    # ================ 餐饮细分 ================
    pdf.add_page()
    pdf.set_font('CN', 'B', 22)
    pdf.set_text_color(*C['text'])
    pdf.cell(0, 12, '餐饮细分', align='L', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font('CN', '', 12)
    pdf.set_text_color(*C['positive'])
    pdf.cell(0, 8, f'日均餐饮 ¥{data["daily_food"]:.2f}', align='L', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(5)

    if charts.get('bar_food'):
        chart_buf = charts['bar_food']
        chart_buf.seek(0)
        pdf.image(chart_buf, x=10, y=pdf.get_y(), w=140)

    # Food analysis
    food_cats = {c[0]: c[1] for c in data.get('food_list', [])}
    total_food = sum(f[1] for f in data.get('food_list', [])) or 1
    cook_amt = food_cats.get('三餐', 0) + food_cats.get('做饭材料', 0)
    cook_ratio = cook_amt / total_food * 100
    pdf.set_xy(155, 45)
    pdf.set_font('CN', 'B', 10)
    pdf.set_text_color(*C['text'])
    if cook_ratio > 60:
        pdf.multi_cell(120, 6, f'✅ 以做饭为主（{cook_ratio:.0f}%），饮食结构健康\n🍳 做饭占比{cook_ratio:.0f}%，日均餐饮¥{data["daily_food"]:.2f}', align='L')
    else:
        pdf.multi_cell(120, 6, f'⚠️ 外卖占比较高，多做饭更省钱\n📊 日均餐饮¥{data["daily_food"]:.2f}', align='L')

    # ================ 每周趋势 ================
    pdf.add_page()
    pdf.set_font('CN', 'B', 22)
    pdf.set_text_color(*C['text'])
    pdf.cell(0, 12, '每周支出趋势', align='L', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(5)

    if charts.get('line_weekly'):
        chart_buf = charts['line_weekly']
        chart_buf.seek(0)
        pdf.image(chart_buf, x=10, y=pdf.get_y(), w=170)

    # Weekly analysis
    wl = data.get('weekly_list', [])
    if len(wl) >= 2:
        peak_week = max(wl, key=lambda x: x[1])
        low_week = min(wl, key=lambda x: x[1])
        weekly_avg = sum(w[1] for w in wl) / len(wl)
        pdf.set_xy(18, 160)
        pdf.set_font('CN', 'B', 10)
        pdf.set_text_color(*C['text'])
        pdf.multi_cell(250, 6,
            f'📈 周均支出 ¥{weekly_avg:,.0f} · 峰值 {peak_week[0]} ¥{peak_week[1]:,.0f} · 低谷 {low_week[0]} ¥{low_week[1]:,.0f}\n'
            f'{"⚠️ 峰值周存在大额支出，建议关注是否可优化" if peak_week[1] > weekly_avg * 1.5 else "✅ 周间支出波动在正常范围"}',
            align='L')

    # ================ 每日热力图 ================
    if charts.get('heatmap'):
        pdf.add_page()
        pdf.set_font('CN', 'B', 22)
        pdf.set_text_color(*C['text'])
        pdf.cell(0, 12, '每日支出热力图', align='L', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.ln(5)

        chart_buf = charts['heatmap']
        chart_buf.seek(0)
        pdf.image(chart_buf, x=10, y=pdf.get_y(), w=210)

    # ================ 预算对比 ================
    if charts.get('budget_bar') and data.get('budget_comparison'):
        pdf.add_page()
        pdf.set_font('CN', 'B', 22)
        pdf.set_text_color(*C['text'])
        pdf.cell(0, 12, '预算 vs 实际', align='L', new_x=XPos.LMARGIN, new_y=YPos.NEXT)

        bc = data['budget_comparison']
        if bc.get('total_pct'):
            pdf.set_font('CN', '', 12)
            color = C['positive'] if bc['total_pct'] <= 100 else C['primary']
            pdf.set_text_color(*color)
            pdf.cell(0, 8, f'总计: ¥{bc["total_actual"]:,.0f} / ¥{bc["total_budget"]:,.0f} ({bc["total_pct"]}%)',
                     align='L', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.ln(5)

        chart_buf = charts['budget_bar']
        chart_buf.seek(0)
        pdf.image(chart_buf, x=10, y=pdf.get_y(), w=180)

        # Budget analysis
        over_count = sum(1 for c in bc.get('categories', []) if c['status'] == 'over')
        no_budget_count = sum(1 for c in bc.get('categories', []) if c['status'] == 'no_budget')
        pdf.set_xy(18, 175)
        pdf.set_font('CN', 'B', 10)
        pdf.set_text_color(*C['text'])
        analysis_parts = []
        if over_count > 0:
            analysis_parts.append(f'🔴 {over_count} 个类别超预算')
        if no_budget_count > 0:
            analysis_parts.append(f'⚪ {no_budget_count} 个类别未列入预算')
        if analysis_parts:
            pdf.cell(0, 6, ' · '.join(analysis_parts), align='L', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf.set_font('CN', '', 9)
            pdf.set_text_color(*C['muted'])
            pdf.cell(0, 5, '建议为所有主要支出类别设定预算，以便全面掌控财务状况', align='L', new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    # ================ 总结 ================
    pdf.set_auto_page_break(auto=False)  # 防止最后一页溢出产生空白页
    pdf.add_page()
    pdf.set_fill_color(*C['darkBg'])
    pdf.rect(0, 0, PW, PH, 'F')

    pdf.set_y(25)
    pdf.set_font('CN', 'B', 28)
    pdf.set_text_color(*C['white'])
    pdf.cell(0, 15, '总结 & 建议', align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(10)

    if data.get('health'):
        h = data['health']
        pdf.set_font('CN', '', 13)
        pdf.set_text_color(*C['white'])
        pdf.cell(0, 10, f"财务健康评分: {h['score']} 分 · {h['grade']} · {h['grade_text']}", align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        for sug in h.get('suggestions', []):
            pdf.set_font('CN', '', 12)
            pdf.set_text_color(*C['pink4'])
            pdf.cell(0, 10, sug, align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.ln(6)

    # Data-driven conclusions
    sr = data.get('savings_rate', 0)
    conclusions = []
    if sr >= 60:
        conclusions.append(f'💰 储蓄率 {sr}%，远高于推荐水平，财富积累能力优秀')
    elif sr >= 40:
        conclusions.append(f'💰 储蓄率 {sr}%，财务状态健康，继续保持')
    else:
        conclusions.append(f'💰 储蓄率 {sr}%，建议设定月度存款目标逐步提高')

    conclusions.append(f'📊 月度结余 ¥{data["balance"]:,.0f} · 日均支出 ¥{data["daily_avg"]:,.0f}')

    # Food structure
    food_cats = {c[0]: c[1] for c in data.get('food_list', [])}
    total_food = sum(f[1] for f in data.get('food_list', [])) or 1
    cook_ratio = (food_cats.get('三餐', 0) + food_cats.get('做饭材料', 0)) / total_food * 100
    if cook_ratio > 60:
        conclusions.append(f'🍳 以做饭为主（{cook_ratio:.0f}%），饮食结构健康，日均餐饮 ¥{data["daily_food"]:.2f}')
    else:
        conclusions.append(f'🍳 外卖占比 {100-cook_ratio:.0f}%，多做饭可有效降低餐饮支出')

    # Spending
    top1 = data['cat1_list'][0] if data.get('cat1_list') else ('无', 0)
    top1_pct = top1[1] / data['total_expense'] * 100
    if top1_pct > 40:
        conclusions.append(f'📊 最大支出「{top1[0]}」占 {top1_pct:.0f}%，可关注是否有优化空间')
    else:
        conclusions.append(f'📊 最大支出仅占 {top1_pct:.0f}%，支出结构分散健康')

    conclusions.append(f'✨ 共 {data.get("transaction_count", 0)} 笔交易 · {data.get("income_count", 0)}收 {data.get("expense_count", 0)}支')

    for c in conclusions:
        pdf.set_font('CN', '', 11)
        pdf.set_text_color(*C['pink4'])
        pdf.cell(0, 9, c, align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.set_y(PH - 25)
    pdf.set_font('CN', '', 9)
    pdf.set_text_color(*C['muted'])
    pdf.cell(0, 8, 'Generated by iCost + AI Analysis  ·  YLYT Family', align='C')

    # 输出
    buf = BytesIO()
    pdf.output(buf)
    buf.seek(0)
    return buf
