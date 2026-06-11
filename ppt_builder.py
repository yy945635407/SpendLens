"""粉色卡通风 + 手绘图表 PPT 生成模块。"""

from io import BytesIO
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE


# ---- 粉色卡通风色板 ----
C = {
    'bg':       RGBColor(0xFF, 0xF0, 0xF5),
    'card':     RGBColor(0xFF, 0xFF, 0xFF),
    'darkBg':   RGBColor(0x5C, 0x2D, 0x3E),
    'primary':  RGBColor(0xFF, 0x6B, 0x8A),
    'pink2':    RGBColor(0xFF, 0x85, 0xA2),
    'pink3':    RGBColor(0xFF, 0x9E, 0xBB),
    'pink4':    RGBColor(0xFF, 0xB3, 0xC6),
    'secondary':RGBColor(0x7B, 0xC8, 0xA4),
    'gold':     RGBColor(0xFF, 0xD4, 0xB8),
    'text':     RGBColor(0x3D, 0x1E, 0x2A),
    'muted':    RGBColor(0xC4, 0x90, 0x9E),
    'positive': RGBColor(0x7B, 0xC8, 0xA4),
    'rose':     RGBColor(0xFF, 0x7E, 0xB3),
    'white':    RGBColor(0xFF, 0xFF, 0xFF),
}

CARD_COLORS = [C['pink2'], C['primary'], C['gold'], C['rose'], C['pink3'], C['pink4']]
SLIDE_W = Inches(10)
SLIDE_H = Inches(5.625)


def _hex_to_rgb(hex_str):
    """'#7BC8A4' → RGBColor"""
    h = hex_str.lstrip('#')
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _add_rect(slide, x, y, w, h, fill=None, shadow=False, accent=None):
    """添加矩形卡片，可选左侧 accent bar。"""
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shape.fill.background()
    if fill:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill
    shape.line.fill.background()
    if shadow:
        shape.shadow.inherit = False
    if accent:
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, Inches(0.06), h)
        bar.fill.solid()
        bar.fill.fore_color.rgb = accent
        bar.line.fill.background()


def _add_text(slide, text, x, y, w, h, size=14, color=C['text'], bold=False,
              align=PP_ALIGN.LEFT, font_name='Arial', margin=0):
    """便捷添加文本。"""
    tf = slide.shapes.add_textbox(x, y, w, h)
    tf.text_frame.word_wrap = True
    tf.text_frame.auto_size = None
    p = tf.text_frame.paragraphs[0]
    p.text = text
    p.font.size = Pt(size)
    p.font.color.rgb = color
    p.font.bold = bold
    p.font.name = font_name
    p.alignment = align
    # Set margin to 0 for precise alignment
    tf.margin_left = Pt(margin) if margin > 0 else Pt(0)
    tf.margin_right = Pt(margin) if margin > 0 else Pt(0)
    tf.margin_top = Pt(0)
    tf.margin_bottom = Pt(0)
    return tf


def _add_status_dot(slide, x, y, size, color):
    """添加一个小圆点用于状态指示（高对比度替代emoji）。"""
    dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, size, size)
    dot.fill.solid()
    dot.fill.fore_color.rgb = color
    dot.line.fill.background()
    return dot


def _add_multiline(slide, lines, x, y, w, h, size=13, color=C['text'],
                   font_name='Arial', auto_fit=True):
    """添加多行文本，lines 为 [(text, bold, color), ...]。

    当 auto_fit=True 时，自动缩小字号以确保所有内容在框内。
    """
    txBox = slide.shapes.add_textbox(x, y, w, h)
    tf = txBox.text_frame
    tf.word_wrap = True

    # 自动适配字号
    actual_size = size
    if auto_fit and lines:
        non_empty = [l for l in lines if l[0]]
        est_lines = len(non_empty)
        line_height = Pt(size) * 1.35  # 估算行高
        total_h = line_height * est_lines
        if total_h > h:
            # 缩小字号使其适配
            ratio = h / total_h
            actual_size = max(7, int(size * ratio * 0.85))

    for i, (text, bold, clr) in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = text
        p.font.size = Pt(actual_size)
        p.font.color.rgb = clr
        p.font.bold = bold
        p.font.name = font_name
        # 减小段落间距
        p.space_before = Pt(0)
        p.space_after = Pt(1)
    txBox.margin_left = Pt(0)
    txBox.margin_top = Pt(0)
    txBox.margin_bottom = Pt(0)
    return txBox


def _add_image(slide, buf, x, y, w=None, h=None, max_w=None, max_h=None):
    """从 BytesIO 添加图片，自动保持宽高比。

    指定 w 则自动算 h；指定 h 则自动算 w；同时指定则两者都使用。
    指定 max_w / max_h 会在超出时等比缩小。
    """
    from PIL import Image as PILImage
    buf.seek(0)
    img = PILImage.open(buf)
    iw, ih = img.size
    buf.seek(0)
    if w is not None and h is None:
        h = int(w * ih / iw)
    elif h is not None and w is None:
        w = int(h * iw / ih)
    elif w is None and h is None:
        raise ValueError('必须指定 w 或 h')
    # 超出最大边界时等比缩小
    if max_w is not None and w > max_w:
        scale = max_w / w
        w = max_w
        h = int(h * scale)
    if max_h is not None and h > max_h:
        scale = max_h / h
        h = max_h
        w = int(w * scale)
    return slide.shapes.add_picture(buf, x, y, w, h)


def _ensure_text_on_top(slide):
    """将所有文本框移到 Z-order 最顶层，确保文字不被图表/形状遮挡。"""
    spTree = slide.shapes._spTree
    ns_a = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    children = list(spTree)
    text_els = []
    other_els = []

    for child in children:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag == 'sp':
            txBody = child.find(f'{{{ns_a}}}txBody')
            if txBody is not None:
                text_els.append(child)
                continue
        other_els.append(child)

    if not text_els:
        return

    # Remove all, re-add: non-text first, text last (on top)
    for child in children:
        spTree.remove(child)
    for el in other_els:
        spTree.append(el)
    for el in text_els:
        spTree.append(el)


def build_ppt(data, charts):
    """根据分析数据和图表 BytesIO 生成 PPT，返回 BytesIO。

    Args:
        data: analyzer.analyze() 返回的分析字典
        charts: charts.all_charts() 返回的 {name: BytesIO} 字典

    Returns:
        BytesIO: 生成的 .pptx 文件
    """
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)

    month = data['month']

    # ================================================================
    # SLIDE 1: 封面（重新设计 — 简洁大气）
    # ================================================================
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = C['darkBg']

    # Tiny accent dots
    for dx, dy in [(3.0, 2.8), (6.5, 3.6), (2.0, 4.2)]:
        dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(dx), Inches(dy), Inches(0.12), Inches(0.12))
        dot.fill.solid(); dot.fill.fore_color.rgb = C['pink4']; dot.line.fill.background()

    # Main title area — centered, clear hierarchy
    _add_text(slide, '🌸  ' + month + '  🌸', Inches(1), Inches(1.6), Inches(8), Inches(0.5),
              size=18, color=C['pink4'], align=PP_ALIGN.CENTER)
    _add_text(slide, 'SpendLens', Inches(0.5), Inches(2.2), Inches(9), Inches(0.9),
              size=44, color=C['white'], bold=True, align=PP_ALIGN.CENTER)
    # Thin decorative line under title
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(4.2), Inches(3.2), Inches(1.6), Pt(2))
    line.fill.solid(); line.fill.fore_color.rgb = C['pink4']; line.line.fill.background()
    _add_text(slide, 'iCost 智能记账  ·  让每一笔都清晰可见', Inches(1.5), Inches(3.5), Inches(7), Inches(0.45),
              size=15, color=C['muted'], align=PP_ALIGN.CENTER)
    _add_text(slide, 'YLYT FAMILY  ·  💖  CONFIDENTIAL 💖', Inches(0), Inches(5.1), Inches(10), Inches(0.25),
              size=9, color=C['muted'], align=PP_ALIGN.CENTER)

    # ================================================================
    # SLIDE 2: 总体概况（纯卡片布局，不重复饼图）
    # ================================================================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background; bg.fill.solid(); bg.fill.fore_color.rgb = C['bg']

    _add_text(slide, '📋 总体概况', Inches(0.6), Inches(0.25), Inches(6), Inches(0.5),
              size=26, color=C['text'], bold=True)
    _add_text(slide, f'{month} 收支总览', Inches(0.6), Inches(0.7), Inches(6), Inches(0.25),
              size=11, color=C['muted'])

    income_count = data.get('income_count', 0)
    expense_count = data.get('expense_count', 0)
    transaction_count = data.get('transaction_count', income_count + expense_count)
    stats = [
        ('总收入', 'Total Income', f'¥{data["total_income"]:,.0f}', f'{income_count} 笔', C['positive']),
        ('总支出', 'Total Expense', f'¥{data["total_expense"]:,.0f}', f'{expense_count} 笔', C['primary']),
        ('结余', 'Balance', f'¥{data["balance"]:,.0f}', f'储蓄率 {data["savings_rate"]}%', C['pink2']),
        ('日均支出', 'Daily Avg', f'¥{data["daily_avg"]:,.0f}', f'{data["month"]}', C['pink3']),
        ('最大支出', 'Top Category', data['cat1_list'][0][0] if data['cat1_list'] else '-', f'¥{data["cat1_list"][0][1]:,.0f}' if data['cat1_list'] else '', C['rose']),
        ('交易笔数', 'Transactions', f'{transaction_count} 笔', f'{income_count}收 · {expense_count}支', C['gold']),
    ]

    cw, ch = Inches(2.7), Inches(1.2)
    gx, gy = Inches(0.22), Inches(0.16)
    sx, sy = Inches(0.6), Inches(1.15)

    for i, (label, en, value, sub, accent) in enumerate(stats):
        col, row = i % 3, i // 3
        cx = sx + col * (cw + gx)
        cy = sy + row * (ch + gy)
        _add_rect(slide, cx, cy, cw, ch, fill=C['card'], shadow=True, accent=accent)
        # Value — dynamic sizing to prevent overflow
        v_len = len(str(value))
        if v_len < 8:
            val_size = 20
        elif v_len < 10:
            val_size = 17
        elif v_len < 12:
            val_size = 15
        else:
            val_size = 13
        _add_text(slide, value, cx + Inches(0.3), cy + Inches(0.08), cw - Inches(0.45), Inches(0.42),
                  size=val_size, color=accent, bold=True)
        # Label (Chinese)
        _add_text(slide, label, cx + Inches(0.3), cy + Inches(0.54), cw - Inches(0.45), Inches(0.18),
                  size=11, color=C['text'])
        # English subtitle
        _add_text(slide, en, cx + Inches(0.3), cy + Inches(0.74), cw - Inches(0.45), Inches(0.16),
                  size=8, color=C['muted'])
        # Sub info
        _add_text(slide, sub, cx + Inches(0.3), cy + Inches(0.92), cw - Inches(0.45), Inches(0.22),
                  size=10, color=accent)

    # ---- 健康评分卡片（底部横条） ----
    if data.get('health'):
        h = data['health']
        h_color = h.get('color', '#7BC8A4')
        if isinstance(h_color, str):
            h_color = _hex_to_rgb(h_color)
        _add_rect(slide, Inches(0.6), Inches(4.0), Inches(8.8), Inches(1.4), fill=C['card'], shadow=True,
                  accent=h_color)
        _add_text(slide, '🌸 财务健康评分', Inches(0.9), Inches(4.1), Inches(2.5), Inches(0.25),
                  size=11, color=C['muted'])
        _add_text(slide, f"{h['score']}分  {h['grade']}", Inches(0.9), Inches(4.35), Inches(2.5), Inches(0.45),
                  size=28, color=h_color, bold=True)
        _add_text(slide, h['grade_text'], Inches(3.3), Inches(4.4), Inches(1.8), Inches(0.4),
                  size=13, color=h_color, bold=True)
        # Detailed health breakdown
        health_lines = []
        if h.get('details'):
            for detail in h['details'][:3]:
                health_lines.append(f"• {detail['dim']}: {detail['comment']}")
        _add_text(slide, '\n'.join(health_lines) if health_lines else '', Inches(5.6), Inches(4.1), Inches(3.6), Inches(1.1),
                  size=9, color=C['text'])
        if h.get('suggestions'):
            _add_text(slide, '💡 ' + ' · '.join(h['suggestions'][:2]), Inches(0.9), Inches(4.85), Inches(8.0), Inches(0.35),
                      size=9, color=C['muted'])

    # ================================================================
    # SLIDE 3: 支出结构
    # ================================================================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background; bg.fill.solid(); bg.fill.fore_color.rgb = C['bg']

    _add_text(slide, '支出结构分析', Inches(0.6), Inches(0.3), Inches(6), Inches(0.55),
              size=28, color=C['text'], bold=True)
    _add_image(slide, charts['pie_spending'], Inches(0.3), Inches(1.0),
               w=Inches(5.5), max_h=Inches(4.0))

    top3 = data['cat1_list'][:3]
    total_exp = data['total_expense']
    insights = []
    for i, (cat_name, cat_amt) in enumerate(top3):
        title = f'{"🔥" if i==0 else "🥈" if i==1 else "🥉"}  Top {i+1}: {cat_name}'
        pct = cat_amt / total_exp * 100
        desc = f'¥{cat_amt:,.2f} · 占比 {pct:.1f}%'
        insights.append((title, desc))
    for i, (title, desc) in enumerate(insights):
        iy = Inches(1.2) + i * Inches(1.2)
        _add_rect(slide, Inches(6.5), iy, Inches(3.2), Inches(1.05), fill=C['card'], shadow=True,
                  accent=[C['primary'], C['pink3'], C['positive']][i])
        _add_text(slide, title, Inches(6.8), iy + Inches(0.1), Inches(2.7), Inches(0.3),
                  size=13, color=C['text'], bold=True)
        _add_text(slide, desc, Inches(6.8), iy + Inches(0.45), Inches(2.7), Inches(0.4),
                  size=10, color=C['muted'])

    # Spending structure analysis
    top1_pct = top3[0][1] / total_exp * 100 if top3 else 0
    cat_count = len(data['cat1_list'])
    structure_analysis = f'📊 支出涵盖 {cat_count} 个类别'
    if top1_pct > 40:
        structure_analysis += f'，最大类别占比 {top1_pct:.0f}%，支出较为集中，可关注是否有优化空间'
    elif top1_pct > 25:
        structure_analysis += f'，最大类别占比 {top1_pct:.0f}%，支出结构基本合理'
    else:
        structure_analysis += f'，最大类别仅占 {top1_pct:.0f}%，支出分散健康'
    _add_text(slide, structure_analysis, Inches(6.8), Inches(4.95), Inches(2.9), Inches(0.4),
              size=10, color=C['text'])

    # ================================================================
    # SLIDE 4: 餐饮细分
    # ================================================================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background; bg.fill.solid(); bg.fill.fore_color.rgb = C['bg']

    _add_text(slide, '餐饮细分', Inches(0.6), Inches(0.25), Inches(6), Inches(0.5),
              size=26, color=C['text'], bold=True)
    _add_text(slide, f'日均餐饮 ¥{data["daily_food"]:.2f} · 结构健康', Inches(0.6), Inches(0.7),
              Inches(6), Inches(0.25), size=11, color=C['secondary'])

    _add_image(slide, charts['bar_food'], Inches(0.3), Inches(1.1),
               w=Inches(5.0), max_h=Inches(4.2))
    _add_rect(slide, Inches(5.8), Inches(1.1), Inches(3.8), Inches(3.8), fill=C['card'], shadow=True)

    _add_text(slide, '💡 餐饮分析', Inches(6.1), Inches(1.25), Inches(3.3), Inches(0.35),
              size=15, color=C['text'], bold=True)

    food_lines = []
    total_food = sum(f[1] for f in data['food_list']) or 1
    for i, (cat, amt) in enumerate(data['food_list']):
        food_lines.append((cat, True, C['text']))
        food_lines.append((f'¥{amt:,.2f} · {amt/total_food*100:.0f}%', False, C['muted']))
    food_lines.append(('', False, C['muted']))
    # Smart food analysis
    food_cats = {c[0]: c[1] for c in data['food_list']}
    cook_ratio = (food_cats.get('三餐', 0) + food_cats.get('做饭材料', 0)) / total_food * 100 if total_food > 0 else 0
    takeout_ratio = (food_cats.get('外卖', 0) + food_cats.get('零食', 0)) / total_food * 100 if total_food > 0 else 0
    if cook_ratio > 60:
        food_lines.append(('✅ 以做饭为主，饮食结构健康', True, C['secondary']))
    else:
        food_lines.append(('⚠️ 外卖占比较高，多做饭更省钱', True, C['primary']))
    food_lines.append((f'🍳 做饭占比 {cook_ratio:.0f}%  🛵 外卖零食 {takeout_ratio:.0f}%', False, C['muted']))
    food_lines.append((f'📊 日均餐饮 ¥{data["daily_food"]:.2f}，月均餐饮 ¥{total_food/30:.2f}/天', False, C['muted']))
    _add_multiline(slide, food_lines, Inches(6.1), Inches(1.8), Inches(3.3), Inches(3.2), size=10)

    # ================================================================
    # SLIDE 5: 每周趋势
    # ================================================================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background; bg.fill.solid(); bg.fill.fore_color.rgb = C['bg']

    _add_text(slide, '每周支出趋势', Inches(0.6), Inches(0.3), Inches(6), Inches(0.55),
              size=28, color=C['text'], bold=True)

    wl = data['weekly_list']
    weekly_avg = sum(w[1] for w in wl) / len(wl) if wl else 0
    peak_week = max(wl, key=lambda x: x[1]) if wl else ('', 0)
    _add_text(slide, f'周均支出 ¥{weekly_avg:,.0f} · {peak_week[0]} 社保扣款周最高',
              Inches(0.6), Inches(0.85), Inches(6), Inches(0.3), size=13, color=C['muted'])

    _add_image(slide, charts['line_weekly'], Inches(0.3), Inches(1.2),
               w=Inches(7.0), max_h=Inches(4.0))

    # Peak callout + analysis
    _add_rect(slide, Inches(8.4), Inches(1.2), Inches(1.3), Inches(2.2), fill=C['card'], shadow=True)
    _add_text(slide, '峰值', Inches(8.4), Inches(1.4), Inches(1.3), Inches(0.3),
              size=13, color=C['primary'], bold=True, align=PP_ALIGN.CENTER)
    _add_text(slide, peak_week[0], Inches(8.4), Inches(1.7), Inches(1.3), Inches(0.25),
              size=11, color=C['text'], align=PP_ALIGN.CENTER)
    _add_text(slide, f'¥{peak_week[1]:,.0f}', Inches(8.4), Inches(2.0), Inches(1.3), Inches(0.3),
              size=15, color=C['text'], bold=True, align=PP_ALIGN.CENTER)

    # Weekly analysis
    if len(wl) >= 2:
        sorted_weeks = sorted(wl, key=lambda x: x[1])
        low_week = sorted_weeks[0]
        trend_analysis = f'📈 周间波动: {low_week[0]} 最低 ¥{low_week[1]:,.0f}'
        if peak_week[1] > weekly_avg * 1.5:
            trend_analysis += '\n⚠️ 峰值周存在大额支出'
        _add_text(slide, trend_analysis, Inches(8.4), Inches(2.55), Inches(1.3), Inches(0.65),
                  size=8, color=C['muted'], align=PP_ALIGN.CENTER)

    # ================================================================
    # SLIDE 5.5: 日历热力图（如存在）
    # ================================================================
    if charts.get('heatmap'):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        bg = slide.background; bg.fill.solid(); bg.fill.fore_color.rgb = C['bg']

        _add_text(slide, '📅 每日支出热力图', Inches(0.6), Inches(0.3), Inches(6), Inches(0.55),
                  size=28, color=C['text'], bold=True)
        _add_text(slide, f'{month} · 颜色越深，支出越高', Inches(0.6), Inches(0.85),
                  Inches(6), Inches(0.3), size=13, color=C['muted'])

        _add_image(slide, charts['heatmap'], Inches(0.5), Inches(1.3), Inches(9.0), Inches(4.0))

        # 找出最高支出日和最低支出日
        if data.get('daily_list'):
            days_with_spend = [(d, a) for d, a in data['daily_list'] if a > 0]
            if days_with_spend:
                max_day = max(days_with_spend, key=lambda x: x[1])
                min_day = min(days_with_spend, key=lambda x: x[1])
                _add_rect(slide, Inches(0.6), Inches(5.1), Inches(4.0), Inches(0.38), fill=C['card'], shadow=False)
                _add_text(slide, f'🔥 最高日: {max_day[0]} ¥{max_day[1]:.0f}    ❄️ 最低日: {min_day[0]} ¥{min_day[1]:.0f}',
                          Inches(0.8), Inches(5.1), Inches(3.8), Inches(0.38), size=11, color=C['text'])

    # ================================================================
    # SLIDE 6: 账户分布
    # ================================================================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background; bg.fill.solid(); bg.fill.fore_color.rgb = C['bg']

    _add_text(slide, '账户支出分布', Inches(0.6), Inches(0.3), Inches(6), Inches(0.55),
              size=28, color=C['text'], bold=True)
    _add_text(slide, f'{data["account_list"][0][0]} 承担 {data["account_list"][0][1]/data["total_expense"]*100:.1f}% 的家庭支出',
              Inches(0.6), Inches(0.85), Inches(6), Inches(0.3), size=13, color=C['muted'])

    _add_image(slide, charts['bar_account'], Inches(0.3), Inches(1.3),
               w=Inches(5.8), max_h=Inches(3.8))

    acc_colors = [C['primary'], C['pink2'], C['pink3'], C['rose']]
    for i, (acct, amt) in enumerate(data['account_list'][:4]):
        pct = amt / data['total_expense'] * 100
        ay = Inches(1.3) + i * Inches(0.95)
        _add_rect(slide, Inches(6.5), ay, Inches(3.2), Inches(0.78), fill=C['card'], shadow=True)
        _add_text(slide, acct, Inches(6.8), ay + Inches(0.08), Inches(1.5), Inches(0.28),
                  size=12, color=C['text'], bold=True)
        _add_text(slide, f'{pct:.1f}%', Inches(8.3), ay + Inches(0.08), Inches(1.2), Inches(0.28),
                  size=16, color=acc_colors[i], bold=True, align=PP_ALIGN.RIGHT)
        _add_text(slide, f'¥{amt:,.2f}', Inches(6.8), ay + Inches(0.42), Inches(2.7), Inches(0.25),
                  size=10, color=C['muted'])

    # Account analysis
    if len(data['account_list']) >= 2:
        top_acc = data['account_list'][0]
        acc_analysis = f'💳 {top_acc[0]} 承担了 {top_acc[1]/data["total_expense"]*100:.1f}% 的家庭支出，为最主要支付渠道'
        if len(data['account_list']) >= 3:
            acc_analysis += f'\n📊 共使用 {len(data["account_list"])} 个账户，支付方式多样化'
        _add_text(slide, acc_analysis, Inches(6.8), Inches(4.8), Inches(2.9), Inches(0.6),
                  size=10, color=C['text'])

    # ================================================================
    # SLIDE 7: 收入来源
    # ================================================================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background; bg.fill.solid(); bg.fill.fore_color.rgb = C['bg']

    _add_text(slide, '收入来源分析', Inches(0.6), Inches(0.3), Inches(6), Inches(0.55),
              size=28, color=C['text'], bold=True)
    _add_text(slide, f'总收入 ¥{data["total_income"]:,.2f} · 月结余 ¥{data["balance"]:,.2f}',
              Inches(0.6), Inches(0.85), Inches(6), Inches(0.3), size=13, color=C['secondary'])

    _add_image(slide, charts['doughnut_income'], Inches(0.3), Inches(1.3),
               h=Inches(3.8), max_w=Inches(3.2))

    # 动态计算收入明细间距
    inc_list = data['income_list'][:6]  # 最多显示6个收入来源
    inc_count = len(inc_list)
    inc_total = data['total_income']

    card_top = Inches(1.3)
    card_h = Inches(4.0)
    title_h = Inches(0.45)
    analysis_h = Inches(0.35)
    content_top = card_top + title_h + Inches(0.05)
    content_bottom = card_top + card_h - analysis_h - Inches(0.1)
    available_h = content_bottom - content_top
    # 每个条目占用的垂直空间
    item_h = available_h / max(inc_count, 1)

    _add_rect(slide, Inches(4.4), card_top, Inches(5.2), card_h, fill=C['card'], shadow=True)
    _add_text(slide, '收入明细', Inches(4.7), Inches(1.45), Inches(4.5), Inches(0.35),
              size=16, color=C['text'], bold=True)

    inc_colors = [C['positive'], C['primary'], C['pink3'], C['muted'], C['rose'], C['pink2']]
    for i, (acct, amt) in enumerate(inc_list):
        pct = amt / inc_total * 100 if inc_total > 0 else 0
        iiy = content_top + i * item_h
        # 自适应字号
        dyn_size = max(9, min(12, int(item_h / 12700 * 0.38)))
        _add_text(slide, acct, Inches(4.7), iiy, Inches(2.5), Inches(item_h * 0.45),
                  size=dyn_size, color=C['text'])
        _add_text(slide, f'¥{amt:,.2f}', Inches(7.5), iiy, Inches(1.8), Inches(item_h * 0.45),
                  size=dyn_size, color=C['text'], bold=True, align=PP_ALIGN.RIGHT)
        _add_text(slide, f'{pct:.1f}%', Inches(4.7), iiy + item_h * 0.48, Inches(2.5), Inches(item_h * 0.35),
                  size=max(8, dyn_size - 2), color=inc_colors[i % len(inc_colors)])

    # Income diversity analysis
    if inc_count >= 3:
        inc_analysis = f'✅ 收入来源 {inc_count} 个，多元化良好'
    elif inc_count >= 2:
        inc_analysis = f'💡 收入来源 {inc_count} 个，基本多元'
    else:
        inc_analysis = f'⚠️ 收入来源单一，建议开拓副业'
    inc_analysis += f' | 结余 ¥{data["balance"]:,.0f}'
    _add_text(slide, inc_analysis, Inches(4.7), Inches(5.0), Inches(4.5), Inches(0.25),
              size=9, color=C['text'])

    # ================================================================
    # SLIDE 7.5: 预算对比（如存在）
    # ================================================================
    if charts.get('budget_bar') and data.get('budget_comparison'):
        bc = data['budget_comparison']
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        bg = slide.background; bg.fill.solid(); bg.fill.fore_color.rgb = C['bg']

        _add_text(slide, '💸 预算 vs 实际', Inches(0.6), Inches(0.3), Inches(6), Inches(0.55),
                  size=28, color=C['text'], bold=True)

        if bc.get('total_pct') is not None:
            status_color = C['secondary'] if bc['total_pct'] <= 100 else C['primary']
            _add_text(slide, f'总支出 ¥{bc["total_actual"]:,.0f} / 预算 ¥{bc["total_budget"]:,.0f} · {bc["total_pct"]}%',
                      Inches(0.6), Inches(0.85), Inches(6), Inches(0.3), size=13, color=status_color)

        _add_image(slide, charts['budget_bar'], Inches(0.3), Inches(1.2),
                   w=Inches(6.2), max_h=Inches(4.0))

        # 右侧：只显示有预算的分类，用彩色圆点替代emoji（高对比度）
        budgeted = [c for c in bc.get('categories', []) if c['status'] != 'no_budget']
        if budgeted:
            count = len(budgeted)
            card_top = Inches(1.2)
            card_h = Inches(4.0)
            card_x = Inches(7.0)
            card_w = Inches(2.7)

            _add_rect(slide, card_x, card_top, card_w, card_h, fill=C['card'], shadow=True)
            _add_text(slide, '预算执行', Inches(7.2), Inches(1.3), Inches(2.4), Inches(0.35),
                      size=15, color=C['text'], bold=True)

            # 动态计算每行高度
            content_top_emu = Inches(1.85)
            content_bottom_emu = Inches(5.05)
            available_h_emu = content_bottom_emu - content_top_emu
            row_h = available_h_emu / max(count, 1)

            for i, c in enumerate(budgeted):
                row_y = content_top_emu + i * row_h
                is_over = c['status'] == 'over'
                dot_color = C['primary'] if is_over else C['secondary']
                text_color = C['primary'] if is_over else C['secondary']
                label = '超支' if is_over else '剩余'
                diff_val = abs(c['diff']) if c['diff'] is not None else 0

                # 彩色实心圆点（取代看不清的emoji）
                dot_size = Inches(0.16)
                _add_status_dot(slide, Inches(7.3), row_y + row_h * 0.1, dot_size, dot_color)

                # 类别名
                cat_size = max(10, min(13, int(row_h / 12700 * 0.55)))
                _add_text(slide, c['category'], Inches(7.55), row_y, Inches(1.5), row_h * 0.55,
                          size=cat_size, color=C['text'], bold=True)
                # 金额信息
                info = f'{label} ¥{diff_val:,.0f}（{c["pct_used"]}%）'
                info_size = max(8, cat_size - 2)
                _add_text(slide, info, Inches(7.55), row_y + row_h * 0.52, Inches(1.9), row_h * 0.42,
                          size=info_size, color=text_color)
        else:
            # 没有预算分类时显示提示
            _add_rect(slide, Inches(7.0), Inches(1.2), Inches(2.7), Inches(3.0), fill=C['card'], shadow=True)
            _add_text(slide, '预算执行', Inches(7.2), Inches(1.3), Inches(2.4), Inches(0.35),
                      size=15, color=C['text'], bold=True)
            _add_text(slide, '请先设置预算\n即可查看执行情况', Inches(7.3), Inches(2.2), Inches(2.2), Inches(0.8),
                      size=11, color=C['muted'])

    # ================================================================
    # SLIDE 8: 总结 & 建议
    # ================================================================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background; bg.fill.solid(); bg.fill.fore_color.rgb = C['darkBg']

    _add_text(slide, '🌸 总结 & 建议 🌸', Inches(0.6), Inches(0.3), Inches(6), Inches(0.55),
              size=28, color=C['white'], bold=True)
    _add_text(slide, f'{month} · 财务健康度评估 💖', Inches(0.6), Inches(0.85), Inches(6), Inches(0.3),
              size=13, color=C['pink4'])

    # Build smart conclusions from data
    conclusions = []

    # 1. Savings rate analysis
    sr = data['savings_rate']
    if sr >= 60:
        conclusions.append(('💰', f'储蓄率 {sr}%，财务状态非常优秀',
            f'远高于推荐的30%储蓄率，月存 ¥{data["balance"]:,.0f}，财富积累能力卓越，可考虑将结余资金进行稳健理财 💖'))
    elif sr >= 40:
        conclusions.append(('💰', f'储蓄率 {sr}%，财务状态健康',
            f'已达推荐的储蓄水平，月存 ¥{data["balance"]:,.0f}，继续保持即可稳步积累财富 ✨'))
    else:
        conclusions.append(('💰', f'储蓄率 {sr}%，有提升空间',
            f'当前月存 ¥{data["balance"]:,.0f}，建议设定月度存款目标逐步提高储蓄率 📈'))

    # 2. Food structure
    food_cats = {c[0]: c[1] for c in data['food_list']}
    cook_amt = food_cats.get('三餐', 0) + food_cats.get('做饭材料', 0)
    total_food = sum(f[1] for f in data['food_list']) or 1
    cook_ratio = cook_amt / total_food * 100
    if cook_ratio > 60:
        conclusions.append(('🍳', f'餐饮结构健康（做饭占{cook_ratio:.0f}%）',
            f'以三餐和做饭材料为主，日均餐饮仅 ¥{data["daily_food"]:.2f}，饮食习惯良好，既健康又省钱 ✨'))
    else:
        conclusions.append(('🍳', f'餐饮中外卖占比{100-cook_ratio:.0f}%，可优化',
            f'日均餐饮 ¥{data["daily_food"]:.2f}，多做饭可有效降低餐饮支出 🍳'))

    # 3. Spending concentration
    top1 = data['cat1_list'][0] if data['cat1_list'] else ('无', 0)
    top1_pct = top1[1] / data['total_expense'] * 100 if data['total_expense'] > 0 else 0
    if top1_pct > 40:
        conclusions.append(('📊', f'最大支出类别「{top1[0]}」占{top1_pct:.1f}%',
            f'支出较为集中，可分析该类支出中是否有可压缩的非刚性消费 🔍'))
    else:
        conclusions.append(('📊', f'最大支出类别「{top1[0]}」仅占{top1_pct:.1f}%',
            f'支出结构分散健康，没有单一类别占比过高，消费习惯良好 🎯'))

    # 4. Essential vs discretionary
    essential_cats = ['餐饮', '住房', '交通', '通讯', '社保医保']
    essential_amt = sum(c[1] for c in data['cat1_list'] if c[0] in essential_cats)
    essential_pct = essential_amt / data['total_expense'] * 100 if data['total_expense'] > 0 else 0
    discretionary = data['total_expense'] - essential_amt
    conclusions.append(('🎯', f'刚性支出占{essential_pct:.0f}%，弹性支出 ¥{discretionary:,.0f}',
        f'扣除刚性支出后，可自由支配金额为 ¥{discretionary:,.0f}，消费克制有度 🍰'))

    # 5. Weekly pattern
    wl = data['weekly_list']
    if len(wl) >= 2:
        peak_week = max(wl, key=lambda x: x[1])
        low_week = min(wl, key=lambda x: x[1])
        conclusions.append(('📅', f'支出峰值在{peak_week[0]}（¥{peak_week[1]:,.0f}），低谷在{low_week[0]}（¥{low_week[1]:,.0f}）',
            f'周间支出有波动属正常现象，关注峰值周是否包含可延后的非必要支出 💕'))

    # 6. Health score summary
    if data.get('health'):
        h = data['health']
        conclusions.append(('🌸', f'财务健康评分 {h["score"]} 分 · {h["grade"]} · {h["grade_text"]}',
            f'综合储蓄率、支出结构、收入多样性和消费稳定性评估，财务状态{h["grade_text"]} 💖'))

    # 动态间距：根据结论数量自适应
    n_conclusions = len(conclusions)
    summary_start = Inches(1.35)
    summary_end = Inches(5.05)  # 留白给底部装饰线
    item_spacing = (summary_end - summary_start) / max(n_conclusions, 1)
    title_h = item_spacing * 0.48
    desc_h = item_spacing * 0.42
    title_size = max(10, min(13, int(item_spacing / 12700 * 0.55)))
    desc_size = max(8, title_size - 2)

    for i, (icon, title, desc) in enumerate(conclusions):
        cy = summary_start + i * item_spacing
        _add_text(slide, icon + '  ' + title, Inches(0.8), cy, Inches(8.8), title_h,
                  size=title_size, color=RGBColor(0xFF, 0xFF, 0xFF), bold=True)
        _add_text(slide, desc, Inches(0.8), cy + title_h, Inches(8.8), desc_h,
                  size=desc_size, color=C['muted'])

    # Bottom line
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(2), Inches(5.2), Inches(6), Pt(1))
    line.fill.solid(); line.fill.fore_color.rgb = C['primary']; line.line.fill.background()
    _add_text(slide, 'Generated by iCost + AI Analysis  ·  YLYT Family',
              Inches(0), Inches(5.3), Inches(10), Inches(0.25),
              size=8, color=C['muted'], align=PP_ALIGN.CENTER)

    # ---- 确保所有文字在最顶层 ----
    for slide in prs.slides:
        _ensure_text_on_top(slide)

    # ---- 保存 ----
    buf = BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf
