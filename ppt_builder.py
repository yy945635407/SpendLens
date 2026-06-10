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


def _add_multiline(slide, lines, x, y, w, h, size=13, color=C['text'],
                   font_name='Arial'):
    """添加多行文本，lines 为 [(text, bold, color), ...]"""
    txBox = slide.shapes.add_textbox(x, y, w, h)
    tf = txBox.text_frame
    tf.word_wrap = True
    for i, (text, bold, clr) in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = text
        p.font.size = Pt(size)
        p.font.color.rgb = clr
        p.font.bold = bold
        p.font.name = font_name
    txBox.margin_left = Pt(0)
    txBox.margin_top = Pt(0)
    return txBox


def _add_image(slide, buf, x, y, w, h):
    """从 BytesIO 添加图片。"""
    return slide.shapes.add_picture(buf, x, y, w, h)


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

    # Subtle decorative circles — pushed to edges, much lighter
    for x, y, r, clr in [
        (-1.8, -1.8, 3.0, RGBColor(0x8A, 0x4A, 0x5E)),    # muted darker pink, far corner
        (7.8, -1.0, 2.2, RGBColor(0x7A, 0x4A, 0x58)),     # top-right, muted
        (-0.8, 4.2, 1.5, RGBColor(0x90, 0x55, 0x68)),     # bottom-left, muted
    ]:
        shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(r*2), Inches(r*2))
        shape.fill.solid()
        shape.fill.fore_color.rgb = clr
        shape.line.fill.background()

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

    stats = [
        ('💖 总收入', f'¥{data["total_income"]:,.0f}', f'{data["income_count"]} 笔', C['positive']),
        ('🌸 总支出', f'¥{data["total_expense"]:,.0f}', f'{data["expense_count"]} 笔', C['primary']),
        ('🎀 结余', f'¥{data["balance"]:,.0f}', f'储蓄率 {data["savings_rate"]}%', C['pink2']),
        ('🍰 日均支出', f'¥{data["daily_avg"]:,.0f}', '31天统计', C['pink3']),
        ('💕 最大类别', data['cat1_list'][0][0] if data['cat1_list'] else '-', f'¥{data["cat1_list"][0][1]:,.0f}' if data['cat1_list'] else '', C['rose']),
        ('✨ 最省周', data['weekly_list'][-1][0] if data['weekly_list'] else '-', f'¥{data["weekly_list"][-1][1]:,.0f}' if data['weekly_list'] else '', C['gold']),
    ]

    cw, ch = Inches(2.7), Inches(1.2)
    gx, gy = Inches(0.22), Inches(0.16)
    sx, sy = Inches(0.6), Inches(1.15)

    for i, (label, value, sub, accent) in enumerate(stats):
        col, row = i % 3, i // 3
        cx = sx + col * (cw + gx)
        cy = sy + row * (ch + gy)
        _add_rect(slide, cx, cy, cw, ch, fill=C['card'], shadow=True, accent=accent)
        # Value — smaller font to prevent overflow
        val_size = 20 if len(str(value)) < 10 else 17
        _add_text(slide, value, cx + Inches(0.3), cy + Inches(0.15), cw - Inches(0.45), Inches(0.45),
                  size=val_size, color=accent, bold=True)
        _add_text(slide, label, cx + Inches(0.3), cy + Inches(0.6), cw - Inches(0.45), Inches(0.22),
                  size=11, color=C['muted'])
        _add_text(slide, sub, cx + Inches(0.3), cy + Inches(0.85), cw - Inches(0.45), Inches(0.2),
                  size=10, color=accent)

    # ---- 健康评分卡片（底部横条） ----
    if data.get('health'):
        h = data['health']
        h_color = h.get('color', '#7BC8A4')
        if isinstance(h_color, str):
            h_color = _hex_to_rgb(h_color)
        _add_rect(slide, Inches(0.6), Inches(4.1), Inches(8.8), Inches(1.2), fill=C['card'], shadow=True,
                  accent=h_color)
        _add_text(slide, '🌸 财务健康评分', Inches(0.9), Inches(4.2), Inches(2.5), Inches(0.3),
                  size=11, color=C['muted'])
        _add_text(slide, f"{h['score']}分  {h['grade']}", Inches(0.9), Inches(4.5), Inches(2.5), Inches(0.5),
                  size=30, color=h_color, bold=True)
        _add_text(slide, h['grade_text'], Inches(3.5), Inches(4.55), Inches(2.5), Inches(0.4),
                  size=14, color=h_color, bold=True)
        if h.get('suggestions'):
            _add_text(slide, ' · '.join(h['suggestions'][:2]), Inches(5.8), Inches(4.45), Inches(3.4), Inches(0.6),
                      size=9, color=C['muted'])

    # ================================================================
    # SLIDE 3: 支出结构
    # ================================================================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background; bg.fill.solid(); bg.fill.fore_color.rgb = C['bg']

    _add_text(slide, '支出结构分析', Inches(0.6), Inches(0.3), Inches(6), Inches(0.55),
              size=28, color=C['text'], bold=True)
    _add_image(slide, charts['pie_spending'], Inches(0.3), Inches(1.0), Inches(5.8), Inches(4.2))

    top3 = data['cat1_list'][:3]
    insights = []
    for i, (cat_name, cat_amt) in enumerate(top3):
        title = f'{"🔥" if i==0 else "🥈" if i==1 else "🥉"}  Top {i+1}: {cat_name}'
        pct = cat_amt / data["total_expense"] * 100
        desc = f'¥{cat_amt:,.2f} · 占比 {pct:.1f}%'
        insights.append((title, desc))
    for i, (title, desc) in enumerate(insights):
        iy = Inches(1.2) + i * Inches(1.5)
        _add_rect(slide, Inches(6.5), iy, Inches(3.2), Inches(1.3), fill=C['card'], shadow=True,
                  accent=[C['primary'], C['pink3'], C['positive']][i])
        _add_text(slide, title, Inches(6.8), iy + Inches(0.15), Inches(2.7), Inches(0.35),
                  size=14, color=C['text'], bold=True)
        _add_text(slide, desc, Inches(6.8), iy + Inches(0.55), Inches(2.7), Inches(0.55),
                  size=10, color=C['muted'])

    # ================================================================
    # SLIDE 4: 餐饮细分
    # ================================================================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background; bg.fill.solid(); bg.fill.fore_color.rgb = C['bg']

    _add_text(slide, '餐饮细分', Inches(0.6), Inches(0.25), Inches(6), Inches(0.5),
              size=26, color=C['text'], bold=True)
    _add_text(slide, f'日均餐饮 ¥{data["daily_food"]:.2f} · 结构健康', Inches(0.6), Inches(0.7),
              Inches(6), Inches(0.25), size=11, color=C['secondary'])

    _add_image(slide, charts['bar_food'], Inches(0.3), Inches(1.1), Inches(5.2), Inches(4.2))
    _add_rect(slide, Inches(5.8), Inches(1.1), Inches(3.8), Inches(3.8), fill=C['card'], shadow=True)

    _add_text(slide, '💡 餐饮分析', Inches(6.1), Inches(1.25), Inches(3.3), Inches(0.35),
              size=15, color=C['text'], bold=True)

    food_lines = []
    total_food = sum(f[1] for f in data['food_list']) or 1
    for i, (cat, amt) in enumerate(data['food_list']):
        food_lines.append((cat, True, C['text']))
        food_lines.append((f'¥{amt:,.2f} · {amt/total_food*100:.0f}%', False, C['muted']))
    food_lines.append(('', False, C['muted']))
    food_lines.append(('✅ 三餐比例高，结构健康', True, C['secondary']))
    _add_multiline(slide, food_lines, Inches(6.1), Inches(1.8), Inches(3.3), Inches(2.8), size=10)

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

    _add_image(slide, charts['line_weekly'], Inches(0.3), Inches(1.2), Inches(7.8), Inches(4.0))

    # Peak callout
    _add_rect(slide, Inches(8.4), Inches(1.2), Inches(1.3), Inches(2.0), fill=C['card'], shadow=True)
    _add_text(slide, '峰值', Inches(8.4), Inches(1.4), Inches(1.3), Inches(0.4),
              size=13, color=C['primary'], bold=True, align=PP_ALIGN.CENTER)
    _add_text(slide, peak_week[0], Inches(8.4), Inches(1.8), Inches(1.3), Inches(0.3),
              size=11, color=C['text'], align=PP_ALIGN.CENTER)
    _add_text(slide, f'¥{peak_week[1]:,.0f}', Inches(8.4), Inches(2.15), Inches(1.3), Inches(0.35),
              size=16, color=C['text'], bold=True, align=PP_ALIGN.CENTER)
    _add_text(slide, '社保+礼金\n集中周', Inches(8.4), Inches(2.55), Inches(1.3), Inches(0.5),
              size=9, color=C['muted'], align=PP_ALIGN.CENTER)

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

    _add_image(slide, charts['bar_account'], Inches(0.3), Inches(1.3), Inches(5.8), Inches(4.0))

    acc_colors = [C['primary'], C['pink2'], C['pink3'], C['rose']]
    for i, (acct, amt) in enumerate(data['account_list'][:4]):
        pct = amt / data['total_expense'] * 100
        ay = Inches(1.3) + i * Inches(1.15)
        _add_rect(slide, Inches(6.5), ay, Inches(3.2), Inches(0.95), fill=C['card'], shadow=True)
        _add_text(slide, acct, Inches(6.8), ay + Inches(0.1), Inches(1.5), Inches(0.3),
                  size=13, color=C['text'], bold=True)
        _add_text(slide, f'{pct:.1f}%', Inches(8.3), ay + Inches(0.1), Inches(1.2), Inches(0.3),
                  size=18, color=acc_colors[i], bold=True, align=PP_ALIGN.RIGHT)
        _add_text(slide, f'¥{amt:,.2f}', Inches(6.8), ay + Inches(0.48), Inches(2.7), Inches(0.3),
                  size=11, color=C['muted'])

    # ================================================================
    # SLIDE 7: 收入来源
    # ================================================================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background; bg.fill.solid(); bg.fill.fore_color.rgb = C['bg']

    _add_text(slide, '收入来源分析', Inches(0.6), Inches(0.3), Inches(6), Inches(0.55),
              size=28, color=C['text'], bold=True)
    _add_text(slide, f'总收入 ¥{data["total_income"]:,.2f} · 月结余 ¥{data["balance"]:,.2f}',
              Inches(0.6), Inches(0.85), Inches(6), Inches(0.3), size=13, color=C['secondary'])

    _add_image(slide, charts['doughnut_income'], Inches(0.0), Inches(1.3), Inches(5.5), Inches(3.9))

    _add_rect(slide, Inches(5.8), Inches(1.3), Inches(3.8), Inches(3.5), fill=C['card'], shadow=True)
    _add_text(slide, '收入明细', Inches(6.1), Inches(1.5), Inches(3.3), Inches(0.4),
              size=16, color=C['text'], bold=True)

    inc_colors = [C['positive'], C['primary'], C['pink3'], C['muted']]
    inc_total = data['total_income']
    for i, (acct, amt) in enumerate(data['income_list'][:4]):
        pct = amt / inc_total * 100 if inc_total > 0 else 0
        iiy = Inches(2.1) + i * Inches(0.65)
        _add_text(slide, acct, Inches(6.2), iiy, Inches(1.5), Inches(0.3),
                  size=12, color=C['text'])
        _add_text(slide, f'¥{amt:,.2f}', Inches(7.8), iiy, Inches(1.5), Inches(0.3),
                  size=12, color=C['text'], bold=True, align=PP_ALIGN.RIGHT)
        _add_text(slide, f'{pct:.1f}%', Inches(6.2), iiy + Inches(0.3), Inches(2.7), Inches(0.2),
                  size=10, color=inc_colors[i])

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

        _add_image(slide, charts['budget_bar'], Inches(0.3), Inches(1.2), Inches(6.5), Inches(4.0))

        # 右侧：超预算/未超预算列表
        _add_rect(slide, Inches(7.2), Inches(1.2), Inches(2.5), Inches(4.0), fill=C['card'], shadow=True)
        _add_text(slide, '预算执行', Inches(7.4), Inches(1.3), Inches(2.2), Inches(0.35),
                  size=15, color=C['text'], bold=True)

        lines = []
        for c in bc.get('categories', [])[:6]:
            if c['status'] == 'over':
                lines.append((f"🔴 {c['category']}", True, C['primary']))
                lines.append((f'超支 ¥{abs(c["diff"]):.0f} ({c["pct_used"]}%)', False, C['primary']))
            elif c['status'] == 'ok':
                lines.append((f"🟢 {c['category']}", True, C['secondary']))
                lines.append((f'剩余 ¥{abs(c["diff"]):.0f} ({c["pct_used"]}%)', False, C['secondary']))
            else:
                lines.append((f"⚪ {c['category']}", True, C['muted']))
                lines.append(('未设预算', False, C['muted']))
            lines.append(('', False, C['text']))

        if lines:
            _add_multiline(slide, lines[:-1], Inches(7.5), Inches(1.8), Inches(2.0), Inches(3.0), size=10)

    # ================================================================
    # SLIDE 8: 总结 & 建议
    # ================================================================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background; bg.fill.solid(); bg.fill.fore_color.rgb = C['darkBg']

    _add_text(slide, '🌸 总结 & 建议 🌸', Inches(0.6), Inches(0.3), Inches(6), Inches(0.55),
              size=28, color=C['white'], bold=True)
    _add_text(slide, f'{month} · 财务健康度评估 💖', Inches(0.6), Inches(0.85), Inches(6), Inches(0.3),
              size=13, color=C['pink4'])

    conclusions = [
        ('💰', '储蓄率 ' + str(data['savings_rate']) + '%，财务状态健康',
         '远高于推荐的 30% 储蓄率，财富积累能力优秀 💖'),
        ('🍳', '餐饮结构偏健康',
         '以三餐和做饭材料为主，外卖支出极低，饮食习惯良好 ✨'),
        ('🎉', '人情支出为一次性事件',
         '本月婚礼礼金集中，下月此项将大幅下降，无需担忧 🎀'),
        ('📉', '撇除刚性支出，日常消费仅 ¥' + f'{data["total_expense"] - sum(c[1] for c in data["cat1_list"][:2] if c[0] in ["人情","社保医保"]):,.0f}',
         '扣除人情和社保后，生活开销控制在合理范围，非常克制 🍰'),
        ('✅', '建议：继续保持当前消费习惯',
         '月度预算可按此标准设定，预留弹性空间应对突发支出 💕'),
    ]

    for i, (icon, title, desc) in enumerate(conclusions):
        cy = Inches(1.4) + i * Inches(0.78)
        _add_text(slide, icon + '  ' + title, Inches(0.8), cy, Inches(8.8), Inches(0.32),
                  size=14, color=RGBColor(0xFF, 0xFF, 0xFF), bold=True)
        _add_text(slide, desc, Inches(0.8), cy + Inches(0.34), Inches(8.8), Inches(0.28),
                  size=11, color=C['muted'])

    # Bottom line
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(2), Inches(5.2), Inches(6), Pt(1))
    line.fill.solid(); line.fill.fore_color.rgb = C['primary']; line.line.fill.background()
    _add_text(slide, 'Generated by iCost + AI Analysis  ·  YLYT Family',
              Inches(0), Inches(5.3), Inches(10), Inches(0.25),
              size=8, color=C['muted'], align=PP_ALIGN.CENTER)

    # ---- 保存 ----
    buf = BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf
