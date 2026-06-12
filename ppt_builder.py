"""PPT 生成模块 — 双主题（粉色/蓝色），无遮挡，Z 序正确。"""

from io import BytesIO
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

try:
    from PIL import Image as PILImage
    _HAS_PIL = True
except ImportError:
    _HAS_PIL = False

_PT = 12700  # EMU per point


# ================================================================
# 工具函数
# ================================================================

def _hex_to_rgb(hex_str):
    """'#7BC8A4' → RGBColor"""
    h = hex_str.lstrip('#')
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


# ================================================================
# 双主题色板 — 所有颜色值均为 hex 字符串
# ================================================================

_PINK_HEX = {
    'bg': '#FFF0F5', 'card': '#FFFFFF', 'darkBg': '#5C2D3E',
    'primary': '#FF6B8A', 'pink2': '#FF85A2', 'pink3': '#FF9EBB', 'pink4': '#FFB3C6',
    'secondary': '#7BC8A4', 'gold': '#FFD4B8', 'text': '#3D1E2A',
    'muted': '#C4909E', 'positive': '#7BC8A4', 'rose': '#FF7EB3', 'white': '#FFFFFF',
    'danger': '#FF3D6A', 'warning': '#FFD4B8',
    'accent1': '#FFB3C6', 'accent2': '#FFE4EC', 'line': '#FFD4DF',
    'font_title': 'Arial', 'font_body': 'Arial',
}

_BLUE_HEX = {
    'bg': '#F0F5FF', 'card': '#FFFFFF', 'darkBg': '#1E2A4A',
    'primary': '#5B8DEF', 'pink2': '#7BA3F5', 'pink3': '#9BB9FB', 'pink4': '#BBCFFF',
    'secondary': '#5BC8A4', 'gold': '#B8D4FF', 'text': '#2A3550',
    'muted': '#90A4C4', 'positive': '#5BC8A4', 'rose': '#8BABF5', 'white': '#FFFFFF',
    'danger': '#EF5B8A', 'warning': '#B8D4FF',
    'accent1': '#C8DDFF', 'accent2': '#E0EBFF', 'line': '#5B8DEF',
    'font_title': 'Arial', 'font_body': 'Arial',
}


def _make_c(hex_dict):
    """将 hex 字典转为 RGBColor 字典。非 hex 值（如字体名）原样保留。"""
    result = {}
    for k, v in hex_dict.items():
        if isinstance(v, str) and v.startswith('#'):
            result[k] = _hex_to_rgb(v)
        else:
            result[k] = v
    return result


# 全局当前主题色板
C = _make_c(_PINK_HEX)

# 卡片强调色列表（用于区分多个同类卡片）
ACCENT_COLORS = ['pink2', 'primary', 'gold', 'rose', 'pink3', 'pink4']


def set_ppt_theme(theme='pink'):
    """切换 PPT 主题：'pink' 或 'blue'。"""
    global C
    h = _BLUE_HEX if theme == 'blue' else _PINK_HEX
    C = _make_c(h)


# ================================================================
# Shape 辅助函数
# ================================================================

def _add_rect(slide, x, y, w, h, fill=None, accent=None):
    """添加矩形卡片，可选左侧 accent 细条。accent 在卡片之后添加。"""
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shape.fill.background()
    if fill:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill
    shape.line.fill.background()
    if accent:
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, Inches(0.06), h)
        bar.fill.solid()
        bar.fill.fore_color.rgb = accent
        bar.line.fill.background()


def _add_text(slide, text, x, y, w, h, size=14, color=None, bold=False,
              align=PP_ALIGN.LEFT, font_name='Arial'):
    """添加文本。color 默认取主题 text 色。"""
    if color is None:
        color = C['text']
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
    tf.margin_left = Pt(0)
    tf.margin_right = Pt(0)
    tf.margin_top = Pt(0)
    tf.margin_bottom = Pt(0)


def _add_status_dot(slide, x, y, size, color):
    """实心圆点，用于状态指示。"""
    dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, size, size)
    dot.fill.solid()
    dot.fill.fore_color.rgb = color
    dot.line.fill.background()
    return dot


def _add_multiline(slide, lines, x, y, w, h, size=13, color=None,
                   font_name='Arial', auto_fit=True):
    """多行文本。lines = [(text, bold, color), ...]。auto_fit 时自动缩字。"""
    if color is None:
        color = C['text']
    txBox = slide.shapes.add_textbox(x, y, w, h)
    tf = txBox.text_frame
    tf.word_wrap = True

    actual_size = size
    if auto_fit and lines:
        non_empty = [l for l in lines if l[0]]
        est_lines = len(non_empty)
        est_total_pt = size * 1.35 * est_lines
        box_h_pt = h / _PT
        if est_total_pt > box_h_pt > 0:
            actual_size = max(7, int(size * box_h_pt / est_total_pt * 0.85))

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
        p.space_before = Pt(0)
        p.space_after = Pt(1)
    txBox.margin_left = Pt(0)
    txBox.margin_top = Pt(0)
    txBox.margin_bottom = Pt(0)
    return txBox


def _add_image(slide, buf, x, y, w=None, h=None, max_w=None, max_h=None):
    """从 BytesIO 添加图片，自动保持宽高比。"""
    buf.seek(0)
    if _HAS_PIL:
        img = PILImage.open(buf)
        iw, ih = img.size
    else:
        iw, ih = 800, 600
    buf.seek(0)

    if w is not None and h is None:
        h = int(w * ih / iw)
    elif h is not None and w is None:
        w = int(h * iw / ih)
    elif w is None and h is None:
        raise ValueError('必须指定 w 或 h')

    if max_w is not None and w > max_w:
        scale = max_w / w
        w = max_w
        h = int(h * scale)
    if max_h is not None and h > max_h:
        scale = max_h / h
        h = max_h
        w = int(w * scale)
    return slide.shapes.add_picture(buf, x, y, w, h)


# ================================================================
# 封面装饰（仅边角，绝不进入内容区）
# ================================================================

def _add_top_accent(slide, color=None):
    """顶部细色条 — 所有幻灯片的统一装饰。"""
    c = color or C['primary']
    bar = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(10), Pt(3))
    bar.fill.solid()
    bar.fill.fore_color.rgb = c
    bar.line.fill.background()


def _add_bottom_accent(slide, color=None):
    """底部细色条。"""
    c = color or C['muted']
    bar = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(5.55), Inches(10), Pt(2))
    bar.fill.solid()
    bar.fill.fore_color.rgb = c
    bar.line.fill.background()


# ================================================================
# 主函数
# ================================================================

def build_ppt(data, charts, theme='pink'):
    """根据分析数据和图表 BytesIO 生成 PPT，返回 BytesIO。

    Args:
        data: analyzer.analyze() 返回的分析字典
        charts: charts.all_charts() 返回的 {name: BytesIO} 字典
        theme: 'pink' 或 'blue'

    Returns:
        BytesIO: .pptx 文件内容
    """
    set_ppt_theme(theme)

    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)

    month = data['month']
    total_exp = data['total_expense']
    total_inc = data['total_income']

    # 常用尺寸常量
    MARGIN = Inches(0.6)
    CONTENT_W = Inches(8.8)

    # ================================================================
    # SLIDE 1 — 封面
    # ================================================================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = C['darkBg']

    # 顶部 + 底部细色条（唯一边角装饰）
    _add_top_accent(slide, C['primary'])
    _add_bottom_accent(slide, C['pink3'])

    # 文字最后添加（自然在最上层）
    _add_text(slide, month, Inches(1), Inches(1.5), Inches(8), Inches(0.45),
              size=18, color=C['pink4'], align=PP_ALIGN.CENTER)
    _add_text(slide, 'SpendLens', Inches(0.5), Inches(2.1), Inches(9), Inches(0.9),
              size=46, color=C['white'], bold=True, align=PP_ALIGN.CENTER)
    _add_text(slide, 'iCost 智能记账  ·   让每一笔都清晰可见',
              Inches(1.5), Inches(3.5), Inches(7), Inches(0.4),
              size=14, color=C['muted'], align=PP_ALIGN.CENTER)
    _add_text(slide, 'YLYT FAMILY', Inches(0), Inches(5.18), Inches(10), Inches(0.25),
              size=9, color=C['muted'], align=PP_ALIGN.CENTER)

    # ================================================================
    # SLIDE 2 — 总体概况（2×3 卡片 + 健康评分）
    # ================================================================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = C['bg']

    # 卡片背景 + accent bar（在文字之前）
    income_count = data.get('income_count', 0)
    expense_count = data.get('expense_count', 0)
    tx_count = data.get('transaction_count', income_count + expense_count)

    stats = [
        ('总收入', 'Total Income', f'¥{total_inc:,.0f}',
         f'{income_count} 笔', C['positive']),
        ('总支出', 'Total Expense', f'¥{total_exp:,.0f}',
         f'{expense_count} 笔', C['primary']),
        ('结余', 'Balance', f'¥{data["balance"]:,.0f}',
         f'储蓄率 {data["savings_rate"]}%', C['pink2']),
        ('日均支出', 'Daily Avg', f'¥{data["daily_avg"]:,.0f}',
         f'{month}', C['pink3']),
        ('最大支出', 'Top Category',
         data['cat1_list'][0][0] if data['cat1_list'] else '-',
         f'¥{data["cat1_list"][0][1]:,.0f}' if data['cat1_list'] else '',
         C['rose']),
        ('交易笔数', 'Transactions', f'{tx_count} 笔',
         f'{income_count}收 · {expense_count}支', C['gold']),
    ]

    cw, ch = Inches(2.7), Inches(1.25)
    gx, gy = Inches(0.22), Inches(0.18)
    sx, sy = MARGIN, Inches(0.55)

    # Phase 1: 卡片背景
    for i in range(6):
        col, row = i % 3, i // 3
        cx = sx + col * (cw + gx)
        cy = sy + row * (ch + gy)
        accent = stats[i][4]
        _add_rect(slide, cx, cy, cw, ch, fill=C['card'], accent=accent)

    # Phase 2: 卡片文字（最后添加）
    for i, (label, en, value, sub, accent) in enumerate(stats):
        col, row = i % 3, i // 3
        cx = sx + col * (cw + gx)
        cy = sy + row * (ch + gy)

        v_len = len(str(value))
        val_size = 20 if v_len < 8 else (17 if v_len < 10 else (15 if v_len < 12 else 13))
        _add_text(slide, value, cx + Inches(0.3), cy + Inches(0.1),
                  cw - Inches(0.45), Inches(0.42), size=val_size, color=accent, bold=True)
        _add_text(slide, label, cx + Inches(0.3), cy + Inches(0.58),
                  cw - Inches(0.45), Inches(0.18), size=11, color=C['text'])
        _add_text(slide, en, cx + Inches(0.3), cy + Inches(0.78),
                  cw - Inches(0.45), Inches(0.16), size=8, color=C['muted'])
        _add_text(slide, sub, cx + Inches(0.3), cy + Inches(0.96),
                  cw - Inches(0.45), Inches(0.22), size=10, color=accent)

    # 健康评分横条
    if data.get('health'):
        h = data['health']
        h_color = h.get('color', '#7BC8A4')
        if isinstance(h_color, str):
            h_color = _hex_to_rgb(h_color)

        hy = Inches(3.55)
        hh = Inches(1.6)
        _add_rect(slide, MARGIN, hy, CONTENT_W, hh, fill=C['card'], accent=h_color)

        _add_text(slide, '财务健康评分', MARGIN + Inches(0.3), hy + Inches(0.1),
                  Inches(2.5), Inches(0.25), size=11, color=C['muted'])
        _add_text(slide, f"{h['score']}分  {h['grade']}",
                  MARGIN + Inches(0.3), hy + Inches(0.38),
                  Inches(2.5), Inches(0.45), size=28, color=h_color, bold=True)
        _add_text(slide, h['grade_text'], MARGIN + Inches(2.9), hy + Inches(0.46),
                  Inches(2.0), Inches(0.4), size=13, color=h_color, bold=True)

        # 健康详情
        health_lines = []
        if h.get('details'):
            for d in h['details'][:3]:
                health_lines.append(f"• {d['dim']}: {d['comment']}")
        _add_text(slide, '\n'.join(health_lines) if health_lines else '',
                  MARGIN + Inches(5.6), hy + Inches(0.12),
                  Inches(3.6), Inches(1.1), size=9, color=C['text'])

        if h.get('suggestions'):
            _add_text(slide, ' · '.join(h['suggestions'][:2]),
                      MARGIN + Inches(0.3), hy + Inches(0.95),
                      Inches(8.0), Inches(0.4), size=9, color=C['muted'])

    # 标题文字（在所有卡片之上）
    _add_text(slide, '总体概况', MARGIN, Inches(0.15), Inches(6), Inches(0.4),
              size=24, color=C['text'], bold=True)
    _add_text(slide, f'{month} 收支总览', MARGIN, Inches(0.38), Inches(6), Inches(0.17),
              size=11, color=C['muted'])

    # ================================================================
    # SLIDE 3 — 支出结构
    # ================================================================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = C['bg']

    # 饼图
    _add_image(slide, charts['pie_spending'], Inches(0.3), Inches(0.9),
               w=Inches(5.5), max_h=Inches(4.2))

    # 右侧 Top 3 卡片
    top3 = data['cat1_list'][:3]
    for i, (cat_name, cat_amt) in enumerate(top3):
        pct = cat_amt / total_exp * 100
        iy = Inches(1.0) + i * Inches(1.2)
        accent = [C['primary'], C['pink3'], C['positive']][i]
        _add_rect(slide, Inches(6.3), iy, Inches(3.3), Inches(1.05),
                  fill=C['card'], accent=accent)
        _add_text(slide, f'Top {i + 1}: {cat_name}',
                  Inches(6.6), iy + Inches(0.1), Inches(2.8), Inches(0.3),
                  size=13, color=C['text'], bold=True)
        _add_text(slide, f'¥{cat_amt:,.2f} · 占比 {pct:.1f}%',
                  Inches(6.6), iy + Inches(0.5), Inches(2.8), Inches(0.4),
                  size=10, color=C['muted'])

    # 支出结构分析
    top1_pct = top3[0][1] / total_exp * 100 if top3 else 0
    cat_count = len(data['cat1_list'])
    analysis = f'支出涵盖 {cat_count} 个类别'
    if top1_pct > 40:
        analysis += f'，最大类别占比 {top1_pct:.0f}%，支出较为集中，可关注是否有优化空间'
    elif top1_pct > 25:
        analysis += f'，最大类别占比 {top1_pct:.0f}%，支出结构基本合理'
    else:
        analysis += f'，最大类别仅占 {top1_pct:.0f}%，支出分散健康'
    _add_text(slide, analysis, Inches(6.6), Inches(4.75), Inches(2.9), Inches(0.5),
              size=10, color=C['text'])

    _add_text(slide, '支出结构分析', MARGIN, Inches(0.15), Inches(6), Inches(0.4),
              size=24, color=C['text'], bold=True)

    # ================================================================
    # SLIDE 4 — 餐饮细分
    # ================================================================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = C['bg']

    _add_image(slide, charts['bar_food'], Inches(0.3), Inches(0.8),
               w=Inches(5.0), max_h=Inches(4.4))

    # 右侧分析卡片
    _add_rect(slide, Inches(5.8), Inches(0.8), Inches(3.8), Inches(4.2),
              fill=C['card'])

    total_food = sum(f[1] for f in data['food_list']) or 1
    food_cats = {c[0]: c[1] for c in data['food_list']}

    food_lines = []
    for cat, amt in data['food_list']:
        food_lines.append((cat, True, C['text']))
        food_lines.append((f'¥{amt:,.2f} · {amt / total_food * 100:.0f}%',
                           False, C['muted']))
    food_lines.append(('', False, C['muted']))

    cook_ratio = (food_cats.get('三餐', 0) + food_cats.get('做饭材料', 0)) / total_food * 100
    takeout_ratio = (food_cats.get('外卖', 0) + food_cats.get('零食', 0)) / total_food * 100

    if cook_ratio > 60:
        food_lines.append(('以做饭为主，饮食结构健康', True, C['secondary']))
    else:
        food_lines.append(('外卖占比较高，多做饭更省钱', True, C['primary']))
    food_lines.append((f'做饭占比 {cook_ratio:.0f}%    外卖零食 {takeout_ratio:.0f}%',
                       False, C['muted']))
    food_lines.append((f'日均餐饮 ¥{data["daily_food"]:.2f}',
                       False, C['muted']))

    _add_multiline(slide, food_lines, Inches(6.1), Inches(1.0),
                   Inches(3.3), Inches(3.8), size=10)

    _add_text(slide, '餐饮细分', MARGIN, Inches(0.15), Inches(6), Inches(0.4),
              size=24, color=C['text'], bold=True)
    _add_text(slide, f'日均餐饮 ¥{data["daily_food"]:.2f}',
              MARGIN, Inches(0.48), Inches(6), Inches(0.2),
              size=11, color=C['secondary'])

    # ================================================================
    # SLIDE 5 — 每周趋势
    # ================================================================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = C['bg']

    _add_image(slide, charts['line_weekly'], Inches(0.2), Inches(0.85),
               w=Inches(7.2), max_h=Inches(4.3))

    wl = data['weekly_list']
    weekly_avg = sum(w[1] for w in wl) / len(wl) if wl else 0
    peak_week = max(wl, key=lambda x: x[1]) if wl else ('', 0)
    low_week = min(wl, key=lambda x: x[1]) if wl else ('', 0)

    # 峰值卡片
    _add_rect(slide, Inches(8.0), Inches(0.85), Inches(1.6), Inches(2.0),
              fill=C['card'])
    _add_text(slide, '周峰值', Inches(8.0), Inches(1.0), Inches(1.6), Inches(0.25),
              size=12, color=C['primary'], bold=True, align=PP_ALIGN.CENTER)
    _add_text(slide, peak_week[0], Inches(8.0), Inches(1.3), Inches(1.6), Inches(0.22),
              size=10, color=C['text'], align=PP_ALIGN.CENTER)
    _add_text(slide, f'¥{peak_week[1]:,.0f}', Inches(8.0), Inches(1.55), Inches(1.6), Inches(0.3),
              size=15, color=C['text'], bold=True, align=PP_ALIGN.CENTER)

    # 趋势分析
    if len(wl) >= 2:
        trend = f'周均 ¥{weekly_avg:,.0f}'
        if peak_week[1] > weekly_avg * 1.5:
            trend += '\n峰值周有大额支出'
        _add_text(slide, trend, Inches(8.0), Inches(2.0), Inches(1.6), Inches(0.6),
                  size=8, color=C['muted'], align=PP_ALIGN.CENTER)

    _add_text(slide, '每周支出趋势', MARGIN, Inches(0.15), Inches(6), Inches(0.4),
              size=24, color=C['text'], bold=True)
    _add_text(slide, f'周均支出 ¥{weekly_avg:,.0f} · {peak_week[0]} 为最高周',
              MARGIN, Inches(0.48), Inches(7), Inches(0.2),
              size=11, color=C['muted'])

    # ================================================================
    # SLIDE 6 — 热力图（可选）
    # ================================================================
    if charts.get('heatmap'):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        bg = slide.background
        bg.fill.solid()
        bg.fill.fore_color.rgb = C['bg']

        _add_image(slide, charts['heatmap'], Inches(0.5), Inches(1.0),
                   Inches(9.0), Inches(4.0))

        # 最高/最低日
        if data.get('daily_list'):
            days = [(d, a) for d, a in data['daily_list'] if a > 0]
            if days:
                max_day = max(days, key=lambda x: x[1])
                min_day = min(days, key=lambda x: x[1])
                _add_text(slide,
                          f'最高日: {max_day[0]} ¥{max_day[1]:.0f}    最低日: {min_day[0]} ¥{min_day[1]:.0f}',
                          MARGIN, Inches(5.12), Inches(8.8), Inches(0.3),
                          size=11, color=C['text'])

        _add_text(slide, '每日支出热力图', MARGIN, Inches(0.15), Inches(6), Inches(0.4),
                  size=24, color=C['text'], bold=True)
        _add_text(slide, f'{month} · 颜色越深，支出越高',
                  MARGIN, Inches(0.48), Inches(6), Inches(0.2),
                  size=11, color=C['muted'])

    # ================================================================
    # SLIDE 7 — 账户分布
    # ================================================================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = C['bg']

    _add_image(slide, charts['bar_account'], Inches(0.2), Inches(0.9),
               w=Inches(5.8), max_h=Inches(4.0))

    acc_colors = [C['primary'], C['pink2'], C['pink3'], C['rose']]
    for i, (acct, amt) in enumerate(data['account_list'][:4]):
        pct = amt / total_exp * 100
        ay = Inches(1.0) + i * Inches(0.95)
        _add_rect(slide, Inches(6.4), ay, Inches(3.2), Inches(0.78),
                  fill=C['card'])
        _add_text(slide, acct, Inches(6.7), ay + Inches(0.08),
                  Inches(1.5), Inches(0.28), size=12, color=C['text'], bold=True)
        _add_text(slide, f'{pct:.1f}%', Inches(8.2), ay + Inches(0.08),
                  Inches(1.2), Inches(0.28), size=16,
                  color=acc_colors[i], bold=True, align=PP_ALIGN.RIGHT)
        _add_text(slide, f'¥{amt:,.2f}', Inches(6.7), ay + Inches(0.44),
                  Inches(2.7), Inches(0.25), size=10, color=C['muted'])

    # 账户分析
    if len(data['account_list']) >= 2:
        top_acc = data['account_list'][0]
        acc_pct = top_acc[1] / total_exp * 100
        acc_analysis = f'{top_acc[0]} 承担了 {acc_pct:.1f}% 的家庭支出，为最主要支付渠道'
        if len(data['account_list']) >= 3:
            acc_analysis += f'\n共使用 {len(data["account_list"])} 个账户，支付方式多样化'
        _add_text(slide, acc_analysis, Inches(6.7), Inches(4.8),
                  Inches(2.9), Inches(0.6), size=10, color=C['text'])

    _add_text(slide, '账户支出分布', MARGIN, Inches(0.15), Inches(6), Inches(0.4),
              size=24, color=C['text'], bold=True)
    _add_text(slide,
              f'{data["account_list"][0][0]} 承担 {data["account_list"][0][1] / total_exp * 100:.1f}% 的家庭支出',
              MARGIN, Inches(0.48), Inches(7), Inches(0.2),
              size=11, color=C['muted'])

    # ================================================================
    # SLIDE 8 — 收入分析
    # ================================================================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = C['bg']

    _add_image(slide, charts['doughnut_income'], Inches(0.5), Inches(1.0),
               h=Inches(3.8), max_w=Inches(3.2))

    inc_list = data['income_list'][:6]
    inc_total = total_inc
    inc_count = len(inc_list)

    card_x = Inches(4.2)
    card_w = Inches(5.4)
    card_h = Inches(4.0)
    card_y = Inches(0.85)
    _add_rect(slide, card_x, card_y, card_w, card_h, fill=C['card'])

    _add_text(slide, '收入明细', card_x + Inches(0.3), card_y + Inches(0.1),
              Inches(4.5), Inches(0.3), size=15, color=C['text'], bold=True)

    inc_colors = [C['positive'], C['primary'], C['pink3'], C['muted'], C['rose'], C['pink2']]
    item_h = Inches(2.75) / max(inc_count, 1)
    content_top = card_y + Inches(0.55)

    for i, (name, amt) in enumerate(inc_list):
        pct = amt / inc_total * 100 if inc_total > 0 else 0
        iy = content_top + i * item_h
        dyn_size = max(9, min(12, int(item_h / _PT * 0.38)))
        _add_text(slide, name, card_x + Inches(0.3), iy,
                  Inches(2.8), item_h * 0.45, size=dyn_size, color=C['text'])
        _add_text(slide, f'¥{amt:,.2f}', card_x + Inches(3.2), iy,
                  Inches(1.8), item_h * 0.45, size=dyn_size,
                  color=C['text'], bold=True, align=PP_ALIGN.RIGHT)
        _add_text(slide, f'{pct:.1f}%', card_x + Inches(0.3),
                  iy + item_h * 0.5, Inches(2.8), item_h * 0.35,
                  size=max(8, dyn_size - 2), color=inc_colors[i % len(inc_colors)])

    # 分析
    if inc_count >= 3:
        summary = f'收入来源 {inc_count} 个，多元化良好'
    elif inc_count >= 2:
        summary = f'收入来源 {inc_count} 个，基本多元'
    else:
        summary = '收入来源单一，建议开拓副业'
    summary += f' | 结余 ¥{data["balance"]:,.0f}'
    _add_text(slide, summary, card_x + Inches(0.3), card_y + card_h - Inches(0.35),
              Inches(4.5), Inches(0.25), size=9, color=C['text'])

    _add_text(slide, '收入来源分析', MARGIN, Inches(0.15), Inches(6), Inches(0.4),
              size=24, color=C['text'], bold=True)
    _add_text(slide, f'总收入 ¥{total_inc:,.2f} · 月结余 ¥{data["balance"]:,.2f}',
              MARGIN, Inches(0.48), Inches(6), Inches(0.2),
              size=11, color=C['secondary'])

    # ================================================================
    # SLIDE 9 — 预算执行（可选）
    # ================================================================
    if charts.get('budget_bar') and data.get('budget_comparison'):
        bc = data['budget_comparison']
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        bg = slide.background
        bg.fill.solid()
        bg.fill.fore_color.rgb = C['bg']

        _add_image(slide, charts['budget_bar'], Inches(0.2), Inches(0.85),
                   w=Inches(6.2), max_h=Inches(4.2))

        budgeted = [c for c in bc.get('categories', []) if c['status'] != 'no_budget']
        if budgeted:
            card_x = Inches(7.0)
            card_w = Inches(2.6)
            card_y = Inches(0.85)
            card_h = Inches(4.2)
            _add_rect(slide, card_x, card_y, card_w, card_h,
                      fill=C['card'])

            _add_text(slide, '预算执行', card_x + Inches(0.2), card_y + Inches(0.1),
                      Inches(2.2), Inches(0.3), size=14, color=C['text'], bold=True)

            count = len(budgeted)
            item_h = Inches(3.4) / max(count, 1)
            top_y = card_y + Inches(0.55)

            for i, c in enumerate(budgeted):
                iy = top_y + i * item_h
                is_over = c['status'] == 'over'
                dot_color = C['primary'] if is_over else C['secondary']
                text_color = C['primary'] if is_over else C['secondary']
                label = '超支' if is_over else '剩余'
                diff_val = abs(c['diff']) if c['diff'] is not None else 0

                _add_status_dot(slide, card_x + Inches(0.22),
                                iy + item_h * 0.15, Inches(0.14), dot_color)

                cat_size = max(9, min(12, int(item_h / _PT * 0.5)))
                _add_text(slide, c['category'],
                          card_x + Inches(0.45), iy,
                          Inches(1.3), item_h * 0.5,
                          size=cat_size, color=C['text'], bold=True)
                info = f'{label} ¥{diff_val:,.0f}（{c["pct_used"]}%）'
                _add_text(slide, info,
                          card_x + Inches(0.45), iy + item_h * 0.5,
                          Inches(1.9), item_h * 0.4,
                          size=max(7, cat_size - 2), color=text_color)
        else:
            _add_rect(slide, Inches(7.0), Inches(0.85), Inches(2.6), Inches(3.0),
                      fill=C['card'])
            _add_text(slide, '预算执行', Inches(7.2), Inches(1.0),
                      Inches(2.2), Inches(0.3), size=14, color=C['text'], bold=True)
            _add_text(slide, '请先设置预算\n即可查看执行情况',
                      Inches(7.3), Inches(2.0), Inches(2.0), Inches(0.8),
                      size=11, color=C['muted'])

        if bc.get('total_pct') is not None:
            status_color = C['secondary'] if bc['total_pct'] <= 100 else C['primary']
            _add_text(slide,
                      f'总支出 ¥{bc["total_actual"]:,.0f} / 预算 ¥{bc["total_budget"]:,.0f} · {bc["total_pct"]}%',
                      MARGIN, Inches(0.48), Inches(7), Inches(0.2),
                      size=11, color=status_color)

        _add_text(slide, '预算 vs 实际', MARGIN, Inches(0.15), Inches(6), Inches(0.4),
                  size=24, color=C['text'], bold=True)

    # ================================================================
    # SLIDE 10 — 总结 & 建议
    # ================================================================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = C['darkBg']

    # 唯一边角装饰
    _add_top_accent(slide, C['primary'])

    # 构建智能结论
    conclusions = []

    sr = data['savings_rate']
    if sr >= 60:
        conclusions.append((f'储蓄率 {sr}%，财务状态非常优秀',
                            f'远高于推荐的30%储蓄率，月存 ¥{data["balance"]:,.0f}，财富积累能力卓越'))
    elif sr >= 40:
        conclusions.append((f'储蓄率 {sr}%，财务状态健康',
                            f'已达推荐储蓄水平，月存 ¥{data["balance"]:,.0f}，继续保持即可稳步积累'))
    else:
        conclusions.append((f'储蓄率 {sr}%，有提升空间',
                            f'当前月存 ¥{data["balance"]:,.0f}，建议设定月度存款目标逐步提高储蓄率'))

    cook_amt = food_cats.get('三餐', 0) + food_cats.get('做饭材料', 0)
    cook_ratio = cook_amt / total_food * 100 if total_food > 0 else 0
    if cook_ratio > 60:
        conclusions.append((f'餐饮结构健康（做饭占{cook_ratio:.0f}%）',
                            f'以三餐和做饭材料为主，日均餐饮仅 ¥{data["daily_food"]:.2f}'))
    else:
        conclusions.append((f'餐饮中外卖占比{100 - cook_ratio:.0f}%，可优化',
                            f'日均餐饮 ¥{data["daily_food"]:.2f}，多做饭可有效降低餐饮支出'))

    top1 = data['cat1_list'][0] if data['cat1_list'] else ('无', 0)
    top1_pct = top1[1] / total_exp * 100 if total_exp > 0 else 0
    if top1_pct > 40:
        conclusions.append((f'最大支出「{top1[0]}」占{top1_pct:.1f}%',
                            '支出较集中，可分析该类支出中是否有可压缩的非刚性消费'))
    else:
        conclusions.append((f'最大支出「{top1[0]}」仅占{top1_pct:.1f}%',
                            '支出结构分散健康，没有单一类别占比过高'))

    essential_cats = ['餐饮', '住房', '交通', '通讯', '社保医保']
    essential_amt = sum(c[1] for c in data['cat1_list'] if c[0] in essential_cats)
    essential_pct = essential_amt / total_exp * 100 if total_exp > 0 else 0
    discretionary = total_exp - essential_amt
    conclusions.append((f'刚性支出占{essential_pct:.0f}%，弹性支出 ¥{discretionary:,.0f}',
                        f'扣除刚性支出后，可自由支配金额 ¥{discretionary:,.0f}，消费克制有度'))

    if len(wl) >= 2:
        peak_week = max(wl, key=lambda x: x[1])
        low_week = min(wl, key=lambda x: x[1])
        conclusions.append((f'支出峰值 {peak_week[0]}（¥{peak_week[1]:,.0f}），低谷 {low_week[0]}（¥{low_week[1]:,.0f}）',
                            '周间支出波动属正常现象，关注峰值周是否有可延后的非必要支出'))

    if data.get('health'):
        h = data['health']
        conclusions.append((f'财务健康评分 {h["score"]} 分 · {h["grade"]} · {h["grade_text"]}',
                            f'综合储蓄率、支出结构、收入多样性和消费稳定性评估，财务状态{h["grade_text"]}'))

    # 动态间距
    n = len(conclusions)
    start_y = Inches(0.85)
    end_y = Inches(4.9)
    spacing = (end_y - start_y) / max(n, 1)

    for i, (title, desc) in enumerate(conclusions):
        cy = start_y + i * spacing
        title_h = spacing * 0.45
        desc_h = spacing * 0.42
        t_size = max(10, min(14, int(spacing / _PT * 0.5)))

        _add_text(slide, title, MARGIN + Inches(0.2), cy,
                  Inches(8.6), title_h, size=t_size, color=C['white'], bold=True)
        _add_text(slide, desc, MARGIN + Inches(0.2), cy + title_h,
                  Inches(8.6), desc_h, size=max(8, t_size - 2), color=C['muted'])

    # 底部线
    _add_text(slide, 'Generated by iCost + AI Analysis  ·  YLYT Family',
              Inches(0), Inches(5.2), Inches(10), Inches(0.2),
              size=8, color=C['muted'], align=PP_ALIGN.CENTER)

    _add_text(slide, '总结 & 建议', MARGIN, Inches(0.15), Inches(6), Inches(0.4),
              size=24, color=C['white'], bold=True)
    _add_text(slide, f'{month} · 财务健康度评估',
              MARGIN, Inches(0.48), Inches(6), Inches(0.2),
              size=13, color=C['pink4'])

    # ================================================================
    # 保存
    # ================================================================
    buf = BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf
