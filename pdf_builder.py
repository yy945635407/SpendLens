"""粉色卡通风 PDF 报告生成模块。使用 fpdf2，横向布局参考 PPT。"""
from io import BytesIO
from fpdf import FPDF
from fpdf.enums import XPos, YPos, Align
import os
import glob


def _find_pdf_font():
    """跨平台查找中文字体，返回 ttf 路径。"""
    candidates = [
        # 项目目录下自带字体（云部署兜底）
        os.path.join(os.path.dirname(os.path.abspath(__file__)), 'fonts', 'NotoSansSC-Regular.ttf'),
        # macOS
        '/Library/Fonts/Arial Unicode.ttf',
        '/System/Library/Fonts/PingFang.ttc',
        '/System/Library/Fonts/STHeiti Light.ttc',
        # Linux (Railway / Ubuntu)
        '/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc',
        '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc',
        '/usr/share/fonts/noto-cjk/NotoSansCJK-Regular.ttc',
    ] + glob.glob('/usr/share/fonts/truetype/noto/NotoSans*') \
      + glob.glob('/usr/share/fonts/opentype/noto/NotoSans*')
    for fp in candidates:
        if os.path.isfile(fp):
            return fp
    raise FileNotFoundError('未找到中文字体，请安装 fonts-noto-cjk 或放置字体到 fonts/ 目录')


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
    stats = [
        ('总收入', f'¥{data["total_income"]:,.0f}', f'{data["income_count"]}笔', C['positive']),
        ('总支出', f'¥{data["total_expense"]:,.0f}', f'{data["expense_count"]}笔', C['primary']),
        ('结余', f'¥{data["balance"]:,.0f}', f'储蓄率{data["savings_rate"]}%', C['pink2']),
        ('日均支出', f'¥{data["daily_avg"]:,.0f}', f'日均{data["expense_count"]/31:.1f}笔', C['pink3']),
        ('最大类别', data['cat1_list'][0][0], f'¥{data["cat1_list"][0][1]:,.0f}', C['rose']),
        ('最节省周', data['weekly_list'][-1][0], f'¥{data["weekly_list"][-1][1]:,.0f}', C['gold']),
    ]

    col_w = 82
    row_h = 32
    gap_x = 8
    start_x = 18
    for i, (label, value, sub, accent) in enumerate(stats):
        x = start_x + (i % 3) * (col_w + gap_x)
        y = 55 + (i // 3) * (row_h + 6)

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
        pdf.set_font('CN', 'B', 16)
        pdf.set_text_color(*C['text'])
        pdf.cell(col_w - 12, 8, value, align='L')

        pdf.set_xy(x + 6, y + 16)
        pdf.set_font('CN', '', 9)
        pdf.set_text_color(*C['muted'])
        pdf.cell(col_w - 12, 6, label, align='L')

        pdf.set_xy(x + 6, y + 22)
        pdf.set_font('CN', '', 8)
        pdf.set_text_color(*accent)
        pdf.cell(col_w - 12, 6, sub, align='L')

    # 健康评分
    if data.get('health'):
        h = data['health']
        pdf.set_xy(start_x, 140)
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
        pdf.cell(0, 6, h['grade_text'], align='L')

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

    # ================ 总结 ================
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
        for sug in h.get('suggestions', []):
            pdf.set_font('CN', '', 14)
            pdf.set_text_color(*C['pink4'])
            pdf.cell(0, 12, sug, align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.ln(8)

    conclusions = [
        f"储蓄率 {data['savings_rate']}%，财务状态健康",
        f"月度结余 ¥{data['balance']:,.0f}",
        f"日均支出 ¥{data['daily_avg']:,.0f}",
        f"共 {data['transaction_count']} 笔交易",
    ]
    for c in conclusions:
        pdf.set_font('CN', '', 12)
        pdf.set_text_color(*C['pink4'])
        pdf.cell(0, 10, c, align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.set_y(PH - 25)
    pdf.set_font('CN', '', 9)
    pdf.set_text_color(*C['muted'])
    pdf.cell(0, 8, 'Generated by iCost + AI Analysis  ·  YLYT Family', align='C')

    # 输出
    buf = BytesIO()
    pdf.output(buf)
    buf.seek(0)
    return buf
