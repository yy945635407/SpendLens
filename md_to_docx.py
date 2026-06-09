"""将竞品调研报告 Markdown 转为精美 Word 文档。"""
import re
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ---- 读取 Markdown ----
with open('竞品调研报告.md', 'r', encoding='utf-8') as f:
    md = f.read()

doc = Document()

# ---- 页面设置 ----
section = doc.sections[0]
section.page_width = Cm(21)
section.page_height = Cm(29.7)
section.top_margin = Cm(2.5)
section.bottom_margin = Cm(2.5)
section.left_margin = Cm(2.5)
section.right_margin = Cm(2.5)

# ---- 样式定义 ----
style = doc.styles['Normal']
style.font.name = 'PingFang SC'
style.font.size = Pt(11)
style.paragraph_format.space_after = Pt(6)
style.paragraph_format.line_spacing = 1.35
style.element.rPr.rFonts.set(qn('w:eastAsia'), 'PingFang SC')

# Helper: 设置段落中文间距
def set_cn_spacing(paragraph, before=0, after=6, line_spacing=1.35):
    pf = paragraph.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line_spacing

def add_heading_styled(text, level=1):
    """添加带样式的标题。"""
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        run.font.name = 'PingFang SC'
        run.element.rPr.rFonts.set(qn('w:eastAsia'), 'PingFang SC')
    if level == 1:
        h.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in h.runs:
            run.font.size = Pt(22)
            run.font.color.rgb = RGBColor(0x5C, 0x2D, 0x3E)
    elif level == 2:
        for run in h.runs:
            run.font.size = Pt(16)
            run.font.color.rgb = RGBColor(0xFF, 0x6B, 0x8A)
    elif level == 3:
        for run in h.runs:
            run.font.size = Pt(13)
            run.font.color.rgb = RGBColor(0x3D, 0x1E, 0x2A)
    return h

def add_para(text, bold=False, italic=False, color=None, size=None, alignment=None):
    """添加格式化段落。"""
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.name = 'PingFang SC'
    run.element.rPr.rFonts.set(qn('w:eastAsia'), 'PingFang SC')
    if bold:
        run.font.bold = True
    if italic:
        run.font.italic = True
    if color:
        run.font.color.rgb = color
    if size:
        run.font.size = size
    if alignment is not None:
        p.alignment = alignment
    return p

def add_table_with_data(headers, rows, col_widths=None):
    """添加格式化表格。"""
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = 'Light Grid Accent 1'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # 表头
    hdr_cells = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr_cells[i].text = h
        for p in hdr_cells[i].paragraphs:
            for run in p.runs:
                run.font.bold = True
                run.font.size = Pt(10)
                run.font.name = 'PingFang SC'
                run.element.rPr.rFonts.set(qn('w:eastAsia'), 'PingFang SC')
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        # 表头背景色
        shading = OxmlElement('w:shd')
        shading.set(qn('w:fill'), 'FF6B8A')
        shading.set(qn('w:val'), 'clear')
        hdr_cells[i]._tc.get_or_add_tcPr().append(shading)

    # 数据行
    for r, row in enumerate(rows):
        cells = table.rows[r + 1].cells
        for c, val in enumerate(row):
            cells[c].text = str(val)
            for p in cells[c].paragraphs:
                for run in p.runs:
                    run.font.size = Pt(9)
                    run.font.name = 'PingFang SC'
                    run.element.rPr.rFonts.set(qn('w:eastAsia'), 'PingFang SC')

    doc.add_paragraph()  # 表后间距
    return table

# =============================================
# 开始构建文档
# =============================================

# ---- 封面区域 ----
# 空行制造视觉空间
for _ in range(2):
    doc.add_paragraph()

add_para('SpendLens', bold=True, size=Pt(28),
         color=RGBColor(0x5C, 0x2D, 0x3E), alignment=WD_ALIGN_PARAGRAPH.CENTER)
add_para('竞 品 调 研 报 告', bold=True, size=Pt(22),
         color=RGBColor(0xFF, 0x6B, 0x8A), alignment=WD_ALIGN_PARAGRAPH.CENTER)

doc.add_paragraph()
add_para('━━━━━━━━━━━━━━━━━━━━━━━━━━━', size=Pt(11),
         color=RGBColor(0xFF, 0xB3, 0xC6), alignment=WD_ALIGN_PARAGRAPH.CENTER)

add_para('调研日期：2026年6月9日', size=Pt(12),
         color=RGBColor(0xC4, 0x90, 0x9E), alignment=WD_ALIGN_PARAGRAPH.CENTER)
add_para('调研范围：国内6款 + 国外7款记账/理财工具', size=Pt(12),
         color=RGBColor(0xC4, 0x90, 0x9E), alignment=WD_ALIGN_PARAGRAPH.CENTER)
add_para('项目：SpendLens (bill-analyzer)', size=Pt(12),
         color=RGBColor(0xC4, 0x90, 0x9E), alignment=WD_ALIGN_PARAGRAPH.CENTER)

doc.add_page_break()

# ---- 目录 ----
add_heading_styled('目  录', level=1)
doc.add_paragraph()

toc_items = [
    ('一', '调研概览', '调研范围与方法'),
    ('二', '工具详情', '12款工具的详细分析'),
    ('三', '综合对比矩阵', '功能特性横向对比'),
    ('四', '核心洞察', '市场空白与机会'),
    ('五', '数据模型参考', '标准化数据结构设计'),
    ('六', '项目建议', '近期优化与长期规划'),
]
for num, title, desc in toc_items:
    p = doc.add_paragraph()
    run = p.add_run(f'{num}. {title}')
    run.font.bold = True
    run.font.size = Pt(14)
    run.font.name = 'PingFang SC'
    run.element.rPr.rFonts.set(qn('w:eastAsia'), 'PingFang SC')
    run.font.color.rgb = RGBColor(0x5C, 0x2D, 0x3E)
    run2 = p.add_run(f'  — {desc}')
    run2.font.size = Pt(11)
    run2.font.name = 'PingFang SC'
    run2.element.rPr.rFonts.set(qn('w:eastAsia'), 'PingFang SC')
    run2.font.color.rgb = RGBColor(0xC4, 0x90, 0x9E)
    set_cn_spacing(p, after=4, line_spacing=1.6)

doc.add_page_break()

# ---- 一、调研概览 ----
add_heading_styled('一、调研概览', level=1)
add_para('本报告对国内外12款主流记账及个人理财工具进行了系统调研，重点关注数据导入方式、分类体系设计、报告生成能力（特别是PPT/PDF导出）以及定价模式。', size=Pt(11))
add_para('调研核心问题：市场上是否存在 "上传Excel → 自动分析 → 生成PPT报告" 的同类产品？', bold=True, size=Pt(11), color=RGBColor(0xFF, 0x6B, 0x8A))

add_heading_styled('调研范围', level=2)

add_table_with_data(
    ['维度', '覆盖'],
    [
        ['国内市场', '鲨鱼记账、随手记、挖财、iCost、支付宝账单、微信记账'],
        ['国外市场', 'YNAB、Monarch Money、Tiller Money、Copilot Money、Spendee、Wallet by BudgetBakers、Lunch Money、PocketSmith'],
        ['关注点', '数据导入格式、支出分类体系、报告生成、PPT/PDF导出、UI设计、定价'],
    ]
)

# ---- 二、工具详情（摘要版） ----
add_heading_styled('二、工具详情', level=1)

tools = [
    {
        'name': '2.1 鲨鱼记账（Shark Accounting）',
        'platform': 'iOS / Android',
        'price': '免费 + 会员 ¥98/年',
        'features': '极简3秒记账，15个一级分类+40个二级分类，月度统计图表，多账本管理',
        'export': 'CSV / Excel（会员）',
        'ppt': '❌',
        'pros': '交互流畅、UI简洁、分类接地气',
        'cons': '不支持PPT/PDF报告、免费版导出受限、分析维度有限',
    },
    {
        'name': '2.2 随手记（Suishouji）— 金蝶旗下',
        'platform': 'iOS / Android / Web',
        'price': '免费 + 会员 ¥128/年',
        'features': '专业级记账(10年+)，丰富账本模板，预算+信贷+投资管理，OCR识别',
        'export': 'Excel / CSV',
        'ppt': '❌',
        'pros': '功能最全面、Web+App同步、金蝶技术背景',
        'cons': '学习成本高、UI老旧、无PPT导出、广告多',
    },
    {
        'name': '2.3 挖财（Wacai）',
        'platform': 'iOS / Android',
        'price': '免费 + 会员',
        'features': '记账+理财社区，部分银行自动导入，语音记账',
        'export': 'Excel',
        'ppt': '❌',
        'pros': '记账+理财一体化、语音记账',
        'cons': '理财导向过重、广告多、用户活跃度下降',
    },
    {
        'name': '2.4 iCost（用户数据源）',
        'platform': 'iOS',
        'price': '买断制 ¥68（终身）',
        'features': '极简记账、多账本、日历热力图、Excel导出含完整字段',
        'export': 'Excel（.xlsx）',
        'ppt': '❌',
        'pros': '导出格式规范、买断制、无广告、极简设计',
        'cons': '仅iOS、无自动分析报告、不支持银行同步',
    },
    {
        'name': '2.5 YNAB（You Need A Budget）',
        'platform': 'Web / iOS / Android',
        'price': '$14.99/月 或 $109/年',
        'features': '零基预算法，银行自动同步(12000+机构)，净值报告，债务管理',
        'export': 'CSV',
        'ppt': '❌',
        'pros': '方法论驱动、自动同步广、报告功能强、教程丰富',
        'cons': '学习曲线陡、不支持国区银行、无PPT/PDF导出、价格高',
    },
    {
        'name': '2.6 Monarch Money',
        'platform': 'Web / iOS / Android',
        'price': '$14.99/月 或 $99.99/年',
        'features': 'Mint替代品，银行/投资同步，交易自动分类+规则引擎，家庭协作',
        'export': 'CSV',
        'ppt': '❌',
        'pros': 'UI现代精美(卡片式)、规则引擎强、家庭成员协作、投资追踪',
        'cons': '不支持PPT/PDF报告、价格高、不支持中国银行',
    },
    {
        'name': '2.7 Tiller Money ⭐',
        'platform': 'Google Sheets / Excel',
        'price': '$79/年',
        'features': '银行交易→Google Sheets/Excel自动同步，预制模板，社区模板库，完全自定义',
        'export': 'Sheets/Excel（原生）',
        'ppt': '❌（但可手动制作）',
        'pros': '最灵活、理念与bill-analyzer最接近、社区模板丰富、隐私友好',
        'cons': '需表格操作能力、不支持中国银行、无移动端App',
    },
    {
        'name': '2.8 Copilot Money',
        'platform': 'iOS / Mac',
        'price': '$13/月 或 $95/年',
        'features': 'Apple原生SwiftUI，AI交易分类，订阅追踪，小组件',
        'export': '无',
        'ppt': '❌',
        'pros': 'UI设计最佳、AI分类准确、订阅追踪独特、隐私优先',
        'cons': '仅Apple设备、不支持任何导出、价格高',
    },
    {
        'name': '2.9 Spendee',
        'platform': 'iOS / Android / Web',
        'price': '免费 + Premium $2.99/月',
        'features': '多钱包、共享钱包、自动银行同步、预算管理、精美图表',
        'export': 'CSV、PDF（Premium）',
        'ppt': '❌',
        'pros': '支持PDF导出、共享钱包成熟、图表精美、价格低',
        'cons': 'PDF模板不可自定义、免费版受限',
    },
    {
        'name': '2.10 Wallet by BudgetBakers',
        'platform': 'iOS / Android / Web',
        'price': '免费 + Premium $3.49/月',
        'features': '银行同步(80+国)、多币种、预算+储蓄目标、详细报告',
        'export': 'CSV、Excel、PDF',
        'ppt': '❌',
        'pros': '多格式导出(CSV/Excel/PDF)、国际化好、Web端完整',
        'cons': '不支持PPT导出、中文不完善',
    },
    {
        'name': '2.11 Lunch Money',
        'platform': 'Web',
        'price': '$10/月',
        'features': '纯Web响应式，多币种，最强规则引擎，开放API，标签系统',
        'export': 'CSV、JSON',
        'ppt': '❌',
        'pros': '开放API、规则引擎最强、独立开发者(更新快)',
        'cons': '无移动端、无银行自动同步、不支持可视化导出',
    },
    {
        'name': '2.12 PocketSmith',
        'platform': 'Web / iOS / Android',
        'price': '免费 + Premium $9.95/月',
        'features': '日历视图、财务预测(最长30年)、场景模拟、净值追踪',
        'export': 'CSV',
        'ppt': '❌',
        'pros': '日历视图独树一帜、财务预测强、What-if分析',
        'cons': '不支持报告导出、界面复杂、学习成本高',
    },
]

for t in tools:
    add_heading_styled(t['name'], level=2)
    p = doc.add_paragraph()
    run = p.add_run(f"📱 平台：{t['platform']}    |    💰 定价：{t['price']}")
    run.font.size = Pt(10)
    run.font.name = 'PingFang SC'
    run.element.rPr.rFonts.set(qn('w:eastAsia'), 'PingFang SC')
    run.font.color.rgb = RGBColor(0xC4, 0x90, 0x9E)

    for label, val in [('✨ 核心功能', t['features']), ('📤 导出格式', t['export']),
                        ('🎯 PPT导出', t['ppt'])]:
        p = doc.add_paragraph()
        run_label = p.add_run(f'{label}：')
        run_label.font.bold = True
        run_label.font.size = Pt(10.5)
        run_label.font.name = 'PingFang SC'
        run_label.element.rPr.rFonts.set(qn('w:eastAsia'), 'PingFang SC')

        run_val = p.add_run(val)
        run_val.font.size = Pt(10.5)
        run_val.font.name = 'PingFang SC'
        run_val.element.rPr.rFonts.set(qn('w:eastAsia'), 'PingFang SC')

    p = doc.add_paragraph()
    run_pro = p.add_run(f'✅ 优点：{t["pros"]}')
    run_pro.font.size = Pt(10.5)
    run_pro.font.name = 'PingFang SC'
    run_pro.element.rPr.rFonts.set(qn('w:eastAsia'), 'PingFang SC')
    run_pro.font.color.rgb = RGBColor(0x7B, 0xC8, 0xA4)

    p = doc.add_paragraph()
    run_con = p.add_run(f'⚠️  不足：{t["cons"]}')
    run_con.font.size = Pt(10.5)
    run_con.font.name = 'PingFang SC'
    run_con.element.rPr.rFonts.set(qn('w:eastAsia'), 'PingFang SC')
    run_con.font.color.rgb = RGBColor(0xE8, 0x81, 0x5F)

    doc.add_paragraph()  # 间距

doc.add_page_break()

# ---- 三、综合对比矩阵 ----
add_heading_styled('三、综合对比矩阵', level=1)
add_para('以下矩阵横向对比所有工具的核心功能特性。✅=支持，❌=不支持，部分=有限支持。')

add_table_with_data(
    ['工具', 'Excel导入', 'PPT导出', 'PDF导出', '自动同步', '开放API', '中文支持', '家庭协作'],
    [
        ['鲨鱼记账', '✅', '❌', '❌', '❌', '❌', '✅', '❌'],
        ['随手记', '✅', '❌', '❌', '❌', '❌', '✅', '✅'],
        ['挖财', '✅', '❌', '❌', '部分', '❌', '✅', '✅'],
        ['iCost', '✅', '❌', '❌', '❌', '❌', '✅', '❌'],
        ['YNAB', '❌', '❌', '❌', '✅', '✅', '❌', '✅'],
        ['Monarch', '❌', '❌', '❌', '✅', '❌', '❌', '✅'],
        ['Tiller', '✅', '❌', '❌', '✅', '✅', '❌', '❌'],
        ['Copilot', '❌', '❌', '❌', '✅', '❌', '❌', '❌'],
        ['Spendee', '❌', '❌', '✅', '✅', '❌', '❌', '✅'],
        ['Wallet', '✅', '❌', '✅', '✅', '❌', '部分', '✅'],
        ['LunchMoney', '✅', '❌', '❌', '❌', '✅', '❌', '❌'],
        ['PocketSmith', '❌', '❌', '❌', '✅', '❌', '❌', '❌'],
        ['bill-analyzer', '✅', '✅', '❌', '❌', '❌', '✅', '❌'],
    ]
)

add_para('🔑 关键发现：全市场没有任何工具能 "导入Excel → 自动生成PPT报告"。bill-analyzer 是唯一的。',
         bold=True, size=Pt(12), color=RGBColor(0xFF, 0x6B, 0x8A))

doc.add_page_break()

# ---- 四、核心洞察 ----
add_heading_styled('四、核心洞察', level=1)

add_heading_styled('4.1 市场空白', level=2)
add_para('"Excel导入 + 自动生成PPT报告" 是一个明确的市场空白。现有工具分为两类：')
p = doc.add_paragraph()
run = p.add_run('① 记账App')
run.font.bold = True
run.font.size = Pt(11)
run.font.name = 'PingFang SC'
run.element.rPr.rFonts.set(qn('w:eastAsia'), 'PingFang SC')
run2 = p.add_run('（鲨鱼、随手记、YNAB等）：有自己的数据录入渠道，不依赖Excel导入，也不生成PPT。')
run2.font.size = Pt(11)
run2.font.name = 'PingFang SC'
run2.element.rPr.rFonts.set(qn('w:eastAsia'), 'PingFang SC')

p = doc.add_paragraph()
run = p.add_run('② 表格工具')
run.font.bold = True
run.font.size = Pt(11)
run.font.name = 'PingFang SC'
run.element.rPr.rFonts.set(qn('w:eastAsia'), 'PingFang SC')
run2 = p.add_run('（Tiller、Lunch Money等）：依赖电子表格，但不生成PPT。')
run2.font.size = Pt(11)
run2.font.name = 'PingFang SC'
run2.element.rPr.rFonts.set(qn('w:eastAsia'), 'PingFang SC')

add_para('bill-analyzer 的独特定位：连接 iCost/银行导出Excel → 自动分析 → 专业PPT报告。', bold=True, size=Pt(12))

add_heading_styled('4.2 分类体系最佳实践', level=2)
add_para('从YNAB的Category Groups → Categories思路，结合中国家庭消费场景，建议分类结构：')

add_table_with_data(
    ['一级分类', '二级分类示例', '来源参考'],
    [
        ['餐饮', '三餐、外卖、饮料、零食、水果、聚餐', '鲨鱼记账 + iCost'],
        ['住房', '房租/房贷、水电、物业、维修、家居', '随手记'],
        ['交通', '公交、打车、加油、停车、保养', '通用'],
        ['购物', '日用品、数码、家居、宠物', '通用'],
        ['娱乐', '电影、游戏、旅行、运动', '通用'],
        ['人情', '礼金、红包、请客', 'iCost特色'],
        ['医疗', '看病、药品、体检、保险', '通用'],
        ['教育', '培训、书籍、文具', '通用'],
        ['通讯', '话费、宽带', '通用'],
        ['服饰', '衣服、鞋子、配饰', 'iCost特色'],
    ]
)

add_para('关键设计原则（参考YNAB + Monarch）：')
p = doc.add_paragraph()
run = p.add_run('● 使用分组（Groups）→ 分类（Categories）层级，而非平铺列表')
run.font.size = Pt(11)
run.font.name = 'PingFang SC'
run.element.rPr.rFonts.set(qn('w:eastAsia'), 'PingFang SC')

p = doc.add_paragraph()
run = p.add_run('● 把"转账（Transfer）"独立于收支之外，避免重复计算（Monarch的做法）')
run.font.size = Pt(11)
run.font.name = 'PingFang SC'
run.element.rPr.rFonts.set(qn('w:eastAsia'), 'PingFang SC')

p = doc.add_paragraph()
run = p.add_run('● 支持自定义标签（Lunch Money风格），不限于固定分类')
run.font.size = Pt(11)
run.font.name = 'PingFang SC'
run.element.rPr.rFonts.set(qn('w:eastAsia'), 'PingFang SC')

add_heading_styled('4.3 报告生成差异化', level=2)

add_table_with_data(
    ['功能', '市场现状', 'bill-analyzer 优势'],
    [
        ['饼图', '几乎所有工具都有', '手绘xkcd风格，更生动有趣'],
        ['柱状图', '主流工具都有', '结合卡通元素 + 自定义配色'],
        ['趋势线', '大部分有', '每周维度 + 峰值标注'],
        ['PPT导出', '全市场无 ❌', '✅ 核心差异化优势'],
        ['PDF导出', '仅Spendee/Wallet', '可考虑加入作为轻量替代'],
    ]
)

add_heading_styled('4.4 值得加入的功能（按优先级）', level=2)

add_para('🔴 高优先级（可直接借鉴）：', bold=True, size=Pt(12))
items_high = [
    '日历热力图：像PocketSmith/iCost一样展示每日支出（基础数据已有）',
    '同比/环比对比：本月 vs 上月、本月 vs 去年同月',
    '预算设置 + 实际vs预算对比：YNAB风格',
    '自定义规则引擎：自动根据商户名/金额给交易打标签',
]
for item in items_high:
    p = doc.add_paragraph(f'  ✓ {item}')
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.space_before = Pt(0)

doc.add_paragraph()
add_para('🟡 中优先级（考虑加入）：', bold=True, size=Pt(12))
items_mid = [
    'What-if 场景模拟：如"停止外卖每月能省多少"',
    'PDF报告导出：作为PPT的轻量替代',
    '分享/协作功能：生成分享链接或邀请家人查看',
    '财务健康评分：综合储蓄率、支出结构、变动趋势给出0-100分',
]
for item in items_mid:
    p = doc.add_paragraph(f'  ○ {item}')
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.space_before = Pt(0)

doc.add_paragraph()
add_para('🟢 低优先级（长期规划）：', bold=True, size=Pt(12))
items_low = [
    '银行/支付宝自动同步：通过开放银行API',
    'AI智能建议：基于消费模式生成个性化省钱建议',
    'Web端 + 移动端适配',
]
for item in items_low:
    p = doc.add_paragraph(f'  · {item}')
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.space_before = Pt(0)

doc.add_page_break()

# ---- 五、数据模型参考 ----
add_heading_styled('五、数据模型参考', level=1)
add_para('综合调研成果，建议 bill-analyzer 内部使用以下标准化数据模型（JSON-like表示）：')

add_heading_styled('5.1 交易（Transaction）', level=2)
code_block = """{
  "date": "2026-05-15",
  "type": "expense",           // expense | income | transfer
  "amount": 45.00,
  "currency": "CNY",
  "category_group": "餐饮",   // 一级分类（参考YNAB Group）
  "category": "外卖",         // 二级分类（参考YNAB Category）
  "account_from": "支付宝",   // 资金来源账户
  "account_to": null,         // 转账目标（transfer类型时使用）
  "merchant": "美团外卖",     // 商户名
  "note": "午餐",             // 备注
  "tags": ["工作日", "午餐"], // 自定义标签（参考Lunch Money）
  "is_recurring": false,      // 是否定期支出
  "source": "icost_import"    // 数据来源
}"""
p = doc.add_paragraph()
run = p.add_run(code_block)
run.font.name = 'Menlo'
run.font.size = Pt(9)
run.font.color.rgb = RGBColor(0x3D, 0x1E, 0x2A)
set_cn_spacing(p, after=6, line_spacing=1.2)

add_heading_styled('5.2 分类（Category）', level=2)
code_block2 = """{
  "name": "外卖",
  "group": "餐饮",
  "type": "expense",
  "icon": "🍔",
  "color": "#FF6B8A",
  "budget_monthly": 500       // 月度预算
}"""
p = doc.add_paragraph()
run = p.add_run(code_block2)
run.font.name = 'Menlo'
run.font.size = Pt(9)
run.font.color.rgb = RGBColor(0x3D, 0x1E, 0x2A)
set_cn_spacing(p, after=6, line_spacing=1.2)

add_heading_styled('5.3 月度报告（MonthlyReport）', level=2)
code_block3 = """{
  "month": "2026-05",
  "total_income": 30636.32,
  "total_expense": 9673.78,
  "balance": 20962.54,
  "savings_rate": 68.4,
  "transaction_count": 167,
  "categories": [...],
  "weekly_trend": [...],
  "account_distribution": [...],
  "highlights": ["储蓄率创新高", "餐饮支出下降12%"],
  "suggestions": ["继续保持当前消费习惯", "餐饮中做饭比例高，很健康"]
}"""
p = doc.add_paragraph()
run = p.add_run(code_block3)
run.font.name = 'Menlo'
run.font.size = Pt(9)
run.font.color.rgb = RGBColor(0x3D, 0x1E, 0x2A)
set_cn_spacing(p, after=6, line_spacing=1.2)

doc.add_page_break()

# ---- 六、项目建议 ----
add_heading_styled('六、对 bill-analyzer 项目的建议', level=1)

add_heading_styled('6.1 当前优势', level=2)
advantages = [
    ('✅ 市场空白', '唯一做到 Excel→分析→PPT 全链路的工具'),
    ('✅ 手绘风格', 'xkcd图表在数据分析领域独树一帜，有辨识度'),
    ('✅ iOS设计', '前端UI和PPT都采用了高级设计语言，品质感强'),
    ('✅ 极简部署', '单目录5文件，Flask单进程，零门槛运行'),
]
for title, desc in advantages:
    p = doc.add_paragraph()
    run = p.add_run(f'{title}：')
    run.font.bold = True
    run.font.size = Pt(11)
    run.font.name = 'PingFang SC'
    run.element.rPr.rFonts.set(qn('w:eastAsia'), 'PingFang SC')
    run.font.color.rgb = RGBColor(0x7B, 0xC8, 0xA4)
    run2 = p.add_run(desc)
    run2.font.size = Pt(11)
    run2.font.name = 'PingFang SC'
    run2.element.rPr.rFonts.set(qn('w:eastAsia'), 'PingFang SC')

add_heading_styled('6.2 近期优化方向', level=2)
improvements = [
    ('📊 环比对比', '本月 vs 上月自动对比，发现消费变化趋势'),
    ('📅 日历热力图', '参考PocketSmith/iCost，直观展示每日支出'),
    ('🏷️ 标签系统', '参考Lunch Money，支持自定义标签而不限于固定分类'),
    ('📤 PDF导出', '参考Spendee/Wallet，作为PPT的轻量替代方案'),
    ('🔄 多文件对比', '上传两个月的账单 → 自动对比分析'),
]
for title, desc in improvements:
    p = doc.add_paragraph()
    run = p.add_run(f'{title}：')
    run.font.bold = True
    run.font.size = Pt(11)
    run.font.name = 'PingFang SC'
    run.element.rPr.rFonts.set(qn('w:eastAsia'), 'PingFang SC')
    run2 = p.add_run(desc)
    run2.font.size = Pt(11)
    run2.font.name = 'PingFang SC'
    run2.element.rPr.rFonts.set(qn('w:eastAsia'), 'PingFang SC')

add_heading_styled('6.3 长期规划', level=2)
longterm = [
    ('🌐 线上部署', '部署为SaaS服务，用户无需本地安装'),
    ('🤖 AI分析', '基于消费模式自动生成个性化文字建议'),
    ('📱 PWA适配', '手机上也能上传账单、查看报告'),
    ('🔗 开放API', '让其他工具能对接，构建生态'),
]
for title, desc in longterm:
    p = doc.add_paragraph()
    run = p.add_run(f'{title}：')
    run.font.bold = True
    run.font.size = Pt(11)
    run.font.name = 'PingFang SC'
    run.element.rPr.rFonts.set(qn('w:eastAsia'), 'PingFang SC')
    run2 = p.add_run(desc)
    run2.font.size = Pt(11)
    run2.font.name = 'PingFang SC'
    run2.element.rPr.rFonts.set(qn('w:eastAsia'), 'PingFang SC')

doc.add_paragraph()
doc.add_paragraph()

# ---- 尾页 ----
add_para('━━━━━━━━━━━━━━━━━━━━━━━━━━━', size=Pt(11),
         color=RGBColor(0xFF, 0xB3, 0xC6), alignment=WD_ALIGN_PARAGRAPH.CENTER)
doc.add_paragraph()
add_para('本报告由 bill-analyzer 项目组调研撰写', size=Pt(11),
         color=RGBColor(0xC4, 0x90, 0x9E), alignment=WD_ALIGN_PARAGRAPH.CENTER)
add_para('数据来源：各产品官网、公开文档、实际使用体验', size=Pt(10),
         color=RGBColor(0xC4, 0x90, 0x9E), alignment=WD_ALIGN_PARAGRAPH.CENTER)
add_para('调研日期：2026-06-09', size=Pt(10),
         color=RGBColor(0xC4, 0x90, 0x9E), alignment=WD_ALIGN_PARAGRAPH.CENTER)

# ---- 保存 ----
output_path = '竞品调研报告.docx'
doc.save(output_path)
print(f'✅ 已生成：{output_path}')
