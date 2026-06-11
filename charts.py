"""xkcd 手绘风格图表生成模块。所有图表输出为 BytesIO，不落盘。"""

import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from io import BytesIO

# ---- 字体配置 ----
import glob

def _find_cn_font():
    base = os.path.dirname(os.path.abspath(__file__))
    candidates = (
        # 项目自带字体（最优先，保证跨平台一致性）
        glob.glob(os.path.join(base, 'fonts', 'wqy-microhei.ttc')) +
        glob.glob(os.path.join(base, 'fonts', '*.ttc')) +
        glob.glob(os.path.join(base, 'fonts', '*.ttf')) +
        # macOS
        glob.glob('/System/Library/AssetsV2/com_apple_MobileAsset_Font8/*/AssetData/Kaiti.ttc') +
        glob.glob('/System/Library/Fonts/PingFang*') +
        glob.glob('/System/Library/Fonts/STHeiti*') +
        glob.glob('/Library/Fonts/*.ttf') +
        # Linux (Railway / Ubuntu)
        glob.glob('/usr/share/fonts/truetype/wqy/wqy-microhei*') +
        glob.glob('/usr/share/fonts/truetype/wqy/wqy-zenhei*') +
        glob.glob('/usr/share/fonts/truetype/noto/NotoSansCJK*') +
        glob.glob('/usr/share/fonts/truetype/noto/NotoSans*') +
        glob.glob('/usr/share/fonts/opentype/noto/NotoSansCJK*') +
        glob.glob('/usr/share/fonts/opentype/noto/NotoSans*') +
        glob.glob('/usr/share/fonts/noto-cjk/NotoSans*') +
        glob.glob('/usr/share/fonts/truetype/droid/DroidSansFallback*') +
        glob.glob('/usr/share/fonts/truetype/arphic/*') +
        glob.glob('/usr/share/fonts/wenquanyi/**/*.ttc') +
        glob.glob('/usr/share/fonts/wenquanyi/**/*.ttf') +
        # 兜底：扫描所有 ttc/ttf
        glob.glob('/usr/share/fonts/**/*.ttc') +
        glob.glob('/usr/share/fonts/**/*.ttf')
    )
    for fp in candidates:
        try:
            fm.FontProperties(fname=fp)
            return fp
        except:
            pass
    return None

CN_FONT_PATH = _find_cn_font()

# 强制刷新 matplotlib 字体缓存（解决 Railway 部署后字体找不到的问题）
if CN_FONT_PATH:
    try:
        fm._load_fontmanager(try_read_cache=False)
    except Exception:
        pass

plt.xkcd(scale=1, length=100, randomness=2)

# ---- 粉色卡通风色板 ----
CC = ['#FF6B8A', '#FF85A2', '#FF9EBB', '#FFB3C6', '#FFD4B8',
      '#FF7EB3', '#C4909E', '#E8A0B4', '#F4B8C8', '#FFCAD4']
DARK = '#3D1E2A'
GRAY = '#C4909E'
LIGHT_GRAY = '#FFD4DF'


def _fp(size=12):
    if CN_FONT_PATH:
        return fm.FontProperties(fname=CN_FONT_PATH, size=size)
    return None


def _set_cn(labels, fp_obj):
    """Apply Chinese font to tick labels."""
    for label in labels:
        if fp_obj:
            label.set_fontproperties(fp_obj)


def pie_spending(cat1_list):
    """支出结构环形图。cat1_list: [(name, amount), ...]"""
    fig, ax = plt.subplots(figsize=(8, 5))
    fig.patch.set_facecolor('none')

    labels = [c[0] for c in cat1_list]
    total = sum(c[1] for c in cat1_list)
    sizes = [c[1] / total * 100 for c in cat1_list]
    explode = [0.05 if i == 0 else 0 for i in range(len(cat1_list))]

    wedges, texts, autotexts = ax.pie(
        sizes, explode=explode, labels=None, colors=CC[:len(cat1_list)],
        autopct='%1.1f%%', startangle=140, pctdistance=0.78,
        wedgeprops={'width': 0.45, 'edgecolor': 'white', 'linewidth': 2},
    )

    for at in autotexts:
        at.set_fontsize(9)
        at.set_color(DARK)
        if CN_FONT_PATH:
            at.set_fontproperties(_fp(9))

    legend_labels = [f'{l}  {s:.1f}%' for l, s in zip(labels, sizes)]
    ax.legend(wedges, legend_labels, loc='center left',
              bbox_to_anchor=(1.02, 0.5), prop=_fp(9), frameon=False)

    ax.set_title('支出结构分布', fontsize=18, color=DARK, pad=20,
                 fontproperties=_fp(18))
    plt.tight_layout()

    buf = BytesIO()
    fig.savefig(buf, dpi=200, format='png', bbox_inches='tight', transparent=True)
    plt.close(fig)
    buf.seek(0)
    return buf


def bar_food(food_list, daily_food):
    """餐饮细分柱状图。food_list: [(name, amount), ...]"""
    fig, ax = plt.subplots(figsize=(7, 4.5))
    fig.patch.set_facecolor('none')
    ax.set_facecolor('none')

    cats = [c[0] for c in food_list]
    vals = [c[1] for c in food_list]
    bc = ['#E8815F', '#D4A853', '#5B9A8B', '#B0A090'][:len(cats)]

    bars = ax.bar(cats, vals, color=bc, width=0.5, edgecolor=DARK, linewidth=1.5, alpha=0.9)
    for bar, val in zip(bars, vals):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + max(vals) * 0.02,
                f'¥{val:.0f}', ha='center', fontsize=12, color=DARK, fontproperties=_fp(12))

    ax.set_title('餐饮细分', fontsize=18, color=DARK, pad=15, fontproperties=_fp(18))
    ax.set_ylabel('金额 (¥)', fontsize=11, color=GRAY, fontproperties=_fp(11))
    ax.tick_params(colors=GRAY, labelsize=10)
    _set_cn(ax.get_xticklabels(), _fp(10))
    _set_cn(ax.get_yticklabels(), _fp(10))
    for spine in ['top', 'right']:
        ax.spines[spine].set_visible(False)
    for spine in ['left', 'bottom']:
        ax.spines[spine].set_color(LIGHT_GRAY)
    ax.set_ylim(0, max(vals) * 1.25)
    plt.tight_layout()

    buf = BytesIO()
    fig.savefig(buf, dpi=200, format='png', bbox_inches='tight', transparent=True)
    plt.close(fig)
    buf.seek(0)
    return buf


def line_weekly(weekly_list):
    """每周支出趋势折线图。weekly_list: [(week_label, amount), ...]"""
    fig, ax = plt.subplots(figsize=(7, 4.5))
    fig.patch.set_facecolor('none')
    ax.set_facecolor('none')

    weeks = [w[0] for w in weekly_list]
    wvals = [w[1] for w in weekly_list]

    ax.plot(weeks, wvals, color='#E8815F', linewidth=2.5, marker='o',
            markersize=10, markerfacecolor='white', markeredgewidth=2,
            markeredgecolor='#E8815F')
    ax.fill_between(range(len(weeks)), wvals, alpha=0.15, color='#E8815F')

    for x, y in zip(weeks, wvals):
        ax.annotate(f'¥{y:,.0f}', (x, y), textcoords="offset points",
                    xytext=(0, 15), ha='center', fontsize=10, color=DARK,
                    fontproperties=_fp(10))

    ax.set_title('每周支出趋势', fontsize=18, color=DARK, pad=15, fontproperties=_fp(18))
    ax.tick_params(colors=GRAY, labelsize=9)
    _set_cn(ax.get_xticklabels(), _fp(9))
    for spine in ['top', 'right']:
        ax.spines[spine].set_visible(False)
    for spine in ['left', 'bottom']:
        ax.spines[spine].set_color(LIGHT_GRAY)
    ax.set_ylim(0, max(wvals) * 1.35)
    plt.tight_layout()

    buf = BytesIO()
    fig.savefig(buf, dpi=200, format='png', bbox_inches='tight', transparent=True)
    plt.close(fig)
    buf.seek(0)
    return buf


def bar_account(account_list):
    """账户支出横向条形图。account_list: [(name, amount), ...]"""
    total = sum(a[1] for a in account_list)
    fig, ax = plt.subplots(figsize=(7, 3.5))
    fig.patch.set_facecolor('none')
    ax.set_facecolor('none')

    accs = [a[0] for a in account_list]
    pcts = [a[1] / total * 100 for a in account_list]
    amts = [a[1] for a in account_list]
    bcols = ['#E8815F', '#5B9A8B', '#D4A853', '#B0A090'][:len(accs)]

    bars = ax.barh(accs, pcts, color=bcols, height=0.45, edgecolor=DARK, linewidth=1.5)
    for bar, pct, amt in zip(bars, pcts, amts):
        ax.text(bar.get_width() + 1, bar.get_y() + bar.get_height() / 2,
                f'{pct:.1f}%  ¥{amt:,.0f}', va='center', fontsize=11,
                color=DARK, fontproperties=_fp(11))

    ax.set_title('各账户支出占比', fontsize=18, color=DARK, pad=15, fontproperties=_fp(18))
    ax.tick_params(colors=GRAY, labelsize=10)
    _set_cn(ax.get_yticklabels(), _fp(10))
    for spine in ['top', 'right']:
        ax.spines[spine].set_visible(False)
    for spine in ['left', 'bottom']:
        ax.spines[spine].set_color(LIGHT_GRAY)
    ax.set_xlim(0, max(pcts) * 1.2)
    plt.tight_layout()

    buf = BytesIO()
    fig.savefig(buf, dpi=200, format='png', bbox_inches='tight', transparent=True)
    plt.close(fig)
    buf.seek(0)
    return buf


def doughnut_income(income_list):
    """收入来源环形图。income_list: [(name, amount), ...]"""
    total = sum(i[1] for i in income_list)
    fig, ax = plt.subplots(figsize=(7, 5.5))
    fig.patch.set_facecolor('none')

    ilabs = [i[0] for i in income_list]
    isizes = [i[1] / total * 100 for i in income_list]
    iexplode = [0.05 if i == 0 else 0 for i in range(len(income_list))]
    ic = ['#5B9A8B', '#E8815F', '#D4A853', '#B0A090'][:len(ilabs)]

    wedges, texts, autotexts = ax.pie(
        isizes, explode=iexplode, labels=None, colors=ic,
        autopct='%1.1f%%', startangle=90, pctdistance=0.78,
        wedgeprops={'width': 0.4, 'edgecolor': 'white', 'linewidth': 2},
    )
    for at in autotexts:
        at.set_fontsize(10)
        at.set_color(DARK)
        if CN_FONT_PATH:
            at.set_fontproperties(_fp(10))

    leg_labels = [f'{l}  {s:.1f}%' for l, s in zip(ilabs, isizes)]
    ax.legend(wedges, leg_labels, loc='upper center',
              bbox_to_anchor=(0.5, -0.08), prop=_fp(9), frameon=False, ncol=3)
    ax.set_title('收入来源分布', fontsize=18, color=DARK, pad=20, fontproperties=_fp(18))
    plt.tight_layout()

    buf = BytesIO()
    fig.savefig(buf, dpi=200, format='png', bbox_inches='tight', transparent=True)
    plt.close(fig)
    buf.seek(0)
    return buf


def all_charts(data):
    """一键生成全部 5 张图表，返回 {name: BytesIO}"""
    return {
        'pie_spending': pie_spending(data['cat1_list']),
        'bar_food': bar_food(data['food_list'], data['daily_food']),
        'line_weekly': line_weekly(data['weekly_list']),
        'bar_account': bar_account(data['account_list']),
        'doughnut_income': doughnut_income(data['income_list']),
    }


# ================================================================
#  新增图表
# ================================================================

import matplotlib.patches as mpatches
import numpy as np


def heatmap_daily(daily_list):
    """日历热力图 — 按周展示每日支出强度。

    daily_list: [(mm-dd, amount), ...] 长度通常28-31
    """
    if not daily_list:
        return None

    fig, ax = plt.subplots(figsize=(8, 5))
    fig.patch.set_facecolor('none')
    ax.set_facecolor('none')

    # 解析日期，构建网格
    days_in_month = len(daily_list)

    # 找到第一天是周几 (假设从 month-day 反向推导)
    import calendar as cal_mod
    first_label = daily_list[0][0]  # mm-dd
    try:
        first_month = int(first_label.split('-')[0])
        first_day = int(first_label.split('-')[1])
    except (ValueError, IndexError):
        first_month, first_day = 1, 1

    # 假定年份是当前年份或从数据推算
    year = 2026
    start_weekday = cal_mod.weekday(year, first_month, first_day)  # 0=Mon

    # 构建7列（周一至周日）的热力数据
    amounts = [d[1] for d in daily_list]
    max_amt = max(amounts) if max(amounts) > 0 else 1

    # 填充网格
    n_rows = (start_weekday + days_in_month + 6) // 7
    grid = np.full((n_rows, 7), np.nan)
    day_idx = 0
    for r in range(n_rows):
        for c in range(7):
            if r == 0 and c < start_weekday:
                continue
            if day_idx >= days_in_month:
                break
            grid[r, c] = amounts[day_idx]
            day_idx += 1

    # 粉色渐变热力
    from matplotlib.colors import LinearSegmentedColormap
    cmap = LinearSegmentedColormap.from_list('pink_heat',
        ['#FFFFFF', '#FFE4EC', '#FFB3C6', '#FF85A2', '#FF6B8A', '#FF3D6A'], N=100)

    im = ax.imshow(grid, cmap=cmap, aspect='equal', vmin=0, vmax=max_amt)

    # 标注金额和日期
    day_labels = [d[0].split('-')[1] if '-' in d[0] else d[0] for d in daily_list]
    amt_labels = [f'¥{d[1]:.0f}' if d[1] > 0 else '' for d in daily_list]

    idx = 0
    for r in range(n_rows):
        for c in range(7):
            if not np.isnan(grid[r, c]):
                amt = grid[r, c]
                # 深色背景用白色文字 + 阴影增强可读性
                text_color = '#FFFFFF' if amt > max_amt * 0.35 else DARK
                ax.text(c, r, f'{day_labels[idx]}\n{amt_labels[idx]}',
                        ha='center', va='center', fontsize=7, color=text_color,
                        fontproperties=_fp(7),
                        bbox=dict(boxstyle='round,pad=0.1', facecolor='black', alpha=0.25, edgecolor='none') if amt > max_amt * 0.35 else dict(boxstyle='round,pad=0.1', facecolor='white', alpha=0.55, edgecolor='none'))
                idx += 1

    # 坐标轴
    weekdays = ['周一', '周二', '周三', '周四', '周五', '周六', '周日']
    ax.set_xticks(range(7))
    ax.set_xticklabels(weekdays, fontsize=9, color=GRAY)
    _set_cn(ax.get_xticklabels(), _fp(9))
    ax.set_yticks(range(n_rows))
    ax.set_yticklabels([f'W{w+1}' for w in range(n_rows)], fontsize=9, color=GRAY)

    ax.set_title('每日支出热力图', fontsize=18, color=DARK, pad=15, fontproperties=_fp(18))

    # 颜色条
    cbar = plt.colorbar(im, ax=ax, shrink=0.8, pad=0.02)
    cbar.set_label('支出金额 (¥)', fontsize=9, color=GRAY)
    cbar.ax.tick_params(labelsize=8, colors=GRAY)

    for spine in ax.spines.values():
        spine.set_visible(False)

    plt.tight_layout()

    buf = BytesIO()
    fig.savefig(buf, dpi=200, format='png', bbox_inches='tight', transparent=True)
    plt.close(fig)
    buf.seek(0)
    return buf


def bar_budget_vs_actual(cat1_list, budget_config):
    """预算 vs 实际对比柱状图。

    cat1_list: [(name, amount), ...]
    budget_config: {"monthly_total": 8000, "categories": {"餐饮": 2000, ...}}
    """
    fig, ax = plt.subplots(figsize=(7, 4.5))
    fig.patch.set_facecolor('none')
    ax.set_facecolor('none')

    cat_budgets = budget_config.get('categories', {})
    total_budget = budget_config.get('monthly_total', 0)

    # 筛选有预算的分类
    cats = []
    actuals = []
    budgets = []
    for cat_name, actual in cat1_list:
        if cat_name in cat_budgets:
            cats.append(cat_name)
            actuals.append(actual)
            budgets.append(cat_budgets[cat_name])

    if not cats:
        ax.text(0.5, 0.5, '暂无预算数据\n请在预算Tab中设置', transform=ax.transAxes,
                ha='center', va='center', fontsize=14, color=GRAY, fontproperties=_fp(14))
        buf = BytesIO()
        fig.savefig(buf, dpi=150, format='png', bbox_inches='tight', transparent=True)
        plt.close(fig)
        buf.seek(0)
        return buf

    x = np.arange(len(cats))
    w = 0.35

    bars1 = ax.bar(x - w/2, actuals, w, color='#FF6B8A', edgecolor=DARK, linewidth=1, label='实际支出')
    bars2 = ax.bar(x + w/2, budgets, w, color='#FFD4DF', edgecolor=DARK, linewidth=1, label='预算')

    for bar, val in zip(bars1, actuals):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + max(actuals)*0.02,
                f'¥{val:.0f}', ha='center', fontsize=8, color=DARK, fontproperties=_fp(8))

    ax.set_xticks(x)
    ax.set_xticklabels(cats, fontsize=10, color=DARK)
    _set_cn(ax.get_xticklabels(), _fp(10))
    ax.set_title('预算 vs 实际支出', fontsize=18, color=DARK, pad=15, fontproperties=_fp(18))
    ax.legend(prop=_fp(10), frameon=False)
    ax.tick_params(colors=GRAY, labelsize=9)
    for spine in ['top', 'right']:
        ax.spines[spine].set_visible(False)
    for spine in ['left', 'bottom']:
        ax.spines[spine].set_color(LIGHT_GRAY)
    if total_budget > 0:
        ax.axhline(y=total_budget, color='#FF3D6A', linestyle='--', linewidth=1.5, alpha=0.6,
                   label=f'总预算 ¥{total_budget:,.0f}')

    plt.tight_layout()

    buf = BytesIO()
    fig.savefig(buf, dpi=200, format='png', bbox_inches='tight', transparent=True)
    plt.close(fig)
    buf.seek(0)
    return buf


def bar_compare_monthly(compare_data):
    """环比对比聚类柱状图。

    compare_data: analyzer.compare() 的返回值
    """
    fig, ax = plt.subplots(figsize=(7, 5))
    fig.patch.set_facecolor('none')
    ax.set_facecolor('none')

    categories_data = compare_data.get('categories', [])
    # 取有实际数据的类别，按变化量排序
    cats_to_show = [c for c in categories_data if c['prev'] > 0 or c['curr'] > 0]
    cats_to_show.sort(key=lambda x: x['curr'], reverse=True)
    cats_to_show = cats_to_show[:8]  # 只显示前8个

    if not cats_to_show:
        buf = BytesIO()
        ax.text(0.5, 0.5, '无对比数据', transform=ax.transAxes,
                ha='center', va='center', fontsize=14, color=GRAY)
        fig.savefig(buf, dpi=150, format='png', bbox_inches='tight', transparent=True)
        plt.close(fig)
        buf.seek(0)
        return buf

    cats = [c['name'] for c in cats_to_show]
    x = np.arange(len(cats))
    w = 0.3

    prev_vals = [c['prev'] for c in cats_to_show]
    curr_vals = [c['curr'] for c in cats_to_show]

    bars1 = ax.bar(x - w/2, prev_vals, w, color='#FFD4DF', edgecolor=DARK, linewidth=1,
                   label=compare_data.get('prev_month', '上月'))
    bars2 = ax.bar(x + w/2, curr_vals, w, color='#FF6B8A', edgecolor=DARK, linewidth=1,
                   label=compare_data.get('curr_month', '本月'))

    # 标注变化
    for i, c in enumerate(cats_to_show):
        chg = c['change_pct']
        if chg is not None and chg != 0:
            arrow = '↑' if chg > 0 else '↓'
            color = '#FF6B8A' if chg > 0 else '#7BC8A4'
            ax.text(i, max(c['curr'], c['prev']) * 1.05,
                    f'{arrow}{abs(chg):.0f}%', ha='center', fontsize=8,
                    color=color, fontproperties=_fp(8))

    ax.set_xticks(x)
    ax.set_xticklabels(cats, fontsize=10, color=DARK)
    _set_cn(ax.get_xticklabels(), _fp(10))
    ax.set_title(f"{compare_data.get('prev_month','上月')} vs {compare_data.get('curr_month','本月')}",
                 fontsize=16, color=DARK, pad=15, fontproperties=_fp(16))
    ax.legend(prop=_fp(10), frameon=False)
    ax.tick_params(colors=GRAY, labelsize=9)
    for spine in ['top', 'right']:
        ax.spines[spine].set_visible(False)
    for spine in ['left', 'bottom']:
        ax.spines[spine].set_color(LIGHT_GRAY)

    plt.tight_layout()

    buf = BytesIO()
    fig.savefig(buf, dpi=200, format='png', bbox_inches='tight', transparent=True)
    plt.close(fig)
    buf.seek(0)
    return buf
