"""SpendLens — Flask 后端."""

import os
import sys
import tempfile
import hashlib
import json
from flask import Flask, request, jsonify, send_file, render_template
from flask_cors import CORS

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from analyzer import analyze, compare
from charts import all_charts, heatmap_daily, bar_budget_vs_actual, bar_compare_monthly
from ppt_builder import build_ppt
from budget import (
    health_score, load_budget, save_budget, compare_budget,
    load_rules, save_rules,
    simulate, create_share, load_share
)

app = Flask(__name__)
CORS(app)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB

# 分析结果缓存（避免图表接口重复分析）
_analysis_cache: dict[str, dict] = {}

BASE_URL = os.environ.get('BASE_URL', 'http://localhost:5050')


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/analyze', methods=['POST'])
def analyze_route():
    """上传 Excel → 返回分析 JSON + cache_id（供后续图表接口使用）。"""
    if 'file' not in request.files:
        return jsonify({'error': '未找到文件'}), 400

    file = request.files['file']
    if not file.filename.endswith(('.xlsx', '.xls')):
        return jsonify({'error': '请上传 .xlsx 或 .xls 文件'}), 400

    try:
        # 保存到临时文件
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            file.save(tmp.name)
            tmp_path = tmp.name

        # 生成 cache_id（文件内容哈希 + 时间戳）
        with open(tmp_path, 'rb') as f:
            file_hash = hashlib.md5(f.read()).hexdigest()[:12]

        data = analyze(tmp_path)
        hs = health_score(data)
        data['health'] = hs
        budget_cfg = load_budget()
        bc = compare_budget(data, budget_cfg)
        data['budget_comparison'] = bc
        os.unlink(tmp_path)

        # 缓存分析数据（供图表接口复用）
        cache_id = file_hash
        _analysis_cache[cache_id] = data

        # 清理旧缓存（保留最近 20 条）
        if len(_analysis_cache) > 20:
            oldest = list(_analysis_cache.keys())[0]
            del _analysis_cache[oldest]

        return jsonify({'success': True, 'data': data, 'cache_id': cache_id})
    except Exception as e:
        return jsonify({'error': f'分析失败: {str(e)}'}), 500


@app.route('/chart/<cache_id>/<chart_name>')
def chart_image(cache_id, chart_name):
    """根据缓存的 analysis_id 生成指定图表 PNG（小程序用）。"""
    data = _analysis_cache.get(cache_id)
    if data is None:
        return jsonify({'error': '缓存已过期，请重新上传文件'}), 404

    valid_charts = {
        'pie_spending', 'bar_food', 'line_weekly', 'doughnut_income',
        'bar_account', 'heatmap', 'budget_bar', 'score_ring',
    }
    if chart_name not in valid_charts:
        return jsonify({'error': f'未知图表类型: {chart_name}'}), 400

    try:
        buf = _generate_single_chart(data, chart_name)
        return send_file(buf, mimetype='image/png')
    except Exception as e:
        return jsonify({'error': f'图表生成失败: {str(e)}'}), 500


def _generate_single_chart(data, name):
    """生成单个图表，返回 BytesIO。"""
    from io import BytesIO
    budget_cfg = load_budget()

    if name == 'pie_spending':
        charts = all_charts(data)
        return charts['pie_spending']
    elif name == 'bar_food':
        charts = all_charts(data)
        return charts['bar_food']
    elif name == 'line_weekly':
        charts = all_charts(data)
        return charts['line_weekly']
    elif name == 'doughnut_income':
        charts = all_charts(data)
        return charts['doughnut_income']
    elif name == 'bar_account':
        charts = all_charts(data)
        return charts['bar_account']
    elif name == 'heatmap':
        return heatmap_daily(data['daily_list'])
    elif name == 'budget_bar':
        return bar_budget_vs_actual(data['cat1_list'], budget_cfg)
    elif name == 'score_ring':
        return _generate_score_ring(data.get('health', {}))
    raise ValueError(f'Unknown chart: {name}')


def _generate_score_ring(health):
    """生成健康评分圆环图 PNG（小程序专用）。"""
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    from io import BytesIO

    score = health.get('score', 0)
    color = health.get('color', '#7BC8A4')
    grade = health.get('grade', '')
    # Parse color hex
    if isinstance(color, str) and color.startswith('#'):
        h = color.lstrip('#')
        rgb = (int(h[0:2], 16) / 255, int(h[2:4], 16) / 255, int(h[4:6], 16) / 255)
    else:
        rgb = (0.48, 0.78, 0.64)  # default green

    fig, ax = plt.subplots(figsize=(3, 3))
    ax.set_xlim(-1.3, 1.3)
    ax.set_ylim(-1.3, 1.3)
    ax.set_aspect('equal')
    ax.axis('off')

    # 背景圆环
    from matplotlib.patches import Wedge
    bg = Wedge((0, 0), 1, 0, 360, width=0.25, color='#f0f0f0')
    ax.add_patch(bg)
    # 分数圆环
    angle = score / 100 * 360
    fg = Wedge((0, 0), 1, 90, 90 - angle, width=0.25, color=rgb)
    ax.add_patch(fg)
    # 中心文字
    ax.text(0, 0.15, f'{score}', ha='center', va='center', fontsize=48, fontweight='bold', color=rgb, fontfamily='sans-serif')
    ax.text(0, -0.35, grade, ha='center', va='center', fontsize=18, color='#666666', fontfamily='sans-serif')

    buf = BytesIO()
    fig.savefig(buf, format='png', dpi=100, bbox_inches='tight', transparent=True)
    plt.close(fig)
    buf.seek(0)
    return buf


@app.route('/generate', methods=['POST'])
def generate_route():
    """上传 Excel → 生成 PPT → 返回文件下载。"""
    if 'file' not in request.files:
        return jsonify({'error': '未找到文件'}), 400

    file = request.files['file']
    if not file.filename.endswith(('.xlsx', '.xls')):
        return jsonify({'error': '请上传 .xlsx 或 .xls 文件'}), 400

    try:
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            file.save(tmp.name)
            tmp_path = tmp.name

        data = analyze(tmp_path)
        # 附加评分和预算
        data['health'] = health_score(data)
        data['budget_comparison'] = compare_budget(data, load_budget())
        charts = all_charts(data)
        # 附加新图表
        charts['heatmap'] = heatmap_daily(data['daily_list'])
        charts['budget_bar'] = bar_budget_vs_actual(data['cat1_list'], load_budget())
        ppt_buf = build_ppt(data, charts)
        os.unlink(tmp_path)

        # 生成文件名
        filename = f"SpendLens_{data['month'].replace('年','').replace('月','')}.pptx"
        return send_file(
            ppt_buf,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
            as_attachment=True,
            download_name=filename,
        )
    except Exception as e:
        return jsonify({'error': f'生成失败: {str(e)}'}), 500


# ---- 预算管理 ----
@app.route('/budget', methods=['GET'])
def budget_get():
    """获取当前预算配置。"""
    return jsonify({'success': True, 'budget': load_budget()})


@app.route('/budget', methods=['POST'])
def budget_save():
    """保存预算配置。"""
    config = request.get_json()
    if not config:
        return jsonify({'error': '无效的预算数据'}), 400
    save_budget(config)
    return jsonify({'success': True, 'budget': config})


# ---- 规则引擎 ----
@app.route('/rules', methods=['GET'])
def rules_get():
    """获取当前规则列表。"""
    return jsonify({'success': True, 'rules': load_rules()})


@app.route('/rules', methods=['POST'])
def rules_save():
    """保存规则列表。"""
    rules = request.get_json()
    if not isinstance(rules, list):
        return jsonify({'error': '规则必须是列表'}), 400
    save_rules(rules)
    return jsonify({'success': True, 'rules': rules})


# ---- What-if 模拟 ----
@app.route('/simulate', methods=['POST'])
def simulate_route():
    """运行 What-if 场景模拟。"""
    body = request.get_json()
    if not body or 'file_data' not in body or 'scenario' not in body:
        return jsonify({'error': '需要 file_data 和 scenario'}), 400

    result = simulate(body['file_data'], body['scenario'])
    return jsonify({'success': True, 'result': result})


# ---- 环比对比 ----
@app.route('/compare', methods=['POST'])
def compare_route():
    """上月 vs 本月对比。需要上传本月文件，自动查找上月数据。"""
    if 'file' not in request.files:
        return jsonify({'error': '未找到文件'}), 400

    file = request.files['file']
    if not file.filename.endswith(('.xlsx', '.xls')):
        return jsonify({'error': '请上传 .xlsx 或 .xls 文件'}), 400

    try:
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            file.save(tmp.name)
            tmp_path = tmp.name

        curr_data = analyze(tmp_path)
        os.unlink(tmp_path)

        # 尝试加载上月数据
        from budget import load_prev_month
        prev_data = load_prev_month(curr_data.get('month', ''))

        if prev_data is None:
            return jsonify({
                'success': True,
                'mode': 'single_month',
                'curr_data': curr_data,
                'message': '未找到上月数据（首次使用？上传第二个月后即可对比）',
            })

        cmp = compare(prev_data, curr_data)
        # 对比图表
        cmp_chart_buf = bar_compare_monthly(cmp)
        import base64
        from io import BytesIO
        cmp_chart_b64 = base64.b64encode(cmp_chart_buf.getvalue()).decode()

        return jsonify({
            'success': True,
            'mode': 'comparison',
            'comparison': cmp,
            'chart_base64': cmp_chart_b64,
        })
    except Exception as e:
        return jsonify({'error': f'对比失败: {str(e)}'}), 500


# ---- PDF 导出 ----
@app.route('/generate-pdf', methods=['POST'])
def generate_pdf_route():
    """上传 Excel → 生成 PDF → 返回文件下载。"""
    if 'file' not in request.files:
        return jsonify({'error': '未找到文件'}), 400

    file = request.files['file']
    if not file.filename.endswith(('.xlsx', '.xls')):
        return jsonify({'error': '请上传 .xlsx 或 .xls 文件'}), 400

    try:
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            file.save(tmp.name)
            tmp_path = tmp.name

        data = analyze(tmp_path)
        data['health'] = health_score(data)
        data['budget_comparison'] = compare_budget(data, load_budget())
        charts = all_charts(data)
        charts['heatmap'] = heatmap_daily(data['daily_list'])
        charts['budget_bar'] = bar_budget_vs_actual(data['cat1_list'], load_budget())

        from pdf_builder import build_pdf
        pdf_buf = build_pdf(data, charts)
        os.unlink(tmp_path)

        filename = f"SpendLens_{data['month'].replace('年','').replace('月','')}.pdf"
        return send_file(
            pdf_buf,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=filename,
        )
    except Exception as e:
        return jsonify({'error': f'PDF生成失败: {str(e)}'}), 500


# ---- 分享 ----
@app.route('/share', methods=['POST'])
def share_create():
    """创建分享快照，返回分享链接。"""
    body = request.get_json()
    if not body or 'data' not in body:
        return jsonify({'error': '需要分析数据'}), 400

    share_id = create_share(body['data'])
    share_url = f'{BASE_URL}/share/{share_id}'
    return jsonify({'success': True, 'share_id': share_id, 'share_url': share_url})


@app.route('/share/<share_id>')
def share_view(share_id):
    """查看分享快照。"""
    snapshot = load_share(share_id)
    if snapshot is None:
        return render_template('share.html', error='分享不存在或已过期', data=None), 404
    return render_template('share.html', error=None, data=snapshot['data'],
                           created_at=snapshot.get('created_at', ''))


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5050))
    debug = os.environ.get('DEBUG', 'true').lower() == 'true'
    print(f'\n🏠 SpendLens 已启动')
    print(f'   → http://0.0.0.0:{port}')
    print(f'   → 拖拽 iCost Excel 文件即可分析\n')
    app.run(host='0.0.0.0', port=port, debug=debug)
