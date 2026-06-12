const app = getApp();
const { request, uploadFile, downloadFile, downloadFileWithProgress } = require('../../utils/api');
const { fmtMoney, fmtInt, fmtPct } = require('../../utils/format');

Page({
  data: {
    // 状态
    theme: 'pink',
    hasFile: false,
    fileName: '',
    fileSize: '',
    loading: false,
    showProgress: false,
    progressTitle: '正在分析账单…',
    progressSub: '',
    progressPct: 0,
    tabActive: 'overview',  // overview | budget | rules | simulate

    // 分析数据
    analysisData: null,
    cacheId: null,
    healthScore: null,
    healthGrade: '',
    healthText: '',
    healthColor: '#7BC8A4',
    stats: [],
    categories: [],
    budgetComparison: null,

    // 图表
    apiBase: '',
    chartPieUrl: '',
    chartFoodUrl: '',
    chartWeeklyUrl: '',
    chartIncomeUrl: '',
    chartAccountUrl: '',
    chartHeatmapUrl: '',
    chartBudgetUrl: '',
    chartScoreUrl: '',
    shareId: '',

    // 预算
    budgetTotal: 8000,
    budgetCats: [],

    // 规则
    rules: [],

    // 模拟
    simCards: [
      { title:'🍔 减少外卖 50%', desc:'若外卖支出减少一半', type:'reduce_category', category:'餐饮', by_pct:50 },
      { title:'💰 储蓄率达 70%', desc:'需要减少多少支出', type:'target_savings_rate', target:70 },
      { title:'📈 收入增加 10%', desc:'加薪/副业的影响', type:'increase_income', by_pct:10 },
      { title:'🎯 储蓄率达 80%', desc:'FIRE 级别储蓄率', type:'target_savings_rate', target:80 }
    ],
    simResult: null,
    simType: 'reduce_category',
    simCat: '',
    simPct: 20,

    // 详情弹窗
    showDetailModal: false,
    detailTitle: '',
    detailSubtitle: '',
    detailFormula: '',
    detailItems: [],
    detailColor: '#FF6B8A',
    detailTotal: '',
    detailTotalPct: ''
  },

  onLoad() {
    // 恢复保存的主题
    const saved = wx.getStorageSync('spendlens-theme') || 'pink';
    this._applyTheme(saved);
    this.loadBudgetConfig();
    this.loadRules();
  },

  _applyTheme(t) {
    const isBlue = t === 'blue';
    const vars = isBlue
      ? '--primary:#5B8DEF;--primary-light:#A8C8FF;--glow:#C8DDFF;--dark:#1E2A4A;--text:#2A3550;--muted:#90A4C4;--positive:#5BC8A4;--gold:#B8D4FF;--bg:#F0F5FF;--card-bg:rgba(255,255,255,0.65);--card-border:rgba(91,141,239,0.08);'
      : '--primary:#FF6B8A;--primary-light:#FFB3C6;--glow:#FFD4DF;--dark:#5C2D3E;--text:#3D1E2A;--muted:#C4909E;--positive:#7BC8A4;--gold:#FFD4B8;--bg:#FFF0F5;--card-bg:rgba(255,255,255,0.65);--card-border:rgba(255,107,138,0.08);';
    this.setData({ theme: t, containerStyle: vars });
    app.globalData.theme = t;
    const bg = isBlue ? '#F0F5FF' : '#FFF0F5';
    wx.setBackgroundColor({ backgroundColor: bg, backgroundColorTop: bg });
    wx.setNavigationBarColor({ frontColor: '#000000', backgroundColor: bg });
  },

  toggleTheme(e) {
    const t = e.currentTarget.dataset.theme;
    this._applyTheme(t);
    wx.setStorageSync('spendlens-theme', t);
  },

  onShareAppMessage() {
    const shareId = this.data.shareId || '';
    return {
      title: 'SpendLens · 我的账单健康报告',
      path: `/pages/share/share?shareId=${shareId}`,
      imageUrl: this.data.chartScoreUrl || ''
    };
  },

  // ============ 文件上传 ============
  chooseFile() {
    wx.chooseMessageFile({
      count: 1,
      type: 'file',
      extension: ['xlsx', 'xls'],
      success: (res) => {
        const file = res.tempFiles[0];
        this.setData({
          hasFile: true,
          fileName: file.name,
          fileSize: (file.size / 1024).toFixed(1) + ' KB'
        });
        this.runAnalysis(file.path);
      },
      fail: () => {
        wx.showToast({ title: '请选择 .xlsx 或 .xls 文件', icon: 'none' });
      }
    });
  },

  // ============ 分析 ============
  async runAnalysis(filePath) {
    this.setData({
      loading: true,
      showProgress: true,
      progressTitle: '正在分析账单…',
      progressSub: '上传文件中…',
      progressPct: 0
    });

    try {
      // 真实上传进度（0-50%），保存 task 引用用于取消
      const uploadPromise = uploadFile(filePath, (progress) => {
        const pct = Math.min(50, Math.floor(progress * 0.5));
        this.setData({
          progressSub: `上传文件中… ${progress}%`,
          progressPct: pct
        });
      });
      this._uploadTask = uploadPromise.task;
      const result = await uploadPromise;
      this._uploadTask = null;

      // 服务器处理完成，数据处理阶段（50-85%）
      this.setData({ progressSub: '分析数据中…', progressPct: 60 });

      const d = result.data;
      app.globalData.analysisData = d;
      app.globalData.cacheId = result.cache_id;

      this.setData({ progressSub: '计算统计数据…', progressPct: 70 });

      this.setData({
        loading: false,
        analysisData: d,
        cacheId: result.cache_id,
        healthScore: d.health ? d.health.score + '分' : '--',
        healthGrade: d.health ? d.health.grade + ' · ' + d.health.grade_text : '',
        healthText: d.health ? (d.health.suggestions || []).slice(0,2).join(' | ') : '',
        healthColor: d.health?.color || '#7BC8A4',
        stats: this.buildStats(d),
        categories: this.buildCategories(d),
        budgetComparison: d.budget_comparison,
        tabActive: 'overview'
      });

      // 加载图表（85-95%）
      this.setData({ progressSub: '生成图表中…', progressPct: 85 });
      this.loadCharts(result.cache_id);

      // 创建分享快照（95-100%）
      this.setData({ progressSub: '创建分享链接…', progressPct: 95 });
      await this.createShareSnapshot(d);

      this.setData({ showProgress: false, progressPct: 100 });
      wx.showToast({ title: '分析完成', icon: 'success' });
    } catch (err) {
      this.setData({ loading: false, showProgress: false, progressPct: 0 });
      if (!err || !err.aborted) {
        wx.showToast({ title: '分析失败: ' + (err.message || '未知错误'), icon: 'none', duration: 3000 });
      }
    }
  },

  buildStats(d) {
    return [
      { label:'总收入', en:'Total Income', value: fmtMoney(d.total_income), sub: d.income_count+'笔', color:'#7BC8A4', accent:'#7BC8A4' },
      { label:'总支出', en:'Total Expense', value: fmtMoney(d.total_expense), sub: d.expense_count+'笔', color:'#FF6B8A', accent:'#FF6B8A' },
      { label:'结余', en:'Balance', value: fmtMoney(d.balance), sub: '储蓄率 '+d.savings_rate+'%', color:'#FF85A2', accent:'#FF85A2' },
      { label:'日均', en:'Daily Average', value: fmtMoney(d.daily_avg), sub: d.month || '', color:'#FFB3C6', accent:'#FFB3C6' },
      { label:'最大支出', en:'Top Category', value: (d.cat1_list[0]||['-'])[0], sub: fmtMoney((d.cat1_list[0]||[0,0])[1]), color:'#FF7EB3', accent:'#FF7EB3' },
      { label:'交易笔数', en:'Transactions', value: d.transaction_count+'笔', sub: d.income_count+'收·'+d.expense_count+'支', color:'#C4909E', accent:'#C4909E' }
    ];
  },

  buildCategories(d) {
    const cc = ['#FF6B8A','#FF85A2','#FF9EBB','#FFB3C6','#FF7EB3','#FFD4B8','#C4909E','#E8A0B4'];
    const maxC = (d.cat1_list[0] || [0,1])[1] || 1;
    return (d.cat1_list || []).slice(0, 8).map((c, i) => ({
      name: c[0],
      amount: fmtMoney(c[1]),
      pct: fmtPct(c[1] / d.total_expense * 100),
      barW: (c[1] / maxC * 100).toFixed(0),
      color: cc[i] || '#FF6B8A'
    }));
  },

  // ============ 图表加载 ============
  loadCharts(cacheId) {
    const apiBase = app.globalData.apiBase;
    this.setData({ apiBase });
    const charts = {
      chartPieUrl: 'pie_spending',
      chartFoodUrl: 'bar_food',
      chartWeeklyUrl: 'line_weekly',
      chartIncomeUrl: 'doughnut_income',
      chartAccountUrl: 'bar_account',
      chartHeatmapUrl: 'heatmap',
      chartBudgetUrl: 'budget_bar',
      chartScoreUrl: 'score_ring',
    };
    const theme = app.globalData.theme || 'pink';
    for (const [key, name] of Object.entries(charts)) {
      this.setData({ [key]: `${apiBase}/chart/${cacheId}/${name}?theme=${theme}` });
    }
  },

  // ============ 标签切换 ============
  switchTab(e) {
    this.setData({ tabActive: e.currentTarget.dataset.tab });
  },

  // ============ 预算 ============
  async loadBudgetConfig() {
    try {
      const res = await request('/budget');
      if (res.success) {
        app.globalData.budgetConfig = res.budget;
        this.setData({
          budgetTotal: res.budget.monthly_total || 8000,
          budgetCats: Object.entries(res.budget.categories || {}).map(([k,v]) => ({ name: k, amount: v }))
        });
      }
    } catch (e) { /* use defaults */ }
  },

  onBudgetTotalChange(e) {
    this.setData({ budgetTotal: parseInt(e.detail.value) || 0 });
  },

  onBudgetCatChange(e) {
    const idx = e.currentTarget.dataset.idx;
    const cats = this.data.budgetCats;
    cats[idx].amount = parseInt(e.detail.value) || 0;
    this.setData({ budgetCats: cats });
  },

  addBudgetCat() {
    wx.showModal({
      title: '添加预算分类',
      editable: true,
      placeholderText: '分类名称（如：餐饮）',
      success: (res) => {
        if (res.confirm && res.content) {
          const cats = this.data.budgetCats;
          cats.push({ name: res.content, amount: 0 });
          this.setData({ budgetCats: cats });
        }
      }
    });
  },

  removeBudgetCat(e) {
    const idx = e.currentTarget.dataset.idx;
    const cats = this.data.budgetCats;
    cats.splice(idx, 1);
    this.setData({ budgetCats: cats });
  },

  async saveBudget() {
    const cfg = {
      monthly_total: this.data.budgetTotal,
      categories: {}
    };
    this.data.budgetCats.forEach(c => { cfg.categories[c.name] = c.amount; });
    try {
      await request('/budget', { method: 'POST', data: cfg });
      app.globalData.budgetConfig = cfg;
      wx.showToast({ title: '预算已保存 ✅', icon: 'success' });
    } catch (e) {
      wx.showToast({ title: '保存失败', icon: 'none' });
    }
  },

  // ============ 规则 ============
  async loadRules() {
    try {
      const res = await request('/rules');
      if (res.success) {
        app.globalData.autoRules = res.rules;
        this.setData({ rules: res.rules.map(r => ({
          ...r,
          desc: `如果 ${r.field} ${r.op} "${r.value}" → 🏷️ ${r.set_tag}`
        }))});
      }
    } catch (e) { /* */ }
  },

  async addRule(e) {
    // 简化：用表单输入
    wx.showModal({
      title: '添加规则',
      editable: true,
      placeholderText: '格式：备注 包含 外卖 → 标签名',
      success: async (res) => {
        if (res.confirm && res.content) {
          const parts = res.content.split(/\s+/);
          if (parts.length >= 4) {
            const rule = {
              id: Date.now(),
              field: parts[0] || 'note',
              op: parts[1] || 'contains',
              value: parts[2] || '',
              set_tag: parts[3] || ''
            };
            app.globalData.autoRules.push(rule);
            try {
              await request('/rules', { method:'POST', data: app.globalData.autoRules });
              this.loadRules();
              wx.showToast({ title: '规则已添加 ✅', icon: 'success' });
            } catch (e) {
              wx.showToast({ title: '保存失败', icon: 'none' });
            }
          } else {
            wx.showToast({ title: '格式错误：备注 包含 关键词 标签名', icon: 'none' });
          }
        }
      }
    });
  },

  async deleteRule(e) {
    const idx = e.currentTarget.dataset.idx;
    app.globalData.autoRules.splice(idx, 1);
    try {
      await request('/rules', { method:'POST', data: app.globalData.autoRules });
      this.loadRules();
      wx.showToast({ title: '规则已删除 ✅', icon: 'success' });
    } catch (e) {
      wx.showToast({ title: '删除失败', icon: 'none' });
    }
  },

  // ============ 模拟 ============
  runSimPreset(e) {
    const preset = e.currentTarget.dataset.preset;
    this.doSimulate(preset.type, preset.category || '', preset.by_pct || preset.target || 0);
  },

  async doSimulate(type, category, pct) {
    const d = this.data.analysisData;
    if (!d) { wx.showToast({ title: '请先上传账单', icon: 'none' }); return; }
    try {
      const res = await request('/simulate', {
        method: 'POST',
        data: { file_data: d, scenario: { type, by_pct: pct, category, target: pct } }
      });
      if (res.success) {
        this.setData({ simResult: res.result });
      } else {
        wx.showToast({ title: res.result?.error || '模拟失败', icon: 'none' });
      }
    } catch (e) {
      wx.showToast({ title: '模拟失败', icon: 'none' });
    }
  },

  runCustomSim() {
    const { simType, simCat, simPct } = this.data;
    this.doSimulate(simType, simCat, simPct);
  },

  onSimTypeChange(e) { this.setData({ simType: e.detail.value }); },
  onSimCatInput(e) { this.setData({ simCat: e.detail.value }); },
  onSimPctInput(e) { this.setData({ simPct: parseFloat(e.detail.value) || 0 }); },

  // ============ 导出 ============
  _handleDownloadError(err, label) {
    this._downloadTask = null;
    this.setData({ showProgress: false });
    if (!err || !err.aborted) {
      const reason = (err && err.message) ? err.message : '网络错误';
      wx.showToast({ title: label + '生成失败: ' + reason, icon: 'none' });
    }
  },

  generatePPT() {
    if (!this.data.cacheId) {
      wx.showToast({ title: '请先上传账单', icon: 'none' });
      return;
    }
    this.setData({
      showProgress: true,
      progressTitle: '正在生成 PPT…',
      progressSub: '服务器生成中…',
      progressPct: 5
    });

    const theme = app.globalData.theme || 'pink';
    const url = app.globalData.apiBase + '/generate-ppt/' + this.data.cacheId + '?theme=' + theme;

    const dlPromise = downloadFileWithProgress(url, (progress) => {
      // 真实下载进度：服务器生成完后开始传输，映射到 10-100%
      const pct = 10 + Math.floor(progress * 0.9);
      this.setData({
        progressSub: `下载中… ${progress}%`,
        progressPct: Math.min(100, pct)
      });
    });
    this._downloadTask = dlPromise.task;
    dlPromise.then((res) => {
      this._downloadTask = null;
      this.setData({ showProgress: false, progressPct: 100 });
      if (res.statusCode === 200) {
        wx.openDocument({
          filePath: res.tempFilePath,
          showMenu: true,
          success: () => wx.showToast({ title: 'PPT 已打开', icon: 'success' }),
          fail: () => wx.showToast({ title: '请安装 WPS 或 Office 打开', icon: 'none' })
        });
      } else {
        wx.showToast({ title: 'PPT 生成失败', icon: 'none' });
      }
    }).catch((err) => this._handleDownloadError(err, 'PPT'));
  },

  generatePDF() {
    if (!this.data.cacheId) {
      wx.showToast({ title: '请先上传账单', icon: 'none' });
      return;
    }
    this.setData({
      showProgress: true,
      progressTitle: '正在生成 PDF…',
      progressSub: '服务器生成中…',
      progressPct: 5
    });

    const theme = app.globalData.theme || 'pink';
    const url = app.globalData.apiBase + '/generate-pdf/' + this.data.cacheId + '?theme=' + theme;

    const dlPromise = downloadFileWithProgress(url, (progress) => {
      // 真实下载进度：服务器生成完后开始传输，映射到 10-100%
      const pct = 10 + Math.floor(progress * 0.9);
      this.setData({
        progressSub: `下载中… ${progress}%`,
        progressPct: Math.min(100, pct)
      });
    });
    this._downloadTask = dlPromise.task;
    dlPromise.then((res) => {
      this._downloadTask = null;
      this.setData({ showProgress: false, progressPct: 100 });
      if (res.statusCode === 200) {
        wx.openDocument({
          filePath: res.tempFilePath,
          showMenu: true,
          success: () => wx.showToast({ title: 'PDF 已打开', icon: 'success' }),
          fail: () => wx.showToast({ title: '请安装 PDF 阅读器打开', icon: 'none' })
        });
      } else {
        wx.showToast({ title: 'PDF 生成失败', icon: 'none' });
      }
    }).catch((err) => this._handleDownloadError(err, 'PDF'));
  },

  async createShareSnapshot(data) {
    try {
      const res = await request('/share', {
        method: 'POST',
        data: { data }
      });
      if (res.success) {
        this.setData({ shareId: res.share_id });
      }
    } catch (e) {
      // 静默失败，不影响主流程
      console.warn('分享快照创建失败:', e);
    }
  },

  resetAll() {
    this.setData({
      hasFile: false, fileName: '', fileSize: '',
      loading: false, showProgress: false,
      analysisData: null, cacheId: null,
      healthScore: null, healthGrade: '', healthText: '',
      stats: [], categories: [], budgetComparison: null,
      apiBase: '',
      chartPieUrl: '', chartFoodUrl: '', chartWeeklyUrl: '',
      chartIncomeUrl: '', chartAccountUrl: '', chartHeatmapUrl: '', chartBudgetUrl: '',
      chartScoreUrl: '',
      shareId: '',
      simResult: null, tabActive: 'overview',
      progressPct: 0
    });
  },

  // ============ 详情弹窗 ============
  showDetail(e) {
    const idx = e.currentTarget.dataset.idx;
    const d = this.data.analysisData;
    if (!d) return;

    let title = '', subtitle = '', formula = '', items = [], color = '', total = '', totalPct = '';

    switch (idx) {
      case 0: // 总收入
        title = '💰 总收入详情';
        subtitle = `合计 ${fmtMoney(d.total_income)} · ${d.income_count} 笔收入`;
        formula = '📐 总收入 = 所有收入来源金额之和';
        color = '#7BC8A4';
        items = (d.income_list || []).map(([name, amt]) => ({
          name,
          amount: `${fmtMoney(amt)}`,
          pct: `${(amt / d.total_income * 100).toFixed(1)}%`
        }));
        total = `${fmtMoney(d.total_income)}`;
        totalPct = '100%';
        break;

      case 1: // 总支出
        title = '📊 总支出详情';
        subtitle = `合计 ${fmtMoney(d.total_expense)} · ${d.expense_count} 笔支出`;
        formula = '📐 总支出 = 所有支出一级分类金额之和';
        color = '#FF6B8A';
        items = (d.cat1_list || []).map(([name, amt]) => ({
          name,
          amount: `${fmtMoney(amt)}`,
          pct: `${(amt / d.total_expense * 100).toFixed(1)}%`
        }));
        total = `${fmtMoney(d.total_expense)}`;
        totalPct = '100%';
        break;

      case 2: // 结余
        title = '🎀 结余详情';
        subtitle = `结余 ${fmtMoney(d.balance)} · 储蓄率 ${d.savings_rate}%`;
        formula = `📐 计算方式：\n结余 = 总收入 − 总支出\n= ${fmtMoney(d.total_income)} − ${fmtMoney(d.total_expense)}\n= ${fmtMoney(d.balance)}\n\n📐 储蓄率 = 结余 ÷ 总收入 × 100%\n= ${fmtMoney(d.balance)} ÷ ${fmtMoney(d.total_income)} × 100%\n= ${d.savings_rate}%`;
        color = '#FF85A2';
        break;

      case 3: // 日均
        title = '🍰 日均支出详情';
        subtitle = `日均 ${fmtMoney(d.daily_avg)} · ${d.month || ''}`;
        const daysInMonth = (d.daily_list && d.daily_list.length) || 30;
        formula = `📐 计算方式：\n日均支出 = 总支出 ÷ ${daysInMonth}天\n= ${fmtMoney(d.total_expense)} ÷ ${daysInMonth}\n= ${fmtMoney(d.daily_avg)}`;
        color = '#FFB3C6';
        items = (d.daily_list || []).map(([day, amt]) => ({
          name: day,
          amount: `${fmtMoney(amt)}`
        }));
        break;

      case 4: // 最大类
        const topCat = d.cat1_list && d.cat1_list[0] ? d.cat1_list[0] : ['无', 0];
        title = '💕 最大支出类别';
        subtitle = `${topCat[0]} · ${fmtMoney(topCat[1])} · 占${(topCat[1]/d.total_expense*100).toFixed(1)}%`;
        formula = '📐 按支出一级分类排序，金额最高者为最大类别';
        color = '#FF7EB3';
        items = (d.cat1_list || []).slice(0, 5).map(([name, amt], idx2) => ({
          name: `${idx2 === 0 ? '👑 ' : ''}${name}`,
          amount: `${fmtMoney(amt)}`,
          pct: `${(amt / d.total_expense * 100).toFixed(1)}%`
        }));
        break;

      case 5: // 交易
        title = '📋 交易总览';
        subtitle = `共 ${d.transaction_count} 笔 · ${d.income_count}收 · ${d.expense_count}支`;
        formula = `📐 计算方式：\n交易总数 = 收入笔数 + 支出笔数\n= ${d.income_count} + ${d.expense_count}\n= ${d.transaction_count} 笔\n\n📊 收支比 = ${d.income_count} : ${d.expense_count}`;
        color = '#C4909E';
        break;
    }

    this.setData({
      showDetailModal: true,
      detailTitle: title,
      detailSubtitle: subtitle,
      detailFormula: formula,
      detailItems: items,
      detailColor: color,
      detailTotal: total,
      detailTotalPct: totalPct
    });
  },

  closeDetail() {
    this.setData({ showDetailModal: false });
  },

  cancelProgress() {
    wx.showModal({
      title: '取消操作',
      content: '确定要取消当前操作吗？',
      success: (res) => {
        if (res.confirm) {
          // 终止正在进行的上传/下载任务
          if (this._uploadTask) {
            this._uploadTask.abort();
            this._uploadTask = null;
          }
          if (this._downloadTask) {
            this._downloadTask.abort();
            this._downloadTask = null;
          }
          this.setData({
            loading: false,
            showProgress: false,
            progressPct: 0
          });
          wx.showToast({ title: '已取消', icon: 'none' });
        }
      }
    });
  }
});
