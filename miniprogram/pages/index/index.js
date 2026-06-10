const app = getApp();
const { request, uploadFile, downloadFile } = require('../../utils/api');
const { fmtMoney, fmtInt, fmtPct } = require('../../utils/format');

Page({
  data: {
    // 状态
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
    simPct: 20
  },

  onLoad() {
    this.loadBudgetConfig();
    this.loadRules();
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
      progressPct: 5
    });

    try {
      const result = await uploadFile(filePath);

      this.setData({ progressSub: '计算统计数据…', progressPct: 30 });

      const d = result.data;
      app.globalData.analysisData = d;
      app.globalData.cacheId = result.cache_id;

      this.setData({ progressSub: '生成健康评分…', progressPct: 60 });

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

      // 加载图表
      this.setData({ progressSub: '生成图表中…', progressPct: 75 });
      this.loadCharts(result.cache_id);

      // 自动创建分享快照
      this.setData({ progressSub: '创建分享链接…', progressPct: 90 });
      this.createShareSnapshot(d);

      this.setData({ showProgress: false, progressPct: 100 });
      wx.showToast({ title: '分析完成', icon: 'success' });
    } catch (err) {
      this.setData({ loading: false, showProgress: false, progressPct: 0 });
      wx.showToast({ title: '分析失败: ' + err.message, icon: 'none', duration: 3000 });
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
    for (const [key, name] of Object.entries(charts)) {
      this.setData({ [key]: `${apiBase}/chart/${cacheId}/${name}` });
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
  generatePPT() {
    if (!this.data.cacheId) {
      wx.showToast({ title: '请先上传账单', icon: 'none' });
      return;
    }
    this.setData({
      showProgress: true,
      progressTitle: '正在生成 PPT…',
      progressSub: '分析数据 · 绘制图表 · 排版中',
      progressPct: 0
    });
    // 模拟进度更新
    const pcts = [25, 50, 75, 95];
    const steps = ['分析数据中…', '绘制图表中…', '排版幻灯片…', '打包完成'];
    let step = 0;
    const timer = setInterval(() => {
      if (step < steps.length) {
        this.setData({ progressSub: steps[step], progressPct: pcts[step] });
        step++;
      }
    }, 800);

    const url = app.globalData.apiBase + '/generate-ppt/' + this.data.cacheId;
    wx.downloadFile({
      url,
      success: (res) => {
        clearInterval(timer);
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
      },
      fail: () => {
        clearInterval(timer);
        this.setData({ showProgress: false });
        wx.showToast({ title: '下载失败，请检查网络', icon: 'none' });
      }
    });
  },

  generatePDF() {
    if (!this.data.cacheId) {
      wx.showToast({ title: '请先上传账单', icon: 'none' });
      return;
    }
    this.setData({
      showProgress: true,
      progressTitle: '正在生成 PDF…',
      progressSub: '分析数据 · 渲染图表 · 排版中',
      progressPct: 0
    });
    const pcts = [25, 50, 75, 95];
    const steps = ['分析数据中…', '渲染图表中…', '排版页面中…', '打包完成'];
    let step = 0;
    const timer = setInterval(() => {
      if (step < steps.length) {
        this.setData({ progressSub: steps[step], progressPct: pcts[step] });
        step++;
      }
    }, 800);

    const url = app.globalData.apiBase + '/generate-pdf/' + this.data.cacheId;
    wx.downloadFile({
      url,
      success: (res) => {
        clearInterval(timer);
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
      },
      fail: () => {
        clearInterval(timer);
        this.setData({ showProgress: false });
        wx.showToast({ title: '下载失败，请检查网络', icon: 'none' });
      }
    });
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

  cancelProgress() {
    wx.showModal({
      title: '取消操作',
      content: '确定要取消当前操作吗？',
      success: (res) => {
        if (res.confirm) {
          this.setData({
            loading: false,
            showProgress: false,
            progressPct: 0
          });
        }
      }
    });
  }
});
