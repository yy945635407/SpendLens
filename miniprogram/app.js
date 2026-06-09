// SpendLens 全局配置
App({
  globalData: {
    // 后端 API 地址 — 部署后替换为 Railway 域名
    // apiBase: 'https://spendlens-production.up.railway.app',
    apiBase: 'http://localhost:5050',
    analysisData: null,
    cacheId: null,
    budgetConfig: { monthly_total: 8000, categories: {} },
    autoRules: []
  }
});
