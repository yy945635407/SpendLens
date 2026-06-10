const app = getApp();
const { request } = require('../../utils/api');
const { fmtMoney } = require('../../utils/format');

Page({
  data: {
    error: null,
    data: null,
    stats: [],
    categories: [],
    suggestions: []
  },

  onLoad(options) {
    const shareId = options.shareId;
    if (!shareId) {
      this.setData({ error: '分享链接无效' });
      return;
    }
    this.loadShare(shareId);
  },

  async loadShare(shareId) {
    try {
      const res = await request(`/api/share/${shareId}`);
      if (res.error) {
        this.setData({ error: res.error });
        return;
      }
      const d = res.data;
      this.setData({
        data: d,
        stats: [
          { label:'💖 总收入', value: fmtMoney(d.total_income) },
          { label:'🌸 总支出', value: fmtMoney(d.total_expense) },
          { label:'🎀 结余', value: fmtMoney(d.balance) },
          { label:'🍰 日均', value: fmtMoney(d.daily_avg) },
          { label:'💕 最大类', value: (d.cat1_list[0]||['-'])[0] },
          { label:'✨ 交易', value: d.transaction_count+'笔' }
        ],
        categories: (d.cat1_list || []).slice(0,8).map(c => ({
          name: c[0],
          amount: fmtMoney(c[1]),
          pct: (c[1] / d.total_expense * 100).toFixed(1) + '%'
        })),
        suggestions: d.health?.suggestions || []
      });
    } catch (e) {
      this.setData({ error: '加载失败: ' + e.message });
    }
  }
});
