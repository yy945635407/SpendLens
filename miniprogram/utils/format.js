// 金额格式化
function fmtMoney(val) {
  if (val == null) return '¥0';
  const n = Number(val);
  const parts = Math.abs(n).toFixed(2).split('.');
  parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return (n < 0 ? '-¥' : '¥') + parts.join('.');
}

function fmtInt(val) {
  if (val == null) return '0';
  return Number(val).toLocaleString ? Number(val).toLocaleString() : String(val);
}

function fmtPct(val, decimals = 1) {
  return Number(val).toFixed(decimals) + '%';
}

function fmtSize(bytes) {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

module.exports = { fmtMoney, fmtInt, fmtPct, fmtSize };
