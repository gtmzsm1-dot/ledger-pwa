// 金额：一律用整数“分”。解析、显示都走字符串，不做浮点小数运算。
export const MAX_CENTS = 10_000_000_000_000; // 单笔上限一千亿元

const FW = { '０':'0','１':'1','２':'2','３':'3','４':'4','５':'5','６':'6','７':'7','８':'8','９':'9',
  '．':'.','，':',','－':'-','−':'-','＋':'+','￥':'¥','　':' ' };

export function normalizeText(s) {
  return String(s ?? '').replace(/[０-９．，－−＋￥　]/g, (c) => FW[c]);
}

/** 解析金额文本 → 分（整数），不合法返回 null。接受 12 / 12.5 / 1,234.56 / ¥12 / 12元 / -3 */
export function parseMoney(raw) {
  let s = normalizeText(raw).replace(/\s+/g, '').replace(/RMB|CNY|人民币|¥|元/gi, '');
  if (!s) return null;
  let neg = false;
  if (s[0] === '-') { neg = true; s = s.slice(1); } else if (s[0] === '+') s = s.slice(1);
  const m = /^(\d{1,3}(?:,\d{3})+|\d+)(?:\.(\d{1,2}))?$/.exec(s);
  if (!m) return null;
  const intDigits = m[1].replace(/,/g, '');
  if (intDigits.length > 13) return null;
  const frac = (m[2] ?? '').padEnd(2, '0');
  const cents = Number(intDigits) * 100 + Number(frac || '0');
  if (!Number.isSafeInteger(cents) || cents > MAX_CENTS) return null;
  return neg ? -cents : cents;
}

/** Excel 数字单元格 → 分。Excel 以双精度存数，只接受“误差 < 0.000001 的两位小数”，否则判为非法 */
export function centsFromExcelNumber(v) {
  if (typeof v !== 'number' || !Number.isFinite(v)) return null;
  const x = v * 100;
  const r = Math.round(x);
  if (Math.abs(x - r) > 1e-4) return null;       // 超过两位小数
  if (!Number.isSafeInteger(r) || Math.abs(r) > MAX_CENTS) return null;
  return r;
}

/** 分 → "1234.50" */
export function centsToPlain(c) {
  const neg = c < 0, a = Math.abs(c);
  return (neg ? '-' : '') + Math.floor(a / 100) + '.' + String(a % 100).padStart(2, '0');
}

/** 分 → "¥1,234.50" */
export function formatMoney(c, { sign = false } = {}) {
  const neg = c < 0, a = Math.abs(c);
  const yuan = String(Math.floor(a / 100)).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  const s = '¥' + yuan + '.' + String(a % 100).padStart(2, '0');
  return neg ? '-' + s : (sign && c > 0 ? '+' + s : s);
}

/** 分 → Excel 数值（只用于写出单元格显示，不参与计算） */
export function centsToExcelNumber(c) { return Number(centsToPlain(c)); }

export function isValidAmount(c) { return Number.isSafeInteger(c) && c > 0 && c <= MAX_CENTS; }
