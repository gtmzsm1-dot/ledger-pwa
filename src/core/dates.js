// 日期统一按北京时间计“自然日”，格式 YYYY-MM-DD；月份 YYYY-MM
export const TZ = 'Asia/Shanghai';
const fmt = new Intl.DateTimeFormat('en-CA', { timeZone: TZ, year: 'numeric', month: '2-digit', day: '2-digit' });
const fmtTime = new Intl.DateTimeFormat('en-GB', { timeZone: TZ, hour: '2-digit', minute: '2-digit', hour12: false });

export function dayOf(dateLike) {
  const d = dateLike instanceof Date ? dateLike : new Date(dateLike);
  return fmt.format(d); // en-CA → 2026-11-12
}
export function timeOf(dateLike) {
  const d = dateLike instanceof Date ? dateLike : new Date(dateLike);
  return fmtTime.format(d);
}
export const monthOf = (day) => day.slice(0, 7);

export function daysInMonth(y, m) {
  return [31, (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0 ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1];
}
export function makeDay(y, m, d) {
  if (!Number.isInteger(y) || !Number.isInteger(m) || !Number.isInteger(d)) return null;
  if (y < 1900 || y > 9999 || m < 1 || m > 12 || d < 1 || d > daysInMonth(y, m)) return null;
  return `${String(y).padStart(4, '0')}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
}
export function isValidDay(s) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s));
  return !!m && makeDay(+m[1], +m[2], +m[3]) === s;
}
export function isValidMonth(s) { return /^\d{4}-(0[1-9]|1[0-2])$/.test(String(s)); }

/** 北京时间某日中午的时间戳（ISO），用于导入/模板生成的发生时间 */
export function noonOf(day) { return new Date(day + 'T12:00:00+08:00').toISOString(); }

/** 解析文本日期：2026-11-12 / 2026/11/12 / 2026.11.12 / 2026年11月12日 / 20261112 */
export function parseDayText(raw) {
  const s = String(raw ?? '').trim().replace(/[０-９]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xFEE0));
  let m = /^(\d{4})\s*[-/.年]\s*(\d{1,2})\s*[-/.月]\s*(\d{1,2})\s*日?(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?$/.exec(s);
  if (m) return makeDay(+m[1], +m[2], +m[3]);
  m = /^(\d{4})(\d{2})(\d{2})$/.exec(s);
  if (m) return makeDay(+m[1], +m[2], +m[3]);
  return null;
}

/** Excel 日期序列号（1900 体系）→ YYYY-MM-DD，用纯整数推算，不受时区影响 */
export function excelSerialToDay(serial) {
  if (typeof serial !== 'number' || !Number.isFinite(serial) || serial < 1) return null;
  const n = Math.floor(serial + 1e-7);
  // 1899-12-30 为 0（已含 Excel 1900 闰年 bug 修正，适用于 1900-03-01 之后）
  const ms = Date.UTC(1899, 11, 30) + n * 86400000;
  const d = new Date(ms);
  return makeDay(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate());
}
export function dayToExcelSerial(day) {
  const [y, m, d] = day.split('-').map(Number);
  return Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);
}
export function addDays(day, n) {
  const [y, m, d] = day.split('-').map(Number);
  const t = new Date(Date.UTC(y, m - 1, d) + n * 86400000);
  return makeDay(t.getUTCFullYear(), t.getUTCMonth() + 1, t.getUTCDate());
}
