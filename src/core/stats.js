import { monthOf } from './dates.js';
import { parseMoney } from './money.js';

/** 三层余额：当日进/出、当月进/出、开始至今（截至 day，含）总进/出与剩余。lineId 为空=三线合计 */
export function balance(ledger, day, lineId = null) {
  const r = { todayIn: 0, todayOut: 0, monthIn: 0, monthOut: 0, totalIn: 0, totalOut: 0, remaining: 0 };
  const month = monthOf(day);
  for (const e of ledger.activeEntries()) {
    if (lineId && e.lineId !== lineId) continue;
    if (e.day > day) continue;
    const inc = e.direction === 'income';
    if (inc) r.totalIn += e.amount; else r.totalOut += e.amount;
    if (monthOf(e.day) === month) { if (inc) r.monthIn += e.amount; else r.monthOut += e.amount; }
    if (e.day === day) { if (inc) r.todayIn += e.amount; else r.todayOut += e.amount; }
  }
  r.remaining = r.totalIn - r.totalOut;
  return r;
}

/** 全局搜索：跨业务线，按金额（精确，或 min/max 区间，单位分）或关键词（对象、备注、业务线、大类、标签） */
export function search(ledger, { text = '', min = null, max = null, lineId = null, includeDeleted = false } = {}) {
  const s = ledger.state;
  const q = String(text).trim().toLowerCase();
  const amt = q ? parseMoney(q) : null;
  const list = Object.values(s.entries).filter((e) => {
    if (!includeDeleted && e.deletedAt) return false;
    if (lineId && e.lineId !== lineId) return false;
    if (min != null && e.amount < min) return false;
    if (max != null && e.amount > max) return false;
    if (!q) return true;
    if (amt != null && amt > 0 && e.amount === amt) return true;
    const hay = [e.counterparty, e.note, s.lines[e.lineId]?.name, s.categories[e.categoryId]?.name, s.tags[e.tagId]?.name, e.day]
      .filter(Boolean).join(' ').toLowerCase();
    return hay.includes(q);
  });
  return list.sort((a, b) => b.occurredAt.localeCompare(a.occurredAt));
}
