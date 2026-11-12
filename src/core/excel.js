// Excel 导出（含预览数据）、导入模板、导入解析与逐行校验。XLSX 由调用方注入（vendor/xlsx.mjs）。
import { centsToExcelNumber, centsFromExcelNumber, parseMoney, isValidAmount, formatMoney } from './money.js';
import { timeOf, excelSerialToDay, parseDayText, isValidDay, noonOf, dayToExcelSerial } from './dates.js';
import { KIND_LABEL } from './ledger.js';

export const IMPORT_COLUMNS = ['日期', '业务线', '收支', '大类', '小标签', '金额', '备注'];
const DIR_LABEL = { income: '收入', expense: '支出' };
const SOURCE_LABEL = { manual: '手记', import: '导入', template: '模板' };

/** 导出数据（预览与写文件用同一份） */
export function buildExport(ledger, { from, to, lineIds = null, includeDeleted = false } = {}) {
  const s = ledger.state;
  const entries = Object.values(s.entries).filter((e) =>
    (includeDeleted || !e.deletedAt) && (!from || e.day >= from) && (!to || e.day <= to) && (!lineIds || lineIds.includes(e.lineId)))
    .sort((a, b) => a.day.localeCompare(b.day) || a.occurredAt.localeCompare(b.occurredAt) || a.id.localeCompare(b.id));
  const header = ['日期', '时间', '业务线', '收支', '大类', '小标签', '金额', '对象/摘要', '备注', '凭证数', '来源', '记录ID'];
  const toRow = (e) => [e.day, timeOf(e.occurredAt), s.lines[e.lineId]?.name ?? '', DIR_LABEL[e.direction],
    s.categories[e.categoryId]?.name ?? '未分类', s.tags[e.tagId]?.name ?? '', e.amount, e.counterparty, e.note,
    (e.photoIds ?? []).length, SOURCE_LABEL[e.source] ?? e.source, e.id];
  const lines = ledger.lines.filter((l) => !lineIds || lineIds.includes(l.id));
  // 汇总：每条线 × 收支 × 大类
  const summary = [];
  let grandIn = 0, grandOut = 0;
  for (const l of lines) {
    const le = entries.filter((e) => e.lineId === l.id && !e.deletedAt);
    const inSum = le.filter((e) => e.direction === 'income').reduce((a, e) => a + e.amount, 0);
    const outSum = le.filter((e) => e.direction === 'expense').reduce((a, e) => a + e.amount, 0);
    grandIn += inSum; grandOut += outSum;
    summary.push([l.name, '收入合计', '', inSum]);
    summary.push([l.name, '支出合计', '', outSum]);
    summary.push([l.name, '结余', '', inSum - outSum]);
    for (const dir of ['income', 'expense']) {
      const byCat = new Map();
      for (const e of le.filter((x) => x.direction === dir)) {
        const k = s.categories[e.categoryId]?.name ?? '未分类';
        byCat.set(k, (byCat.get(k) ?? 0) + e.amount);
      }
      for (const [k, v] of [...byCat].sort()) summary.push([l.name, DIR_LABEL[dir], k, v]);
    }
  }
  summary.push(['全部', '收入合计', '', grandIn], ['全部', '支出合计', '', grandOut], ['全部', '结余', '', grandIn - grandOut]);
  return {
    range: { from: from ?? null, to: to ?? null },
    summary: { header: ['业务线', '项目', '大类', '金额'], rows: summary, moneyCol: 3 },
    detail: { header, rows: entries.map(toRow), moneyCol: 6 },
    perLine: lines.map((l) => ({ name: l.name, header, rows: entries.filter((e) => e.lineId === l.id).map(toRow), moneyCol: 6 })),
    count: entries.length,
  };
}

function sheetFrom(XLSX, table) {
  const aoa = [table.header, ...table.rows.map((r) => r.map((v, i) => (i === table.moneyCol ? centsToExcelNumber(v) : v)))];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  // 金额列格式 0.00
  const range = XLSX.utils.decode_range(ws['!ref']);
  for (let R = 1; R <= range.e.r; R++) {
    const cell = ws[XLSX.utils.encode_cell({ r: R, c: table.moneyCol })];
    if (cell && cell.t === 'n') cell.z = '0.00';
  }
  ws['!cols'] = table.header.map((h, i) => ({ wch: i === table.header.length - 1 ? 38 : Math.max(8, String(h).length * 2 + 2) }));
  return ws;
}
function safeSheetName(n, used) {
  let s = String(n).replace(/[\[\]:*?/\\]/g, ' ').slice(0, 28) || '表';
  let k = s, i = 2;
  while (used.has(k)) k = `${s}${i++}`;
  used.add(k);
  return k;
}

/** 生成导出工作簿（ArrayBuffer / Uint8Array） */
export function writeExportWorkbook(XLSX, exp, { opsSheet = null } = {}) {
  const wb = XLSX.utils.book_new();
  const used = new Set();
  XLSX.utils.book_append_sheet(wb, sheetFrom(XLSX, exp.summary), safeSheetName('汇总', used));
  XLSX.utils.book_append_sheet(wb, sheetFrom(XLSX, exp.detail), safeSheetName('全部明细', used));
  for (const p of exp.perLine) XLSX.utils.book_append_sheet(wb, sheetFrom(XLSX, p), safeSheetName(p.name, used));
  if (opsSheet) XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(opsSheet), safeSheetName('操作日志', used));
  return XLSX.write(wb, { type: 'array', bookType: 'xlsx', compression: true });
}

export function opsTable(ledger) {
  return [['版本', '时间', '操作', '说明', '变更条数'], ...ledger.ops.map((o) => [o.seq, o.timestamp, KIND_LABEL[o.kind] ?? o.kind, o.summary, o.changes.length])];
}

/** 导入模板（含填写说明与示例，示例行以 # 开头，导入时会跳过） */
export function writeImportTemplate(XLSX, ledger = null) {
  const wb = XLSX.utils.book_new();
  const today = '2026-01-15';
  const ws = XLSX.utils.aoa_to_sheet([
    IMPORT_COLUMNS,
    ['#示例 ' + today, '业务线一', '支出', '采购', '', 128.5, '打印纸 两箱（以 # 开头的行不会导入）'],
    ['#示例 ' + today, '业务线二', '收入', '学费', '', 3000, '示例：删掉本行再填写'],
  ]);
  ws['!cols'] = [{ wch: 14 }, { wch: 14 }, { wch: 6 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 40 }];
  XLSX.utils.book_append_sheet(wb, ws, '流水');
  const help = [
    ['列', '必填', '说明'],
    ['日期', '是', '2026-01-15 或 2026/1/15 或 2026年1月15日，也可用 Excel 日期格式'],
    ['业务线', '是', '必须与 App 里的业务线名称完全一致'],
    ['收支', '是', '填“收入”或“支出”'],
    ['大类', '否', '必须是该业务线、该收支方向下已有的大类；空着=未分类'],
    ['小标签', '否', '必须是所填大类下已有的小标签；填小标签必须先填大类'],
    ['金额', '是', '大于 0，最多两位小数，不要写负号；可带 ¥ 或千分位'],
    ['备注', '否', '任意文字'],
    ['', '', ''],
    ['规则', '', '错误行会标红且不导入；与已有流水“同业务线、同收支、同金额、同日”的行会提示可能重复；一次导入可整批撤销'],
  ];
  if (ledger) {
    help.push(['', '', ''], ['当前可用的业务线 / 大类 / 小标签', '', '']);
    for (const l of ledger.lines) for (const dir of ['income', 'expense']) for (const c of ledger.categoriesOf(l.id, dir)) {
      help.push([l.name, DIR_LABEL[dir], c.name + (ledger.tagsOf(c.id).length ? '：' + ledger.tagsOf(c.id).map((t) => t.name).join('、') : '')]);
    }
  }
  const hs = XLSX.utils.aoa_to_sheet(help);
  hs['!cols'] = [{ wch: 22 }, { wch: 8 }, { wch: 80 }];
  XLSX.utils.book_append_sheet(wb, hs, '填写说明');
  return XLSX.write(wb, { type: 'array', bookType: 'xlsx', compression: true });
}

function cellText(v) { return v == null ? '' : String(v).trim(); }

/**
 * 解析导入文件并逐行校验。返回 { rows: [{ rowNumber, values, errors, draft, duplicateOf, duplicateInFile }], headerError }
 * rowNumber 为 Excel 中的行号（表头为第 1 行）。
 */
export function parseImport(XLSX, data, ledger) {
  const wb = XLSX.read(data, { type: 'array', cellDates: false, raw: true });
  const ws = wb.Sheets['流水'] ?? wb.Sheets[wb.SheetNames[0]];
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
  const header = (aoa[0] ?? []).map(cellText);
  const idx = Object.fromEntries(IMPORT_COLUMNS.map((c) => [c, header.indexOf(c)]));
  const missing = ['日期', '业务线', '收支', '金额'].filter((c) => idx[c] < 0);
  if (missing.length) return { rows: [], headerError: `表头缺少列：${missing.join('、')}（请使用导入模板）` };

  const s = ledger.state;
  const lineByName = new Map(ledger.lines.map((l) => [l.name, l]));
  const existing = ledger.activeEntries();
  const seen = new Map();
  const rows = [];
  for (let i = 1; i < aoa.length; i++) {
    const r = aoa[i] ?? [];
    const get = (c) => (idx[c] >= 0 ? r[idx[c]] : null);
    const values = Object.fromEntries(IMPORT_COLUMNS.map((c) => [c, get(c)]));
    if (IMPORT_COLUMNS.every((c) => cellText(values[c]) === '')) continue;           // 空行
    if (cellText(values['日期']).startsWith('#')) continue;                            // 示例/注释行
    const errors = [];
    // 日期
    let day = null;
    const dv = values['日期'];
    if (typeof dv === 'number') day = excelSerialToDay(dv);
    else if (dv instanceof Date) day = null;
    else day = parseDayText(dv);
    if (!day || !isValidDay(day)) errors.push('日期无效');
    // 业务线
    const line = lineByName.get(cellText(values['业务线']));
    if (!line) errors.push(`业务线“${cellText(values['业务线'])}”不存在`);
    // 收支
    const dirText = cellText(values['收支']);
    const direction = dirText === '收入' ? 'income' : dirText === '支出' ? 'expense' : null;
    if (!direction) errors.push('收支须填“收入”或“支出”');
    // 金额
    const av = values['金额'];
    let amount = typeof av === 'number' ? centsFromExcelNumber(av) : parseMoney(av);
    if (amount == null) errors.push('金额格式错误（最多两位小数）');
    else if (!isValidAmount(amount)) { errors.push(amount <= 0 ? '金额必须大于 0' : '金额超过上限'); amount = null; }
    // 大类 / 小标签
    let categoryId = null, tagId = null;
    const catName = cellText(values['大类']), tagName = cellText(values['小标签']);
    if (catName && line && direction) {
      const cat = ledger.categoriesOf(line.id, direction).find((c) => c.name === catName);
      if (!cat) errors.push(`大类“${catName}”在该业务线的${dirText}下不存在`);
      else {
        categoryId = cat.id;
        if (tagName) {
          const tag = ledger.tagsOf(cat.id).find((t) => t.name === tagName);
          if (!tag) errors.push(`小标签“${tagName}”不存在`); else tagId = tag.id;
        }
      }
    } else if (tagName && !catName) errors.push('填了小标签但没填大类');

    const row = { rowNumber: i + 1, values, errors, draft: null, duplicateOf: [], duplicateInFile: null };
    if (!errors.length) {
      row.draft = { lineId: line.id, direction, amount, occurredAt: noonOf(day), categoryId, tagId, counterparty: '', note: cellText(values['备注']) };
      const key = `${line.id}|${direction}|${amount}|${day}`;
      row.duplicateOf = existing.filter((e) => e.lineId === line.id && e.direction === direction && e.amount === amount && e.day === day).map((e) => e.id);
      if (seen.has(key)) row.duplicateInFile = seen.get(key); else seen.set(key, row.rowNumber);
    }
    rows.push(row);
  }
  return { rows, headerError: null };
}

/** 导入预览汇总 */
export function importSummary(parsed) {
  const ok = parsed.rows.filter((r) => !r.errors.length);
  return { total: parsed.rows.length, valid: ok.length, errors: parsed.rows.length - ok.length,
    duplicates: ok.filter((r) => r.duplicateOf.length || r.duplicateInFile).length,
    amount: ok.reduce((a, r) => a + r.draft.amount, 0) };
}

/** 选出要导入的行：默认跳过错误行与疑似重复行；includeDuplicates 可放行重复行 */
export function selectImportDrafts(parsed, { includeDuplicates = false, rowNumbers = null } = {}) {
  return parsed.rows.filter((r) => !r.errors.length && (includeDuplicates || (!r.duplicateOf.length && !r.duplicateInFile)) &&
    (!rowNumbers || rowNumbers.includes(r.rowNumber))).map((r) => r.draft);
}

export { formatMoney, dayToExcelSerial };
