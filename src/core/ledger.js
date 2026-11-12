// 记账核心：所有改动经 commit（校验 → 应用 → 追加操作日志）。纯逻辑，无 DOM、无存储。
import { isValidAmount, formatMoney } from './money.js';
import { dayOf, monthOf, isValidMonth, daysInMonth, makeDay, noonOf } from './dates.js';
import { deepEqual, clone } from './util.js';

export class LedgerError extends Error {
  constructor(code, message, extra = {}) { super(message); this.code = code; Object.assign(this, extra); }
}

export const DEFAULT_SETTINGS = {
  duplicateWindowSeconds: 300,   // 防重复：5 分钟内同金额
  anomalyMultiplier: 2,          // 异常：超过历史均值 2 倍
  anomalyMinSamples: 3,          // 至少 3 笔历史才判断
  authValiditySeconds: 120,      // 月锁授权有效期
};

export const KIND_LABEL = {
  setup: '初始化', addLine: '新增业务线', updateLine: '修改业务线', addCategory: '新增大类', updateCategory: '修改大类',
  addTag: '新增标签', updateTag: '修改标签', addEntry: '记一笔', updateEntry: '修改流水', deleteEntry: '删除流水',
  restoreEntry: '恢复流水', batchTag: '批量打标签', importBatch: '导入 Excel', templateGenerate: '模板生成',
  addTemplate: '新增模板', updateTemplate: '修改模板', lockMonth: '锁定月份', unlockMonth: '解锁月份',
  settings: '修改设置', undo: '撤销', revert: '版本回退',
};

const COLL = { line: 'lines', category: 'categories', tag: 'tags', entry: 'entries', template: 'templates' };

// 浅拷贝：对象记录本身视为不可变（applyChange 总是放入新拷贝），只复制集合
function shallowState(s) {
  return { lines: { ...s.lines }, categories: { ...s.categories }, tags: { ...s.tags }, entries: { ...s.entries },
    templates: { ...s.templates }, lockedMonths: [...s.lockedMonths], settings: { ...s.settings } };
}

export function emptyState() {
  return { lines: {}, categories: {}, tags: {}, entries: {}, templates: {}, lockedMonths: [], settings: { ...DEFAULT_SETTINGS } };
}

// ---------- 变更（Change）----------
function objId(c) { return (c.before ?? c.after).id; }
export function invertChange(c) {
  if (c.t === 'lock') return { t: 'lock', month: c.month, locked: !c.locked };
  return { t: c.t, before: clone(c.after), after: clone(c.before) };
}
function beforeMatches(c, s) {
  if (c.t === 'lock') return s.lockedMonths.includes(c.month) === !c.locked;
  if (c.t === 'settings') return deepEqual(s.settings, c.before);
  const cur = s[COLL[c.t]][objId(c)];
  return deepEqual(cur ?? null, c.before ?? null);
}
export function applyChange(c, s) {
  if (c.t === 'lock') {
    const set = new Set(s.lockedMonths);
    c.locked ? set.add(c.month) : set.delete(c.month);
    s.lockedMonths = [...set].sort();
    return;
  }
  if (c.t === 'settings') { s.settings = clone(c.after); return; }
  const coll = s[COLL[c.t]], id = objId(c);
  if (c.after == null) delete coll[id]; else coll[id] = clone(c.after);
}
function changeDays(c) { return c.t === 'entry' ? [c.before?.day, c.after?.day].filter(Boolean) : []; }
function changeEntryId(c) { return c.t === 'entry' ? objId(c) : null; }

export function replay(ops) {
  const s = emptyState();
  for (const op of ops) for (const c of op.changes) applyChange(c, s);
  return s;
}

export function diffStates(a, b) {
  const out = [];
  for (const [t, coll] of Object.entries(COLL)) {
    const ids = [...new Set([...Object.keys(a[coll]), ...Object.keys(b[coll])])].sort();
    for (const id of ids) {
      if (!deepEqual(a[coll][id] ?? null, b[coll][id] ?? null)) out.push({ t, before: clone(a[coll][id] ?? null), after: clone(b[coll][id] ?? null) });
    }
  }
  for (const m of a.lockedMonths) if (!b.lockedMonths.includes(m)) out.push({ t: 'lock', month: m, locked: false });
  for (const m of b.lockedMonths) if (!a.lockedMonths.includes(m)) out.push({ t: 'lock', month: m, locked: true });
  if (!deepEqual(a.settings, b.settings)) out.push({ t: 'settings', before: clone(a.settings), after: clone(b.settings) });
  return out;
}

// ---------- 默认数据（中性名称；真实业务线名称由用户在 App 内填写，不入代码库）----------
export const DEFAULT_PRESET = [
  { name: '业务线一', income: ['拓新收入', '续费收入', '营业外收入'], expense: ['采购', '工资', '报销'] },
  { name: '业务线二', income: ['学费'], expense: ['工资', '备用金', '其他杂项'] },
  { name: '业务线三', income: ['房租', '其他收入'], expense: ['垫付车马费', '请客吃饭', '设备工具', '其他支出'] },
];

export class Ledger {
  constructor({ state = emptyState(), ops = [], deviceId = 'device', now = () => new Date(), uuid = () => crypto.randomUUID() } = {}) {
    this.state = state; this.ops = ops; this.deviceId = deviceId; this.now = now; this.uuid = uuid;
    this.onCommit = null;
  }

  static bootstrap(opts = {}, preset = DEFAULT_PRESET) {
    const l = new Ledger(opts);
    const changes = [];
    preset.forEach((p, i) => {
      const line = { id: l.uuid(), name: p.name, sortOrder: i, archived: false };
      changes.push({ t: 'line', before: null, after: line });
      for (const dir of ['income', 'expense']) {
        p[dir].forEach((n, j) => changes.push({ t: 'category', before: null,
          after: { id: l.uuid(), lineId: line.id, direction: dir, name: n, sortOrder: j, archived: false } }));
      }
    });
    changes.push({ t: 'settings', before: clone(l.state.settings), after: { ...DEFAULT_SETTINGS } });
    l._commit('setup', '初始化三条业务线与默认大类', changes);
    return l;
  }

  nowIso() { return this.now().toISOString(); }

  _commit(kind, summary, changes, extra = {}) {
    const s = shallowState(this.state);
    for (const c of changes) {
      if (!beforeMatches(c, s)) throw new LedgerError('inconsistent', '数据状态与预期不一致，已拒绝提交');
      applyChange(c, s);
    }
    const op = { id: this.uuid(), seq: this.ops.length + 1, timestamp: this.nowIso(), deviceId: this.deviceId,
      kind, summary, changes: clone(changes), ...extra };
    this.state = s;
    this.ops.push(op);
    this.onCommit?.(op);
    return op;
  }

  // ---------- 查询 ----------
  get lines() { return Object.values(this.state.lines).filter((l) => !l.archived).sort((a, b) => a.sortOrder - b.sortOrder); }
  categoriesOf(lineId, direction) {
    return Object.values(this.state.categories).filter((c) => c.lineId === lineId && c.direction === direction && !c.archived)
      .sort((a, b) => a.sortOrder - b.sortOrder);
  }
  tagsOf(categoryId) { return Object.values(this.state.tags).filter((t) => t.categoryId === categoryId && !t.archived); }
  activeEntries() { return Object.values(this.state.entries).filter((e) => !e.deletedAt); }
  trash() { return Object.values(this.state.entries).filter((e) => e.deletedAt).sort((a, b) => b.deletedAt.localeCompare(a.deletedAt)); }
  uncategorized(lineId) {
    return this.activeEntries().filter((e) => !e.categoryId && (!lineId || e.lineId === lineId)).sort((a, b) => b.occurredAt.localeCompare(a.occurredAt));
  }
  isLocked(day) { return this.state.lockedMonths.includes(monthOf(day)); }

  // ---------- 月锁 ----------
  checkAuth(auth) {
    if (!auth || !auth.issuedAt) throw new LedgerError('authRequired', '需要验证密码');
    const age = (this.now().getTime() - new Date(auth.issuedAt).getTime()) / 1000;
    if (age < 0 || age > this.state.settings.authValiditySeconds) throw new LedgerError('authExpired', '验证已过期，请重新输入密码');
  }
  _checkLock(days, auth, extraMonths = []) {
    const months = [...new Set([...days.map(monthOf), ...extraMonths])].filter((m) => this.state.lockedMonths.includes(m)).sort();
    if (!months.length) return;
    if (!auth) throw new LedgerError('monthLocked', `${months[0]} 已锁定，修改需要输入密码`, { month: months[0] });
    this.checkAuth(auth);
  }
  lockMonth(month) {
    if (!isValidMonth(month)) throw new LedgerError('invalid', '月份格式错误');
    if (this.state.lockedMonths.includes(month)) return null;
    return this._commit('lockMonth', `锁定 ${month}`, [{ t: 'lock', month, locked: true }]);
  }
  unlockMonth(month, auth) {
    if (!this.state.lockedMonths.includes(month)) return null;
    this.checkAuth(auth);
    return this._commit('unlockMonth', `解锁 ${month}`, [{ t: 'lock', month, locked: false }]);
  }

  // ---------- 业务线 / 大类 / 标签 ----------
  _name(s) { const n = String(s ?? '').trim(); if (!n) throw new LedgerError('emptyName', '名称不能为空'); return n; }
  addLine(name) {
    const n = this._name(name);
    if (this.lines.some((l) => l.name === n)) throw new LedgerError('duplicateName', `名称重复：${n}`);
    const line = { id: this.uuid(), name: n, sortOrder: Math.max(-1, ...Object.values(this.state.lines).map((l) => l.sortOrder)) + 1, archived: false };
    this._commit('addLine', `新增业务线 ${n}`, [{ t: 'line', before: null, after: line }]);
    return line;
  }
  updateLine(id, patch) {
    const old = this.state.lines[id]; if (!old) throw new LedgerError('notFound', '业务线不存在');
    const nw = { ...old, ...patch }; if (patch.name !== undefined) nw.name = this._name(patch.name);
    if (deepEqual(old, nw)) return null;
    return this._commit('updateLine', `修改业务线 ${old.name}${nw.name !== old.name ? '→' + nw.name : ''}`, [{ t: 'line', before: old, after: nw }]);
  }
  addCategory(lineId, direction, name) {
    if (!this.state.lines[lineId]) throw new LedgerError('notFound', '业务线不存在');
    if (!['income', 'expense'].includes(direction)) throw new LedgerError('invalid', '收支方向错误');
    const n = this._name(name);
    if (this.categoriesOf(lineId, direction).some((c) => c.name === n)) throw new LedgerError('duplicateName', `名称重复：${n}`);
    const cat = { id: this.uuid(), lineId, direction, name: n, sortOrder: this.categoriesOf(lineId, direction).length, archived: false };
    this._commit('addCategory', `新增大类 ${n}`, [{ t: 'category', before: null, after: cat }]);
    return cat;
  }
  updateCategory(id, patch) {
    const old = this.state.categories[id]; if (!old) throw new LedgerError('notFound', '大类不存在');
    const nw = { ...old, ...patch }; if (patch.name !== undefined) nw.name = this._name(patch.name);
    if (deepEqual(old, nw)) return null;
    return this._commit('updateCategory', `修改大类 ${old.name}`, [{ t: 'category', before: old, after: nw }]);
  }
  addTag(categoryId, name) {
    if (!this.state.categories[categoryId]) throw new LedgerError('notFound', '大类不存在');
    const n = this._name(name);
    if (this.tagsOf(categoryId).some((t) => t.name === n)) throw new LedgerError('duplicateName', `名称重复：${n}`);
    const tag = { id: this.uuid(), categoryId, name: n, archived: false };
    this._commit('addTag', `新增标签 ${n}`, [{ t: 'tag', before: null, after: tag }]);
    return tag;
  }
  updateTag(id, patch) {
    const old = this.state.tags[id]; if (!old) throw new LedgerError('notFound', '标签不存在');
    const nw = { ...old, ...patch }; if (patch.name !== undefined) nw.name = this._name(patch.name);
    if (deepEqual(old, nw)) return null;
    return this._commit('updateTag', `修改标签 ${old.name}`, [{ t: 'tag', before: old, after: nw }]);
  }

  // ---------- 流水 ----------
  validateDraft(d) {
    if (!this.state.lines[d.lineId]) throw new LedgerError('lineNotFound', '业务线不存在');
    if (!['income', 'expense'].includes(d.direction)) throw new LedgerError('invalid', '收支方向错误');
    if (!isValidAmount(d.amount)) throw new LedgerError('invalidAmount', '金额必须大于 0（以分为单位的整数）');
    if (!d.occurredAt || Number.isNaN(new Date(d.occurredAt).getTime())) throw new LedgerError('invalidDate', '日期错误');
    if (d.categoryId) {
      const c = this.state.categories[d.categoryId];
      if (!c) throw new LedgerError('categoryNotFound', '大类不存在');
      if (c.lineId !== d.lineId) throw new LedgerError('categoryMismatch', '大类不属于该业务线');
      if (c.direction !== d.direction) throw new LedgerError('categoryMismatch', '大类的收支方向不一致');
    }
    if (d.tagId) {
      const t = this.state.tags[d.tagId];
      if (!t) throw new LedgerError('tagNotFound', '小标签不存在');
      if (!d.categoryId || t.categoryId !== d.categoryId) throw new LedgerError('categoryMismatch', '小标签不属于所选大类');
    }
  }

  /** 保存前提醒：防重复、异常金额 */
  warnings(d, excludeId = null) {
    const out = [];
    const win = this.state.settings.duplicateWindowSeconds * 1000;
    const now = this.now().getTime(), occ = new Date(d.occurredAt).getTime();
    const dups = this.activeEntries().filter((e) => e.id !== excludeId && e.lineId === d.lineId && e.direction === d.direction && e.amount === d.amount &&
      (Math.abs(new Date(e.createdAt).getTime() - now) <= win || Math.abs(new Date(e.occurredAt).getTime() - occ) <= win));
    if (dups.length) out.push({ type: 'duplicate', entryIds: dups.map((e) => e.id), message: `${Math.round(win / 60000)} 分钟内已记过同金额 ${dups.length} 笔，可能重复` });
    if (d.direction === 'expense' && (d.tagId || d.categoryId)) {
      const hist = this.activeEntries().filter((e) => e.id !== excludeId && e.lineId === d.lineId && e.direction === 'expense' &&
        (d.tagId ? e.tagId === d.tagId : e.categoryId === d.categoryId));
      const k = this.state.settings.anomalyMultiplier;
      if (hist.length >= this.state.settings.anomalyMinSamples) {
        const sum = hist.reduce((a, e) => a + e.amount, 0);
        if (d.amount * hist.length > k * sum) {   // 整数比较：amount > k × 均值
          out.push({ type: 'anomaly', mean: Math.floor(sum / hist.length), samples: hist.length,
            message: `超过该分类历史均值 ${formatMoney(Math.floor(sum / hist.length))}（${hist.length} 笔）的 ${k} 倍` });
        }
      }
    }
    return out;
  }

  _makeEntry(d) {
    const t = this.nowIso();
    return { id: this.uuid(), lineId: d.lineId, direction: d.direction, amount: d.amount,
      day: dayOf(d.occurredAt), occurredAt: new Date(d.occurredAt).toISOString(),
      categoryId: d.categoryId || null, tagId: d.tagId || null,
      counterparty: String(d.counterparty ?? '').trim(), note: String(d.note ?? '').trim(),
      photoIds: [...(d.photoIds ?? [])], source: d.source ?? 'manual', batchId: d.batchId ?? null,
      createdAt: t, updatedAt: t, deletedAt: null, revision: 1 };
  }

  /** 记一笔：有提醒时需 confirmed=true 才保存 */
  addEntry(d, { confirmed = false, auth = null } = {}) {
    this.validateDraft(d);
    const e = this._makeEntry(d);
    this._checkLock([e.day], auth);
    if (!confirmed) {
      const w = this.warnings(d);
      if (w.length) throw new LedgerError('needsConfirmation', w.map((x) => x.message).join('；'), { warnings: w });
    }
    this._commit('addEntry', `${d.direction === 'income' ? '收入' : '支出'} ${formatMoney(d.amount)} ${e.counterparty}`.trim(), [{ t: 'entry', before: null, after: e }]);
    return e;
  }

  updateEntry(id, d, { confirmed = false, auth = null } = {}) {
    const old = this.state.entries[id];
    if (!old || old.deletedAt) throw new LedgerError('notFound', '流水不存在');
    const merged = { lineId: old.lineId, direction: old.direction, amount: old.amount, occurredAt: old.occurredAt, categoryId: old.categoryId,
      tagId: old.tagId, counterparty: old.counterparty, note: old.note, photoIds: old.photoIds, ...d };
    this.validateDraft(merged);
    const e = { ...old, lineId: merged.lineId, direction: merged.direction, amount: merged.amount,
      occurredAt: new Date(merged.occurredAt).toISOString(), day: dayOf(merged.occurredAt),
      categoryId: merged.categoryId || null, tagId: merged.tagId || null,
      counterparty: String(merged.counterparty ?? '').trim(), note: String(merged.note ?? '').trim(), photoIds: [...(merged.photoIds ?? [])] };
    const same = deepEqual({ ...old, updatedAt: 0, revision: 0 }, { ...e, updatedAt: 0, revision: 0 });
    if (same) return old;
    e.updatedAt = this.nowIso(); e.revision = old.revision + 1;
    this._checkLock([old.day, e.day], auth);
    if (!confirmed && (e.amount !== old.amount || e.tagId !== old.tagId || e.categoryId !== old.categoryId)) {
      const w = this.warnings(merged, id).filter((x) => x.type !== 'duplicate');
      if (w.length) throw new LedgerError('needsConfirmation', w.map((x) => x.message).join('；'), { warnings: w });
    }
    this._commit('updateEntry', `修改 ${formatMoney(old.amount)}${e.amount !== old.amount ? '→' + formatMoney(e.amount) : ''} ${e.counterparty}`.trim(), [{ t: 'entry', before: old, after: e }]);
    return e;
  }

  deleteEntry(id, { auth = null } = {}) {
    const old = this.state.entries[id];
    if (!old || old.deletedAt) throw new LedgerError('notFound', '流水不存在');
    this._checkLock([old.day], auth);
    const e = { ...old, deletedAt: this.nowIso(), updatedAt: this.nowIso(), revision: old.revision + 1 };
    return this._commit('deleteEntry', `删除 ${formatMoney(old.amount)} ${old.counterparty}`.trim(), [{ t: 'entry', before: old, after: e }]);
  }

  restoreEntry(id, { auth = null } = {}) {
    const old = this.state.entries[id];
    if (!old || !old.deletedAt) throw new LedgerError('notFound', '回收站里没有这笔');
    this._checkLock([old.day], auth);
    const e = { ...old, deletedAt: null, updatedAt: this.nowIso(), revision: old.revision + 1 };
    return this._commit('restoreEntry', `恢复 ${formatMoney(old.amount)} ${old.counterparty}`.trim(), [{ t: 'entry', before: old, after: e }]);
  }

  /** 批量打标签：整批一条日志，可整批撤销 */
  batchAssign(entryIds, categoryId, tagId = null, { auth = null } = {}) {
    const cat = this.state.categories[categoryId];
    if (!cat) throw new LedgerError('categoryNotFound', '大类不存在');
    if (tagId && this.state.tags[tagId]?.categoryId !== categoryId) throw new LedgerError('categoryMismatch', '小标签不属于所选大类');
    const changes = [];
    for (const id of [...new Set(entryIds)].sort()) {
      const old = this.state.entries[id];
      if (!old || old.deletedAt) throw new LedgerError('notFound', '流水不存在');
      if (old.lineId !== cat.lineId || old.direction !== cat.direction) throw new LedgerError('categoryMismatch', '所选流水与大类的业务线或收支方向不一致');
      if (old.categoryId === categoryId && old.tagId === tagId) continue;
      changes.push({ t: 'entry', before: old, after: { ...old, categoryId, tagId, updatedAt: this.nowIso(), revision: old.revision + 1 } });
    }
    if (!changes.length) return null;
    this._checkLock(changes.flatMap(changeDays), auth);
    return this._commit('batchTag', `批量打标签 ${changes.length} 笔 → ${cat.name}`, changes);
  }

  // ---------- 导入批次 ----------
  commitImport(drafts, fileName, { auth = null } = {}) {
    if (!drafts.length) throw new LedgerError('empty', '没有可导入的行');
    const batchId = this.uuid();
    const changes = drafts.map((d) => {
      const dd = { ...d, source: 'import', batchId };
      this.validateDraft(dd);
      return { t: 'entry', before: null, after: this._makeEntry(dd) };
    });
    this._checkLock(changes.flatMap(changeDays), auth);
    this._commit('importBatch', `导入 ${fileName} 共 ${changes.length} 笔`, changes, { batchId });
    return batchId;
  }
  undoImport(batchId, { auth = null } = {}) {
    const op = this.ops.find((o) => o.kind === 'importBatch' && o.batchId === batchId);
    if (!op) throw new LedgerError('notFound', '导入批次不存在');
    return this.undo(op.id, { auth });
  }
  importBatches() {
    const undone = this._undoneIds();
    return this.ops.filter((o) => o.kind === 'importBatch').map((o) => ({ batchId: o.batchId, opId: o.id, summary: o.summary, timestamp: o.timestamp, count: o.changes.length, undone: undone.has(o.id) }));
  }

  // ---------- 定期模板 ----------
  addTemplate(t) {
    const tpl = { id: this.uuid(), name: this._name(t.name), lineId: t.lineId, direction: t.direction, amount: t.amount,
      categoryId: t.categoryId || null, tagId: t.tagId || null, counterparty: t.counterparty ?? '', note: t.note ?? '',
      dayOfMonth: Math.min(31, Math.max(1, Math.trunc(t.dayOfMonth || 1))), active: t.active ?? true };
    this.validateDraft({ ...tpl, occurredAt: this.nowIso() });
    this._commit('addTemplate', `新增模板 ${tpl.name}`, [{ t: 'template', before: null, after: tpl }]);
    return tpl;
  }
  updateTemplate(id, patch) {
    const old = this.state.templates[id]; if (!old) throw new LedgerError('notFound', '模板不存在');
    const nw = { ...old, ...patch };
    this.validateDraft({ ...nw, occurredAt: this.nowIso() });
    if (deepEqual(old, nw)) return null;
    return this._commit('updateTemplate', `修改模板 ${old.name}`, [{ t: 'template', before: old, after: nw }]);
  }
  pendingTemplates(month) {
    return Object.values(this.state.templates).filter((t) => t.active &&
      !this.activeEntries().some((e) => e.source === 'template' && e.templateId === t.id && monthOf(e.day) === month))
      .sort((a, b) => a.name.localeCompare(b.name));
  }
  /** 一键生成某月的模板流水；同月已生成的不会重复生成 */
  generateTemplates(month, { only = null, auth = null } = {}) {
    if (!isValidMonth(month)) throw new LedgerError('invalid', '月份格式错误');
    const todo = this.pendingTemplates(month).filter((t) => !only || only.includes(t.id));
    if (!todo.length) return [];
    const [y, m] = month.split('-').map(Number);
    const entries = todo.map((t) => {
      const day = makeDay(y, m, Math.min(t.dayOfMonth, daysInMonth(y, m)));
      const e = this._makeEntry({ lineId: t.lineId, direction: t.direction, amount: t.amount, occurredAt: noonOf(day),
        categoryId: t.categoryId, tagId: t.tagId, counterparty: t.counterparty, note: t.note, source: 'template' });
      e.templateId = t.id;
      return e;
    });
    this._checkLock(entries.map((e) => e.day), auth);
    this._commit('templateGenerate', `${month} 模板生成 ${entries.length} 笔`, entries.map((e) => ({ t: 'entry', before: null, after: e })), { batchId: this.uuid() });
    return entries;
  }

  updateSettings(patch) {
    const nw = { ...this.state.settings, ...patch };
    if (deepEqual(nw, this.state.settings)) return null;
    return this._commit('settings', '修改设置', [{ t: 'settings', before: this.state.settings, after: nw }]);
  }

  // ---------- 撤销 / 版本回退 ----------
  _undoneIds() { return new Set(this.ops.filter((o) => o.undoes).map((o) => o.undoes)); }
  lastUndoable() {
    const undone = this._undoneIds();
    for (let i = this.ops.length - 1; i >= 0; i--) {
      const o = this.ops[i];
      if (o.kind !== 'undo' && o.kind !== 'setup' && !undone.has(o.id)) return o;
    }
    return null;
  }
  undoLast(opts = {}) {
    const op = this.lastUndoable();
    if (!op) throw new LedgerError('nothingToUndo', '没有可撤销的操作');
    return this.undo(op.id, opts);
  }
  /** 撤销指定操作。若相关记录之后又被改过，拒绝撤销（不覆盖后来的改动） */
  undo(opId, { auth = null } = {}) {
    const op = this.ops.find((o) => o.id === opId);
    if (!op) throw new LedgerError('notFound', '版本不存在');
    if (op.kind === 'setup') throw new LedgerError('notUndoable', '初始化不能撤销');
    if (this._undoneIds().has(opId)) throw new LedgerError('alreadyUndone', '该操作已撤销');
    const inv = [...op.changes].reverse().map(invertChange);
    const probe = shallowState(this.state);
    for (const c of inv) {
      if (!beforeMatches(c, probe)) throw new LedgerError('undoConflict', `「${op.summary}」之后相关记录又被改过，请先撤销后面的操作，或用版本回退`);
      applyChange(c, probe);
    }
    this._checkLock(inv.flatMap(changeDays), auth, inv.filter((c) => c.t === 'lock').map((c) => c.month));
    return this._commit('undo', `撤销：${op.summary}`, inv, { undoes: op.id, batchId: op.batchId ?? null });
  }
  stateAfter(opId) {
    const i = this.ops.findIndex((o) => o.id === opId);
    if (i < 0) throw new LedgerError('notFound', '版本不存在');
    return replay(this.ops.slice(0, i + 1));
  }
  /** 回退到某版本：生成一条“回退”操作（本身可撤销，历史不丢） */
  revertTo(opId, { auth = null } = {}) {
    const i = this.ops.findIndex((o) => o.id === opId);
    if (i < 0) throw new LedgerError('notFound', '版本不存在');
    const changes = diffStates(this.state, replay(this.ops.slice(0, i + 1)));
    if (!changes.length) return null;
    this._checkLock(changes.flatMap(changeDays), auth, changes.filter((c) => c.t === 'lock').map((c) => c.month));
    return this._commit('revert', `回退到版本 #${i + 1}（${this.ops[i].summary}）`, changes);
  }
  historyOf(entryId) { return this.ops.filter((o) => o.changes.some((c) => changeEntryId(c) === entryId)); }

  // ---------- 自检 ----------
  verifyIntegrity() {
    const issues = [];
    if (!deepEqual(replay(this.ops), this.state)) issues.push('操作日志重放结果与当前数据不一致');
    if (new Set(this.ops.map((o) => o.id)).size !== this.ops.length) issues.push('操作日志编号重复');
    for (const e of Object.values(this.state.entries)) {
      if (!isValidAmount(e.amount)) issues.push(`流水 ${e.id} 金额非法`);
      if (!this.state.lines[e.lineId]) issues.push(`流水 ${e.id} 业务线不存在`);
      if (e.categoryId) {
        const c = this.state.categories[e.categoryId];
        if (!c) issues.push(`流水 ${e.id} 大类不存在`);
        else if (c.lineId !== e.lineId || c.direction !== e.direction) issues.push(`流水 ${e.id} 大类不匹配`);
      }
      if (e.tagId && this.state.tags[e.tagId]?.categoryId !== e.categoryId) issues.push(`流水 ${e.id} 标签不匹配`);
      if (dayOf(e.occurredAt) !== e.day) issues.push(`流水 ${e.id} 日期字段不一致`);
    }
    return issues;
  }
}
