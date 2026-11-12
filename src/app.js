// 页面层：只调用 core 的公开 API；所有数据存在本机 IndexedDB
import * as XLSX from '../vendor/xlsx.mjs';
import { Ledger, KIND_LABEL } from './core/ledger.js';
import { balance, search } from './core/stats.js';
import { parseMoney, formatMoney, centsToPlain } from './core/money.js';
import { dayOf, timeOf, monthOf, noonOf } from './core/dates.js';
import { buildExport, writeExportWorkbook, writeImportTemplate, parseImport, importSummary, selectImportDrafts, opsTable } from './core/excel.js';
import { makeBackup, readBackup, referencedPhotoIds } from './core/backup.js';
import * as db from './db.js';
import { makePinRecord, checkPin, isValidPin } from './pin.js';
import { esc, sheet, confirmDlg, alertDlg, promptDlg, toast, pinPad, askPin } from './ui.js';

export const APP_VERSION = '1.0.0';
const AUTO_LOCK_MS = 3 * 60 * 1000;
const $app = document.getElementById('app');
const S = { ledger: null, pinRec: null, unlocked: false, view: { name: 'home' }, lineFilter: null, hiddenAt: 0,
  persisted: null, integrity: [], photoUrls: new Map(), saveChain: Promise.resolve(), lastSaveError: null };

const isStandalone = () => window.matchMedia?.('(display-mode: standalone)').matches || window.navigator.standalone === true;
const today = () => dayOf(new Date());
const lineName = (id) => S.ledger.state.lines[id]?.name ?? '（已删除）';
const catName = (id) => (id ? S.ledger.state.categories[id]?.name ?? '' : '未分类');
const tagName = (id) => (id ? S.ledger.state.tags[id]?.name ?? '' : '');

// ---------- 启动 ----------
async function boot() {
  if ('serviceWorker' in navigator) navigator.serviceWorker.register('./sw.js').catch(() => {});
  try { S.lastExport = await db.kvGet('lastExport'); } catch { S.lastExport = null; }
  try { S.persisted = navigator.storage?.persisted ? await navigator.storage.persisted() : null;
    if (S.persisted === false && navigator.storage?.persist) S.persisted = await navigator.storage.persist(); } catch { S.persisted = null; }
  let deviceId = await db.kvGet('deviceId');
  if (!deviceId) { deviceId = crypto.randomUUID(); await db.kvSet('deviceId', deviceId); }
  const rec = await db.loadLedger();
  if (rec) S.ledger = new Ledger({ state: rec.state, ops: rec.ops, deviceId });
  else {
    S.ledger = Ledger.bootstrap({ deviceId });
    await db.saveLedger(S.ledger.state, S.ledger.ops);
    await db.addSnapshot('初始化', S.ledger.state, S.ledger.ops);
  }
  attachLedger();
  S.integrity = S.ledger.verifyIntegrity();
  const snaps = await db.listSnapshots();
  if (!snaps.length || dayOf(snaps[0].createdAt) !== today()) await db.addSnapshot('每日自动备份', S.ledger.state, S.ledger.ops);
  S.pinRec = await db.kvGet('pin');
  document.addEventListener('visibilitychange', () => {
    if (document.hidden) S.hiddenAt = Date.now();
    else if (S.unlocked && S.pinRec && Date.now() - S.hiddenAt > AUTO_LOCK_MS) { S.unlocked = false; render(); }
  });
  render();
}

function attachLedger() {
  S.ledger.onCommit = (op) => {
    const { state, ops } = S.ledger;
    S.saveChain = S.saveChain.then(() => db.saveLedger(state, ops)).then(() => { S.lastSaveError = null; })
      .catch((e) => { S.lastSaveError = e; alertDlg('保存失败', '数据没有写入手机存储：' + (e?.message ?? e) + '\n请立即导出完整备份。'); });
    if (['importBatch', 'revert'].includes(op.kind)) db.addSnapshot(`${KIND_LABEL[op.kind]}后`, state, ops).catch(() => {});
  };
}

// ---------- 通用动作（确认 / 月锁授权）----------
async function run(action, { undoToast = true, okText = '已保存' } = {}) {
  const opts = { confirmed: false, auth: null };
  for (let i = 0; i < 4; i++) {
    try {
      const before = S.ledger.ops.length;
      const r = action(opts);
      if (S.ledger.ops.length > before && undoToast) toast(okText, { label: '撤销', run: () => doUndo() });
      return r ?? true;
    } catch (e) {
      if (e.code === 'needsConfirmation') {
        if (!(await confirmDlg('请确认', e.message + '\n\n确认无误再保存。', '确认无误，保存'))) return null;
        opts.confirmed = true; continue;
      }
      if (['monthLocked', 'authRequired', 'authExpired'].includes(e.code)) {
        if (!(await confirmDlg('月份已锁定', `${e.month ?? '该月份'} 已锁定。确定要修改吗？下一步需要输入密码。`, '继续修改'))) return null;
        if (!(await askPin('输入密码确认修改', (p) => checkPin(p, S.pinRec)))) return null;
        opts.auth = { issuedAt: new Date().toISOString(), method: 'pin' }; continue;
      }
      await alertDlg('没有保存', e.message ?? String(e));
      return null;
    }
  }
  return null;
}
async function doUndo() {
  const op = S.ledger.lastUndoable();
  if (!op) return toast('没有可撤销的操作');
  const r = await run((o) => S.ledger.undo(op.id, o), { undoToast: false });
  if (r) { toast('已撤销：' + op.summary); render(); }
}

function go(name, params = {}) { S.view = { name, ...params }; render(); window.scrollTo(0, 0); }

// ---------- 渲染 ----------
function render() {
  if (!S.ledger) return;
  if (!S.pinRec) return renderPinSetup();
  if (!S.unlocked) return renderLock();
  const v = S.view;
  const fn = VIEWS[v.name] ?? VIEWS.home;
  $app.innerHTML = fn(v);
  AFTER[v.name]?.(v);
}

function top(title, { back = 'home', right = '' } = {}) {
  return `<div class="top">${back ? `<button class="iconbtn back" data-act="go" data-to="${back}">‹ 返回</button>` : '<span></span>'}
    <h1>${esc(title)}</h1>${right || '<span class="iconbtn"></span>'}</div>`;
}

function renderPinSetup() {
  let first = null;
  // 两次输入：用闭包实现
  const step2 = async (p) => {
    if (p !== first) { first = null; renderPinSetup(); toast('两次不一致，请重新设置'); return null; }
    S.pinRec = await makePinRecord(p); await db.kvSet('pin', S.pinRec); S.unlocked = true; toast('密码已设置'); render(); return null;
  };
  pinPad($app, { title: '设置 6 位密码', sub: '用于打开 App 和修改已锁定月份。请牢记，忘记后只能用备份恢复到新安装。',
    onDone: async (p) => { first = p; pinPad($app, { title: '再输入一次', sub: '确认密码', onDone: step2 }); return null; } });
}

function renderLock() {
  pinPad($app, { title: '记账', sub: '输入密码', onDone: async (p) => {
    if (await checkPin(p, S.pinRec)) { S.unlocked = true; render(); return null; }
    return '密码不对，请重试';
  } });
}

// 有流水且从未导出或距上次导出超过 7 天时提醒
function backupDue(L) {
  if (!L.activeEntries().length) return false;
  if (!S.lastExport) return true;
  return Date.now() - Date.parse(S.lastExport) > 7 * 864e5;
}

const VIEWS = {};
const AFTER = {};

// 首页
VIEWS.home = () => {
  const L = S.ledger, d = today();
  const b = balance(L, d, S.lineFilter);
  const lines = L.lines;
  const recent = search(L, { lineId: S.lineFilter }).slice(0, 8);
  const banner = !isStandalone()
    ? `<div class="banner" id="a2hs"><b>请添加到主屏幕使用</b>点 Safari 底部的「分享」→「添加到主屏幕」，以后从主屏图标打开。只在浏览器标签页里用，系统可能清理数据。</div>` : '';
  const warn = [
    S.integrity.length ? `<div class="banner"><b>数据自检发现问题</b>${esc(S.integrity[0])}。请到「备份与恢复」用最近的备份恢复。</div>` : '',
    S.persisted === false ? `<div class="banner"><b>存储未获永久保留</b>系统空间紧张时可能清理，请定期「导出完整备份」。</div>` : '',
    backupDue(L) ? `<div class="banner" id="backup-due"><b>${S.lastExport ? `已 ${Math.floor((Date.now() - Date.parse(S.lastExport)) / 864e5)} 天没导出完整备份` : '还没导出过完整备份'}</b>数据只在这台手机上，请到「菜单 → 备份与恢复」导出到“文件”App。</div>` : '',
  ].join('');
  return `${top('记账', { back: '', right: '<button class="iconbtn" data-act="go" data-to="menu" aria-label="菜单">☰</button>' })}
  ${banner}${warn}
  <div class="page">
    <div class="seg"><button class="chip ${!S.lineFilter ? 'on' : ''}" data-act="filter" data-id="">全部</button>
      ${lines.map((l) => `<button class="chip ${S.lineFilter === l.id ? 'on' : ''}" data-act="filter" data-id="${l.id}">${esc(l.name)}</button>`).join('')}</div>
    <div class="card">
      <div class="bal-label">开始至今剩余</div>
      <div class="bal-main num" id="bal-remaining">${formatMoney(b.remaining)}</div>
      <div class="grid3">
        <div><div class="k">今日进</div><div class="v num in">${formatMoney(b.todayIn)}</div></div>
        <div><div class="k">今日出</div><div class="v num">${formatMoney(b.todayOut)}</div></div>
        <div><div class="k">&nbsp;</div><div class="v"></div></div>
        <div><div class="k">本月进</div><div class="v num in">${formatMoney(b.monthIn)}</div></div>
        <div><div class="k">本月出</div><div class="v num">${formatMoney(b.monthOut)}</div></div>
        <div><div class="k">本月结余</div><div class="v num">${formatMoney(b.monthIn - b.monthOut)}</div></div>
        <div><div class="k">累计进</div><div class="v num in">${formatMoney(b.totalIn)}</div></div>
        <div><div class="k">累计出</div><div class="v num">${formatMoney(b.totalOut)}</div></div>
      </div>
    </div>
    <div class="section">最近记录</div>
    ${entryList(recent, '还没有记录，点下方「记一笔」开始')}
    ${recent.length ? `<button class="btn ghost" data-act="go" data-to="list">查看全部 / 搜索</button>` : ''}
  </div>
  <div class="fab-wrap"><button class="btn primary big" data-act="new-entry" id="btn-new">记一笔</button></div>`;
};

function entryList(list, emptyText = '没有记录', { selectable = null } = {}) {
  if (!list.length) return `<div class="list"><div class="empty">${esc(emptyText)}</div></div>`;
  return `<div class="list">${list.map((e) => {
    const sub = [e.day, lineName(e.lineId), catName(e.categoryId) + (e.tagId ? '·' + tagName(e.tagId) : ''), e.note].filter(Boolean).join(' · ');
    const pick = selectable ? `<input type="checkbox" data-pick="${e.id}" ${selectable.has(e.id) ? 'checked' : ''}>` : '';
    return `<${selectable ? 'label' : 'button'} class="row" ${selectable ? '' : `data-act="detail" data-id="${e.id}"`}>${pick}
      <div class="main"><div class="t">${esc(e.counterparty || catName(e.categoryId))}${(e.photoIds ?? []).length ? ' 📎' : ''}</div><div class="s">${esc(sub)}</div></div>
      <div class="amt num ${e.direction === 'income' ? 'in' : ''}">${e.direction === 'income' ? '+' : '−'}${formatMoney(e.amount)}</div></${selectable ? 'label' : 'button'}>`;
  }).join('')}</div>`;
}

// 记一笔 / 编辑
function newDraft() {
  const L = S.ledger;
  const last = L.activeEntries().filter((e) => e.source === 'manual').sort((a, b) => b.createdAt.localeCompare(a.createdAt))[0];
  const lineId = S.lineFilter ?? last?.lineId ?? L.lines[0].id;
  const direction = last && last.lineId === lineId ? last.direction : 'expense';
  return { lineId, direction, amountText: '', day: today(), categoryId: last && last.lineId === lineId && last.direction === direction ? last.categoryId : null,
    tagId: null, counterparty: '', note: '', photoIds: [] };
}
VIEWS.form = (v) => {
  const L = S.ledger, d = v.draft;
  const cats = L.categoriesOf(d.lineId, d.direction);
  const tags = d.categoryId ? L.tagsOf(d.categoryId) : [];
  const locked = L.isLocked(d.day);
  return `${top(v.id ? '修改流水' : '记一笔', { back: v.id ? '' : 'home', right: '' })}
  <div class="page" id="form">
    ${v.id ? `<button class="iconbtn back" data-act="detail" data-id="${v.id}">‹ 返回</button>` : ''}
    <div class="toggle"><button data-act="dir" data-v="expense" class="${d.direction === 'expense' ? 'on' : ''}">支出</button><button data-act="dir" data-v="income" class="${d.direction === 'income' ? 'on' : ''}">收入</button></div>
    <input class="amount-input" id="f-amount" inputmode="decimal" placeholder="¥0.00" value="${esc(d.amountText)}" autocomplete="off" aria-label="金额">
    <div class="hint" id="f-amount-hint"></div>
    <div class="field"><span>业务线</span><div class="chips">${L.lines.map((l) => `<button class="chip ${d.lineId === l.id ? 'on' : ''}" data-act="line" data-id="${l.id}">${esc(l.name)}</button>`).join('')}</div></div>
    <div class="field"><span>大类</span><div class="chips">
      ${cats.map((c) => `<button class="chip ${d.categoryId === c.id ? 'on' : ''}" data-act="cat" data-id="${c.id}">${esc(c.name)}</button>`).join('')}
      <button class="chip add" data-act="add-cat">＋ 大类</button></div></div>
    ${d.categoryId ? `<div class="field"><span>小标签</span><div class="chips">
      ${tags.map((t) => `<button class="chip ${d.tagId === t.id ? 'on' : ''}" data-act="tag" data-id="${t.id}">${esc(t.name)}</button>`).join('')}
      <button class="chip add" data-act="add-tag">＋ 标签</button></div></div>` : ''}
    <label class="field"><span>${d.direction === 'expense' ? '买了什么 / 给谁（工资、报销）' : '收谁的 / 什么款'}</span><input class="input" id="f-cp" value="${esc(d.counterparty)}"></label>
    <label class="field"><span>备注</span><input class="input" id="f-note" value="${esc(d.note)}"></label>
    <label class="field"><span>日期${locked ? '（该月已锁定，保存需密码）' : ''}</span><input class="input" id="f-day" type="date" value="${esc(d.day)}"></label>
    <div class="field"><span>凭证照片（只存本机）</span><div class="thumbs" id="f-thumbs">
      ${d.photoIds.map((id) => `<button class="thumb" data-act="photo" data-id="${id}"><img data-photo="${id}" alt="凭证"></button>`).join('')}
      <label class="file-btn">拍照<input type="file" accept="image/*" capture="environment" data-photo-input></label>
      <label class="file-btn">选图<input type="file" accept="image/*" multiple data-photo-input></label></div></div>
  </div>
  <div class="fab-wrap"><button class="btn primary big" data-act="save-entry" id="btn-save">保存</button></div>`;
};
AFTER.form = (v) => {
  const amt = document.getElementById('f-amount');
  const hint = document.getElementById('f-amount-hint');
  const sync = () => {
    v.draft.amountText = amt.value; v.draft.counterparty = document.getElementById('f-cp').value;
    v.draft.note = document.getElementById('f-note').value; v.draft.day = document.getElementById('f-day').value || today();
    const c = parseMoney(amt.value);
    hint.textContent = amt.value && (c == null || c <= 0) ? '金额格式不对：大于 0，最多两位小数' : '';
  };
  $app.querySelectorAll('#form input').forEach((i) => i.addEventListener('input', sync));
  document.getElementById('f-day').addEventListener('change', () => { sync(); });
  v.sync = sync;
  if (!v.id && !v.draft.amountText) setTimeout(() => amt.focus(), 60);
  loadThumbs();
};

async function saveEntry(v) {
  v.sync?.();
  const d = v.draft;
  const amount = parseMoney(d.amountText);
  if (amount == null || amount <= 0) return alertDlg('金额不对', '请输入大于 0 的金额，最多两位小数');
  const occurredAt = d.day === today() && !v.id ? new Date().toISOString() : (v.id && S.ledger.state.entries[v.id].day === d.day ? S.ledger.state.entries[v.id].occurredAt : noonOf(d.day));
  const draft = { lineId: d.lineId, direction: d.direction, amount, occurredAt, categoryId: d.categoryId, tagId: d.tagId,
    counterparty: d.counterparty, note: d.note, photoIds: d.photoIds };
  const r = await run((o) => (v.id ? S.ledger.updateEntry(v.id, draft, o) : S.ledger.addEntry(draft, o)), { okText: v.id ? '已修改' : `已记 ${formatMoney(amount)}` });
  if (r) { if (v.id) go('detail', { id: v.id }); else go('home'); }
}

// 凭证照片：本地压缩后存 IndexedDB
async function compressImage(file) {
  const url = URL.createObjectURL(file);
  try {
    const img = await new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = rej; i.src = url; });
    const max = 1600, scale = Math.min(1, max / Math.max(img.naturalWidth, img.naturalHeight));
    const c = document.createElement('canvas');
    c.width = Math.round(img.naturalWidth * scale); c.height = Math.round(img.naturalHeight * scale);
    c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
    const blob = await new Promise((res) => c.toBlob(res, 'image/jpeg', 0.72));
    return new Uint8Array(await blob.arrayBuffer());
  } finally { URL.revokeObjectURL(url); }
}
async function addPhotos(files, v) {
  for (const f of files) {
    try {
      const bytes = await compressImage(f);
      const id = crypto.randomUUID();
      await db.putPhoto({ id, mime: 'image/jpeg', createdAt: new Date().toISOString(), bytes });
      v.draft.photoIds.push(id);
    } catch { toast('有一张图片读取失败'); }
  }
  render();
}
async function photoUrl(id) {
  if (S.photoUrls.has(id)) return S.photoUrls.get(id);
  const p = await db.getPhoto(id);
  if (!p) return '';
  const u = URL.createObjectURL(new Blob([p.bytes], { type: p.mime }));
  S.photoUrls.set(id, u); return u;
}
function loadThumbs() { document.querySelectorAll('img[data-photo]').forEach(async (img) => { img.src = await photoUrl(img.dataset.photo); }); }
async function showPhoto(id, v) {
  const u = await photoUrl(id);
  const canRemove = v?.name === 'form';
  const r = await sheet({ html: `<img class="photo-view" src="${u}" alt="凭证">`, buttons: [
    ...(canRemove ? [{ label: '从这笔移除', value: 'remove', kind: 'danger' }] : []), { label: '关闭', value: 'close' }] });
  if (r === 'remove') { v.draft.photoIds = v.draft.photoIds.filter((x) => x !== id); render(); }
}

// 详情
VIEWS.detail = (v) => {
  const e = S.ledger.state.entries[v.id];
  if (!e) return top('流水') + '<div class="page"><div class="empty">记录不存在</div></div>';
  const hist = S.ledger.historyOf(e.id).slice().reverse();
  return `${top('流水详情', { back: v.from ?? 'home' })}<div class="page">
    <div class="card"><div class="bal-label">${e.direction === 'income' ? '收入' : '支出'}${e.deletedAt ? '（在回收站）' : ''}</div>
      <div class="bal-main num ${e.direction === 'income' ? 'in' : ''}">${formatMoney(e.amount)}</div>
      <div class="kv"><span>业务线</span><span>${esc(lineName(e.lineId))}</span></div>
      <div class="kv"><span>大类 / 标签</span><span>${esc(catName(e.categoryId))}${e.tagId ? ' · ' + esc(tagName(e.tagId)) : ''}</span></div>
      <div class="kv"><span>对象</span><span>${esc(e.counterparty || '—')}</span></div>
      <div class="kv"><span>备注</span><span>${esc(e.note || '—')}</span></div>
      <div class="kv"><span>日期</span><span>${esc(e.day)} ${esc(timeOf(e.occurredAt))}${S.ledger.isLocked(e.day) ? ' 🔒' : ''}</span></div>
      <div class="kv"><span>来源</span><span>${esc({ manual: '手记', import: '导入', template: '模板' }[e.source] ?? e.source)} · 第 ${e.revision} 版</span></div>
      ${(e.photoIds ?? []).length ? `<div class="thumbs">${e.photoIds.map((id) => `<button class="thumb" data-act="photo" data-id="${id}"><img data-photo="${id}" alt="凭证"></button>`).join('')}</div>` : ''}
    </div>
    ${e.deletedAt ? `<button class="btn primary" data-act="restore" data-id="${e.id}">从回收站恢复</button>`
      : `<button class="btn" data-act="edit" data-id="${e.id}">修改</button><button class="btn danger" data-act="delete" data-id="${e.id}">删除（进回收站）</button>`}
    <div class="section">修改记录</div>
    <div class="list">${hist.map((o) => `<div class="row"><div class="main"><div class="t">${esc(KIND_LABEL[o.kind])}：${esc(o.summary)}</div><div class="s">${esc(dayOf(o.timestamp))} ${esc(timeOf(o.timestamp))} · 版本 #${o.seq}</div></div></div>`).join('')}</div>
  </div>`;
};
AFTER.detail = loadThumbs;

// 菜单
const MENU = [
  ['list', '🔍', '全部流水 / 搜索'], ['uncat', '🏷', '未分类 · 批量打标签'], ['cats', '🗂', '业务线、大类与标签'],
  ['templates', '🔁', '定期模板（房租、工资）'], ['months', '🔒', '月份锁定'], ['history', '🕘', '版本记录 · 撤销 · 回退'],
  ['trash', '🗑', '回收站'], ['export', '📤', '导出 Excel（先预览）'], ['import', '📥', '导入 Excel'],
  ['backup', '💾', '完整备份与恢复'], ['settings', '⚙️', '设置'],
];
VIEWS.menu = () => `${top('菜单')}<div class="page menu"><div class="list">
  ${MENU.map(([to, ico, label]) => `<button class="row" data-act="go" data-to="${to}"><span class="ico">${ico}</span><div class="main"><div class="t">${esc(label)}</div></div><span class="s">›</span></button>`).join('')}
  </div><button class="btn" data-act="undo-last">撤销最近一次操作</button><button class="btn" data-act="lock-now">立即锁定</button>
  <p class="hint">版本 ${APP_VERSION} · 数据只存在本机，不会上传。</p></div>`;

// 全部 / 搜索
VIEWS.list = (v) => {
  const res = search(S.ledger, { text: v.text ?? '', lineId: v.lineId ?? null });
  const sum = res.reduce((a, e) => a + (e.direction === 'income' ? e.amount : -e.amount), 0);
  return `${top('全部流水', { back: 'menu' })}<div class="page">
    <input class="input" id="q" placeholder="搜金额或关键词（对象、备注、分类）" value="${esc(v.text ?? '')}" inputmode="search">
    <div class="seg"><button class="chip ${!v.lineId ? 'on' : ''}" data-act="list-line" data-id="">全部</button>${S.ledger.lines.map((l) => `<button class="chip ${v.lineId === l.id ? 'on' : ''}" data-act="list-line" data-id="${l.id}">${esc(l.name)}</button>`).join('')}</div>
    <p class="hint">${res.length} 笔，净额 ${formatMoney(sum)}</p>
    ${entryList(res.slice(0, 300))}${res.length > 300 ? '<p class="hint">只显示前 300 笔，请缩小搜索范围</p>' : ''}</div>`;
};
AFTER.list = (v) => {
  const q = document.getElementById('q');
  q.addEventListener('change', () => { v.text = q.value; render(); });
  q.addEventListener('keydown', (e) => { if (e.key === 'Enter') { v.text = q.value; render(); } });
};

// 未分类批量打标签
VIEWS.uncat = (v) => {
  v.sel ??= new Set();
  v.lineId ??= S.ledger.lines[0].id;
  v.dir ??= 'expense';
  const list = S.ledger.uncategorized(v.lineId).filter((e) => e.direction === v.dir);
  const cats = S.ledger.categoriesOf(v.lineId, v.dir);
  return `${top('批量打标签', { back: 'menu' })}<div class="page">
    <div class="seg">${S.ledger.lines.map((l) => `<button class="chip ${v.lineId === l.id ? 'on' : ''}" data-act="uncat-line" data-id="${l.id}">${esc(l.name)}</button>`).join('')}</div>
    <div class="toggle"><button data-act="uncat-dir" data-v="expense" class="${v.dir === 'expense' ? 'on' : ''}">支出</button><button data-act="uncat-dir" data-v="income" class="${v.dir === 'income' ? 'on' : ''}">收入</button></div>
    <p class="hint">未分类 ${list.length} 笔，已选 ${v.sel.size} 笔。勾选后点下方大类即可整批贴上（可撤销）。</p>
    ${entryList(list, '没有未分类的流水', { selectable: v.sel })}
    ${list.length ? `<button class="btn ghost" data-act="uncat-all">全选 / 全不选</button>` : ''}
    <div class="section">贴到大类</div><div class="chips">${cats.map((c) => `<button class="chip" data-act="uncat-apply" data-id="${c.id}">${esc(c.name)}</button>`).join('')}</div></div>`;
};
AFTER.uncat = (v) => {
  $app.querySelectorAll('input[data-pick]').forEach((cb) => cb.addEventListener('change', () => {
    cb.checked ? v.sel.add(cb.dataset.pick) : v.sel.delete(cb.dataset.pick);
    const h = $app.querySelector('.hint'); if (h) h.textContent = h.textContent.replace(/已选 \d+ 笔/, `已选 ${v.sel.size} 笔`);
  }));
};

// 分类管理
VIEWS.cats = (v) => {
  v.lineId ??= S.ledger.lines[0].id;
  const L = S.ledger;
  const block = (dir) => L.categoriesOf(v.lineId, dir).map((c) => `<div class="row"><div class="main"><div class="t">${esc(c.name)}</div>
    <div class="chips">${L.tagsOf(c.id).map((t) => `<button class="chip" data-act="rename-tag" data-id="${t.id}">${esc(t.name)}</button>`).join('')}<button class="chip add" data-act="cats-add-tag" data-id="${c.id}">＋ 标签</button></div></div>
    <button class="iconbtn" data-act="rename-cat" data-id="${c.id}">改名</button></div>`).join('');
  return `${top('业务线与分类', { back: 'menu' })}<div class="page">
    <div class="seg">${L.lines.map((l) => `<button class="chip ${v.lineId === l.id ? 'on' : ''}" data-act="cats-line" data-id="${l.id}">${esc(l.name)}</button>`).join('')}</div>
    <button class="btn" data-act="rename-line" data-id="${v.lineId}">业务线改名（真实名称只存本机）</button>
    <div class="section">收入大类</div><div class="list">${block('income')}</div><button class="btn ghost" data-act="cats-add" data-dir="income">＋ 收入大类</button>
    <div class="section">支出大类</div><div class="list">${block('expense')}</div><button class="btn ghost" data-act="cats-add" data-dir="expense">＋ 支出大类</button></div>`;
};

// 定期模板
VIEWS.templates = () => {
  const L = S.ledger, month = monthOf(today());
  const tpls = Object.values(L.state.templates).sort((a, b) => a.name.localeCompare(b.name));
  const pending = new Set(L.pendingTemplates(month).map((t) => t.id));
  return `${top('定期模板', { back: 'menu' })}<div class="page">
    <div class="list">${tpls.length ? tpls.map((t) => `<div class="row"><div class="main"><div class="t">${esc(t.name)}${t.active ? '' : '（停用）'}</div>
      <div class="s">${esc(lineName(t.lineId))} · ${t.direction === 'income' ? '收入' : '支出'} · ${esc(catName(t.categoryId))} · 每月 ${t.dayOfMonth} 日 · ${pending.has(t.id) ? '本月未生成' : '本月已生成'}</div></div>
      <div class="amt num">${formatMoney(t.amount)}</div><button class="iconbtn" data-act="tpl-toggle" data-id="${t.id}">${t.active ? '停用' : '启用'}</button></div>`).join('') : '<div class="empty">还没有模板</div>'}</div>
    <button class="btn primary" data-act="tpl-gen" ${pending.size ? '' : 'disabled'}>一键生成本月（${pending.size} 笔）</button>
    <button class="btn" data-act="tpl-gen-month">生成其他月份…</button>
    <button class="btn ghost" data-act="tpl-add">＋ 新建模板</button>
    <p class="hint">同一模板同一个月只会生成一次；生成后可整批撤销。</p></div>`;
};

// 月份锁定
VIEWS.months = () => {
  const L = S.ledger;
  const months = [...new Set([...L.activeEntries().map((e) => monthOf(e.day)), monthOf(today())])].sort().reverse();
  return `${top('月份锁定', { back: 'menu' })}<div class="page"><p class="hint">锁定后该月仍可修改，但每次都要确认并输入密码。</p>
    <div class="list">${months.map((m) => { const lk = L.state.lockedMonths.includes(m); return `<div class="row"><div class="main"><div class="t">${m}${lk ? ' 🔒' : ''}</div></div>
      <button class="iconbtn" data-act="${lk ? 'unlock-month' : 'lock-month'}" data-m="${m}">${lk ? '解锁' : '锁定'}</button></div>`; }).join('')}</div></div>`;
};

// 版本记录
VIEWS.history = () => {
  const L = S.ledger;
  const undone = new Set(L.ops.filter((o) => o.undoes).map((o) => o.undoes));
  const ops = L.ops.slice().reverse().slice(0, 300);
  return `${top('版本记录', { back: 'menu' })}<div class="page"><p class="hint">每次改动都有一条记录。可撤销单条，或把全部数据回退到某个版本（回退本身也能撤销）。</p>
    <div class="list">${ops.map((o) => `<div class="row"><div class="main"><div class="t">#${o.seq} ${esc(KIND_LABEL[o.kind])}${undone.has(o.id) ? '（已撤销）' : ''}</div><div class="s">${esc(o.summary)} · ${esc(dayOf(o.timestamp))} ${esc(timeOf(o.timestamp))}</div></div>
      ${o.kind !== 'setup' && o.kind !== 'undo' && !undone.has(o.id) ? `<button class="iconbtn" data-act="undo-op" data-id="${o.id}">撤销</button>` : ''}
      <button class="iconbtn" data-act="revert-op" data-id="${o.id}">回到此版</button></div>`).join('')}</div></div>`;
};

// 回收站
VIEWS.trash = () => `${top('回收站', { back: 'menu' })}<div class="page"><p class="hint">删除的流水都在这里，可随时恢复；不会自动清空。</p>${entryList(S.ledger.trash(), '回收站是空的')}</div>`;

// 导出
VIEWS.export = (v) => {
  v.from ??= monthOf(today()) + '-01'; v.to ??= today();
  v.lineIds ??= S.ledger.lines.map((l) => l.id);
  const exp = buildExport(S.ledger, { from: v.from, to: v.to, lineIds: v.lineIds });
  v.exp = exp;
  return `${top('导出 Excel', { back: 'menu' })}<div class="page">
    <div class="grid2"><label class="field"><span>开始日期</span><input class="input" type="date" id="x-from" value="${v.from}"></label>
    <label class="field"><span>结束日期</span><input class="input" type="date" id="x-to" value="${v.to}"></label></div>
    <div class="chips">${[['month', '本月'], ['lastmonth', '上月'], ['year', '今年'], ['all', '全部']].map(([k, n]) => `<button class="chip" data-act="x-range" data-k="${k}">${n}</button>`).join('')}</div>
    <div class="section">业务线</div><div class="chips">${S.ledger.lines.map((l) => `<button class="chip ${v.lineIds.includes(l.id) ? 'on' : ''}" data-act="x-line" data-id="${l.id}">${esc(l.name)}</button>`).join('')}</div>
    <div class="section">预览 · 汇总</div>${table(exp.summary, 50)}
    <div class="section">预览 · 明细（${exp.count} 笔）</div>${table({ ...exp.detail, header: exp.detail.header.slice(0, 9), rows: exp.detail.rows.map((r) => r.slice(0, 9)) }, 100)}
    <button class="btn primary" data-act="x-go" ${exp.count ? '' : 'disabled'}>导出 Excel</button></div>`;
};
AFTER.export = (v) => {
  for (const [id, k] of [['x-from', 'from'], ['x-to', 'to']]) document.getElementById(id).addEventListener('change', (e) => { v[k] = e.target.value; render(); });
};
function table(t, limit) {
  return `<div class="tbl-wrap"><table><thead><tr>${t.header.map((h) => `<th>${esc(h)}</th>`).join('')}</tr></thead><tbody>
    ${t.rows.slice(0, limit).map((r) => `<tr>${r.map((c, i) => (i === t.moneyCol ? `<td class="r">${esc(centsToPlain(c))}</td>` : `<td>${esc(c)}</td>`)).join('')}</tr>`).join('')}
    </tbody></table></div>${t.rows.length > limit ? `<p class="hint">预览前 ${limit} 行，导出文件包含全部 ${t.rows.length} 行</p>` : ''}`;
}

// 导入
VIEWS.import = (v) => {
  const p = v.parsed;
  let body = `<p class="hint">先下载模板，按列填好（日期、业务线、收支、大类、小标签、金额、备注），再选择文件。错误行标红、不导入；疑似重复行标黄、默认不导入。</p>
    <button class="btn" data-act="tpl-download">下载导入模板</button>
    <label class="btn primary" for="imp-file">选择 Excel 文件</label><input id="imp-file" class="hidden" type="file" accept=".xlsx,.xls,.csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet">`;
  if (v.error) body += `<div class="banner"><b>无法读取</b>${esc(v.error)}</div>`;
  if (p) {
    const s = importSummary(p);
    const sel = selectImportDrafts(p, { includeDuplicates: v.includeDup });
    body += `<div class="section">${esc(v.fileName)}：共 ${s.total} 行，可导入 ${s.valid} 行（${formatMoney(s.amount)}），错误 ${s.errors} 行，疑似重复 ${s.duplicates} 行</div>
    <div class="tbl-wrap"><table><thead><tr><th>行</th><th>结果</th><th>日期</th><th>业务线</th><th>收支</th><th>大类</th><th>小标签</th><th>金额</th><th>备注</th></tr></thead><tbody>
    ${p.rows.map((r) => { const dup = r.duplicateOf.length || r.duplicateInFile; const cls = r.errors.length ? 'err' : dup ? 'warn' : '';
      const res = r.errors.length ? r.errors.join('；') : dup ? (r.duplicateOf.length ? '与已有流水重复' : `与第 ${r.duplicateInFile} 行重复`) : '可导入';
      const val = (k) => esc(r.values[k] ?? '');
      return `<tr class="${cls}"><td>${r.rowNumber}</td><td>${esc(res)}</td><td>${r.draft ? esc(dayOf(r.draft.occurredAt)) : val('日期')}</td><td>${val('业务线')}</td><td>${val('收支')}</td><td>${val('大类')}</td><td>${val('小标签')}</td><td class="r">${r.draft ? esc(centsToPlain(r.draft.amount)) : val('金额')}</td><td>${val('备注')}</td></tr>`; }).join('')}
    </tbody></table></div>
    ${s.duplicates ? `<label class="field"><input type="checkbox" id="imp-dup" ${v.includeDup ? 'checked' : ''}> 疑似重复的行也导入</label>` : ''}
    <button class="btn primary" data-act="imp-go" ${sel.length ? '' : 'disabled'}>导入 ${sel.length} 笔</button>`;
  }
  const batches = S.ledger.importBatches().slice().reverse();
  if (batches.length) body += `<div class="section">导入记录（可整批撤销）</div><div class="list">${batches.map((b) => `<div class="row"><div class="main"><div class="t">${esc(b.summary)}${b.undone ? '（已撤销）' : ''}</div><div class="s">${esc(dayOf(b.timestamp))} ${esc(timeOf(b.timestamp))}</div></div>${b.undone ? '' : `<button class="iconbtn" data-act="imp-undo" data-id="${b.batchId}">整批撤销</button>`}</div>`).join('')}</div>`;
  return `${top('导入 Excel', { back: 'menu' })}<div class="page">${body}</div>`;
};
AFTER.import = (v) => {
  document.getElementById('imp-file').addEventListener('change', async (e) => {
    const f = e.target.files[0]; if (!f) return;
    try { const parsed = parseImport(XLSX, new Uint8Array(await f.arrayBuffer()), S.ledger);
      if (parsed.headerError) { v.parsed = null; v.error = parsed.headerError; } else { v.parsed = parsed; v.error = null; v.fileName = f.name; }
    } catch (err) { v.parsed = null; v.error = '文件无法解析：' + (err.message ?? err); }
    render();
  });
  document.getElementById('imp-dup')?.addEventListener('change', (e) => { v.includeDup = e.target.checked; render(); });
};

// 备份与恢复
VIEWS.backup = (v) => {
  const L = S.ledger;
  return `${top('完整备份与恢复', { back: 'menu' })}<div class="page">
    <div class="card"><div class="kv"><span>流水</span><span>${Object.keys(L.state.entries).length} 笔</span></div>
      <div class="kv"><span>版本记录</span><span>${L.ops.length} 条</span></div>
      <div class="kv"><span>存储永久保留</span><span>${S.persisted === true ? '已获准' : S.persisted === false ? '未获准' : '未知'}</span></div>
      <div class="kv"><span>上次导出备份</span><span id="last-export">—</span></div></div>
    <button class="btn primary" data-act="bk-export">导出完整备份（含凭证照片 + Excel）</button>
    <p class="hint">生成一个 .zip 文件，用「存储到文件」存到 iCloud 云盘或电脑。换手机时在新手机上「从备份恢复」。</p>
    <label class="btn" for="bk-file">从备份文件恢复</label><input id="bk-file" class="hidden" type="file" accept=".zip,application/zip">
    <div class="section">本机自动备份（保留最近 ${db.SNAPSHOT_KEEP} 份，每天、导入、回退、恢复前自动生成）</div>
    <div class="list" id="snaps"><div class="empty">读取中…</div></div></div>`;
};
AFTER.backup = async () => {
  const last = await db.kvGet('lastExport'); const el = document.getElementById('last-export'); if (el && last) el.textContent = `${dayOf(last)} ${timeOf(last)}`;
  document.getElementById('bk-file').addEventListener('change', async (e) => { const f = e.target.files[0]; if (f) await restoreFromFile(f); });
  const snaps = await db.listSnapshots();
  const box = document.getElementById('snaps'); if (!box) return;
  box.innerHTML = snaps.map((s) => `<div class="row"><div class="main"><div class="t">${esc(s.reason)}</div><div class="s">${esc(dayOf(s.createdAt))} ${esc(timeOf(s.createdAt))} · ${s.entries} 笔</div></div>
    <button class="iconbtn" data-act="snap-restore" data-id="${esc(s.id)}">恢复到此</button></div>`).join('') || '<div class="empty">暂无</div>';
};

async function exportFullBackup() {
  const L = S.ledger;
  const ids = referencedPhotoIds(L);
  const photos = (await db.allPhotos()).filter((p) => ids.has(p.id)).map((p) => ({ ...p, bytes: new Uint8Array(p.bytes) }));
  const zipBytes = await makeBackup(L, photos, { appVersion: APP_VERSION });
  // 校验：导出后立即读回
  await readBackup(zipBytes);
  const xlsx = new Uint8Array(writeExportWorkbook(XLSX, buildExport(L, {}), { opsSheet: opsTable(L) }));
  const { unzipStore, zipStore } = await import('./core/zip.js');
  const files = unzipStore(zipBytes);
  const all = zipStore([...files, { name: '账目-全部.xlsx', data: xlsx }]);
  const stamp = new Date().toISOString().slice(0, 16).replace(/[-:T]/g, '');
  await saveFile(all, `记账完整备份-${stamp}.zip`, 'application/zip');
  S.lastExport = new Date().toISOString();
  await db.kvSet('lastExport', S.lastExport);
}

async function restoreFrom(state, ops, photos, label) {
  await db.addSnapshot('恢复前自动备份', S.ledger.state, S.ledger.ops);
  for (const p of photos) await db.putPhoto(p);
  S.ledger = new Ledger({ state, ops, deviceId: S.ledger.deviceId });
  attachLedger();
  await db.saveLedger(state, ops);
  S.integrity = S.ledger.verifyIntegrity();
  toast(`${label}完成；恢复前的数据已自动备份，可在本页回到恢复前`);
  go('backup');
}
async function restoreFromFile(f) {
  let r;
  try { r = await readBackup(new Uint8Array(await f.arrayBuffer())); } catch (e) { return alertDlg('备份无法恢复', e.message + '\n当前数据没有任何改动。'); }
  const ok = await confirmDlg('确认从备份恢复？', `备份时间：${dayOf(r.meta.exportedAt)} ${timeOf(r.meta.exportedAt)}\n流水 ${r.meta.counts.entries} 笔，版本记录 ${r.meta.counts.ops} 条，凭证 ${r.meta.counts.photos} 张。\n\n校验通过。当前数据会先自动备份，可随时回到恢复前。`, '恢复');
  if (!ok) return;
  if (!(await askPin('输入密码确认恢复', (p) => checkPin(p, S.pinRec)))) return;
  await restoreFrom(r.state, r.ops, r.photos, '从备份恢复');
}

async function saveFile(bytes, name, mime) {
  const file = new File([bytes], name, { type: mime });
  if (navigator.canShare?.({ files: [file] })) {
    try { await navigator.share({ files: [file], title: name }); return; } catch (e) { if (e?.name === 'AbortError') return; }
  }
  const url = URL.createObjectURL(file);
  const a = document.createElement('a'); a.href = url; a.download = name; document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30000);
}

// 设置
VIEWS.settings = () => {
  const st = S.ledger.state.settings;
  return `${top('设置', { back: 'menu' })}<div class="page"><div class="list">
    <button class="row" data-act="set-num" data-k="duplicateWindowSeconds" data-unit="60" data-label="防重复提醒时间窗（分钟）"><div class="main"><div class="t">防重复提醒</div><div class="s">${st.duplicateWindowSeconds / 60} 分钟内同金额再记会提醒</div></div></button>
    <button class="row" data-act="set-num" data-k="anomalyMultiplier" data-unit="1" data-label="异常预警倍数"><div class="main"><div class="t">异常金额预警</div><div class="s">支出超过该分类历史均值 ${st.anomalyMultiplier} 倍（至少 ${st.anomalyMinSamples} 笔历史）提醒</div></div></button>
    <button class="row" data-act="change-pin"><div class="main"><div class="t">修改密码</div></div></button></div>
    <p class="hint">数据只存在本机浏览器存储（IndexedDB），页面禁止任何外部连接。删除主屏图标或清除 Safari 网站数据会删掉数据，操作前请先导出完整备份。</p></div>`;
};

// ---------- 事件 ----------
document.addEventListener('change', (e) => {
  if (e.target.matches('input[data-photo-input]') && S.view.name === 'form') addPhotos([...e.target.files], S.view);
});

document.addEventListener('click', async (e) => {
  const b = e.target.closest('[data-act]');
  if (!b || !S.unlocked) return;
  const act = b.dataset.act, id = b.dataset.id, v = S.view, L = S.ledger;
  try {
    switch (act) {
      case 'go': return go(b.dataset.to);
      case 'filter': S.lineFilter = id || null; return render();
      case 'new-entry': return go('form', { draft: newDraft() });
      case 'detail': return go('detail', { id, from: v.name === 'detail' ? 'home' : v.name === 'form' ? 'home' : v.name });
      case 'edit': { const x = L.state.entries[id];
        return go('form', { id, draft: { lineId: x.lineId, direction: x.direction, amountText: centsToPlain(x.amount), day: x.day, categoryId: x.categoryId, tagId: x.tagId, counterparty: x.counterparty, note: x.note, photoIds: [...(x.photoIds ?? [])] } }); }
      case 'dir': v.sync?.(); v.draft.direction = b.dataset.v; v.draft.categoryId = null; v.draft.tagId = null; return render();
      case 'line': v.sync?.(); v.draft.lineId = id; v.draft.categoryId = null; v.draft.tagId = null; return render();
      case 'cat': v.sync?.(); v.draft.categoryId = v.draft.categoryId === id ? null : id; v.draft.tagId = null; return render();
      case 'tag': v.sync?.(); v.draft.tagId = v.draft.tagId === id ? null : id; return render();
      case 'add-cat': { v.sync?.(); const n = await promptDlg('新增大类', { placeholder: '例如：餐食后勤' }); if (!n) return;
        const c = await run(() => L.addCategory(v.draft.lineId, v.draft.direction, n), { okText: '已新增大类' }); if (c?.id) v.draft.categoryId = c.id; return render(); }
      case 'add-tag': { v.sync?.(); const n = await promptDlg('新增小标签', { placeholder: '例如：运营工资' }); if (!n) return;
        const t = await run(() => L.addTag(v.draft.categoryId, n), { okText: '已新增标签' }); if (t?.id) v.draft.tagId = t.id; return render(); }
      case 'save-entry': return saveEntry(v);
      case 'photo': return showPhoto(id, v);
      case 'delete': if (await confirmDlg('删除这笔？', '会放进回收站，可以恢复，也可以撤销。', '删除', true)) { if (await run((o) => L.deleteEntry(id, o), { okText: '已删除' })) go('home'); } return;
      case 'restore': if (await run((o) => L.restoreEntry(id, o), { okText: '已恢复' })) render(); return;
      case 'undo-last': return doUndo();
      case 'lock-now': S.unlocked = false; return render();
      case 'list-line': v.lineId = id || null; return render();
      case 'uncat-line': v.lineId = id; v.sel = new Set(); return render();
      case 'uncat-dir': v.dir = b.dataset.v; v.sel = new Set(); return render();
      case 'uncat-all': { const list = L.uncategorized(v.lineId).filter((x) => x.direction === v.dir); v.sel = v.sel.size === list.length ? new Set() : new Set(list.map((x) => x.id)); return render(); }
      case 'uncat-apply': { if (!v.sel.size) return toast('先勾选流水');
        const cat = L.state.categories[id];
        if (!(await confirmDlg('批量打标签', `把 ${v.sel.size} 笔贴到「${cat.name}」？可整批撤销。`, '确定'))) return;
        if (await run((o) => L.batchAssign([...v.sel], id, null, o), { okText: `已贴 ${v.sel.size} 笔` })) { v.sel = new Set(); render(); } return; }
      case 'cats-line': v.lineId = id; return render();
      case 'cats-add': { const n = await promptDlg(b.dataset.dir === 'income' ? '新增收入大类' : '新增支出大类'); if (n) { await run(() => L.addCategory(v.lineId, b.dataset.dir, n)); render(); } return; }
      case 'cats-add-tag': { const n = await promptDlg('新增小标签'); if (n) { await run(() => L.addTag(id, n)); render(); } return; }
      case 'rename-cat': { const c = L.state.categories[id]; const n = await promptDlg('大类改名', { value: c.name }); if (n && n !== c.name) { await run(() => L.updateCategory(id, { name: n })); render(); } return; }
      case 'rename-tag': { const t = L.state.tags[id]; const n = await promptDlg('标签改名', { value: t.name }); if (n && n !== t.name) { await run(() => L.updateTag(id, { name: n })); render(); } return; }
      case 'rename-line': { const l = L.state.lines[id]; const n = await promptDlg('业务线改名', { value: l.name }); if (n && n !== l.name) { await run(() => L.updateLine(id, { name: n })); render(); } return; }
      case 'tpl-add': return addTemplateFlow();
      case 'tpl-toggle': { const t = L.state.templates[id]; await run(() => L.updateTemplate(id, { active: !t.active })); return render(); }
      case 'tpl-gen': { const n = L.pendingTemplates(monthOf(today())).length;
        if (await confirmDlg('一键生成本月', `生成 ${n} 笔模板流水？`, '生成')) { await run((o) => L.generateTemplates(monthOf(today()), o), { okText: `已生成 ${n} 笔` }); render(); } return; }
      case 'tpl-gen-month': { const m = await promptDlg('生成哪个月？', { value: monthOf(today()), placeholder: '2026-11' }); if (!m) return;
        await run((o) => L.generateTemplates(m.trim(), o), { okText: '已生成' }); return render(); }
      case 'lock-month': if (await confirmDlg('锁定 ' + b.dataset.m, '锁定后修改该月任何流水都要输入密码。', '锁定')) { await run(() => L.lockMonth(b.dataset.m), { okText: '已锁定' }); render(); } return;
      case 'unlock-month': if (await confirmDlg('解锁 ' + b.dataset.m, '确定要解锁吗？下一步需要输入密码。', '继续')) {
        if (await askPin('输入密码解锁', (p) => checkPin(p, S.pinRec))) { await run(() => L.unlockMonth(b.dataset.m, { issuedAt: new Date().toISOString(), method: 'pin' }), { okText: '已解锁' }); render(); } } return;
      case 'undo-op': { const op = L.ops.find((o) => o.id === id); if (await confirmDlg('撤销这条操作？', op.summary, '撤销')) { await run((o) => L.undo(id, o), { undoToast: false }); toast('已撤销'); render(); } return; }
      case 'revert-op': { const op = L.ops.find((o) => o.id === id);
        if (!(await confirmDlg(`回到版本 #${op.seq}？`, `全部数据会回到「${op.summary}」之后的样子。之后的记录不会丢，回退本身也可撤销。`, '回退'))) return;
        await db.addSnapshot('回退前自动备份', L.state, L.ops);
        await run((o) => L.revertTo(id, o), { okText: '已回退' }); return render(); }
      case 'x-range': { const t = today(), m = monthOf(t);
        if (b.dataset.k === 'month') { v.from = m + '-01'; v.to = t; }
        if (b.dataset.k === 'lastmonth') { const [y, mm] = m.split('-').map(Number); const pm = mm === 1 ? `${y - 1}-12` : `${y}-${String(mm - 1).padStart(2, '0')}`; v.from = pm + '-01'; v.to = pm + '-31'; }
        if (b.dataset.k === 'year') { v.from = t.slice(0, 4) + '-01-01'; v.to = t; }
        if (b.dataset.k === 'all') { v.from = '1900-01-01'; v.to = '9999-12-31'; }
        return render(); }
      case 'x-line': v.lineIds = v.lineIds.includes(id) ? v.lineIds.filter((x) => x !== id) : [...v.lineIds, id]; return render();
      case 'x-go': { const buf = new Uint8Array(writeExportWorkbook(XLSX, v.exp)); return saveFile(buf, `账目-${v.from}至${v.to}.xlsx`, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'); }
      case 'tpl-download': return saveFile(new Uint8Array(writeImportTemplate(XLSX, L)), '记账导入模板.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      case 'imp-go': { const drafts = selectImportDrafts(v.parsed, { includeDuplicates: v.includeDup });
        if (!(await confirmDlg('确认导入', `导入 ${drafts.length} 笔？整批可撤销。`, '导入'))) return;
        await db.addSnapshot('导入前自动备份', L.state, L.ops);
        if (await run((o) => L.commitImport(drafts, v.fileName, o), { okText: `已导入 ${drafts.length} 笔` })) { v.parsed = null; render(); } return; }
      case 'imp-undo': if (await confirmDlg('整批撤销导入？', '这一批导入的流水会全部移除，数据回到导入前。', '整批撤销', true)) { await run((o) => L.undoImport(id, o), { undoToast: false }); toast('已整批撤销'); render(); } return;
      case 'bk-export': b.disabled = true; b.textContent = '正在生成…'; try { await exportFullBackup(); toast('备份已生成'); } finally { render(); } return;
      case 'snap-restore': { const s = await db.getSnapshot(id);
        if (!(await confirmDlg('恢复到这份本机备份？', `${s.reason} · ${dayOf(s.createdAt)} ${timeOf(s.createdAt)} · ${s.entries} 笔。当前数据会先自动备份。`, '恢复'))) return;
        const chk = new Ledger({ state: s.state, ops: s.ops }).verifyIntegrity(); if (chk.length) return alertDlg('这份备份自检未通过', chk[0]);
        if (!(await askPin('输入密码确认恢复', (p) => checkPin(p, S.pinRec)))) return;
        return restoreFrom(s.state, s.ops, [], '恢复本机备份'); }
      case 'set-num': { const k = b.dataset.k, unit = +b.dataset.unit; const cur = L.state.settings[k] / unit;
        const n = await promptDlg(b.dataset.label, { value: String(cur), inputmode: 'numeric' }); if (n == null) return;
        const x = Math.round(Number(n)); if (!Number.isInteger(x) || x < 1 || x > 1000) return toast('请输入 1–1000 的整数');
        await run(() => L.updateSettings({ [k]: x * unit })); return render(); }
      case 'change-pin': { if (!(await askPin('输入当前密码', (p) => checkPin(p, S.pinRec)))) return;
        const n1 = await promptDlg('新密码（6 位数字）', { type: 'password', inputmode: 'numeric' }); if (!n1 || !isValidPin(n1)) return toast('需要 6 位数字');
        const n2 = await promptDlg('再输入一次', { type: 'password', inputmode: 'numeric' }); if (n1 !== n2) return toast('两次不一致');
        S.pinRec = await makePinRecord(n1); await db.kvSet('pin', S.pinRec); return toast('密码已修改'); }
    }
  } catch (err) { alertDlg('出错了', err.message ?? String(err)); }
});

async function addTemplateFlow() {
  const L = S.ledger;
  const lines = L.lines;
  const html = `<label class="field"><span>名称</span><input class="input" id="t-name" placeholder="例如：房租"></label>
    <label class="field"><span>业务线</span><select class="input" id="t-line">${lines.map((l) => `<option value="${l.id}">${esc(l.name)}</option>`).join('')}</select></label>
    <label class="field"><span>收支</span><select class="input" id="t-dir"><option value="expense">支出</option><option value="income">收入</option></select></label>
    <label class="field"><span>大类</span><select class="input" id="t-cat"></select></label>
    <label class="field"><span>金额</span><input class="input" id="t-amt" inputmode="decimal"></label>
    <label class="field"><span>每月几号</span><input class="input" id="t-day" inputmode="numeric" value="1"></label>
    <label class="field"><span>对象 / 备注</span><input class="input" id="t-cp"></label>`;
  const res = await sheet({ title: '新建定期模板', html, onMount: (w) => {
      const fill = () => { const lid = w.querySelector('#t-line').value, dir = w.querySelector('#t-dir').value;
        w.querySelector('#t-cat').innerHTML = '<option value="">未分类</option>' + L.categoriesOf(lid, dir).map((c) => `<option value="${c.id}">${esc(c.name)}</option>`).join(''); };
      w.querySelector('#t-line').addEventListener('change', fill); w.querySelector('#t-dir').addEventListener('change', fill); fill(); },
    buttons: [{ label: '保存', kind: 'primary', value: (w) => ({ name: w.querySelector('#t-name').value, lineId: w.querySelector('#t-line').value, direction: w.querySelector('#t-dir').value,
      categoryId: w.querySelector('#t-cat').value || null, amount: parseMoney(w.querySelector('#t-amt').value), dayOfMonth: Number(w.querySelector('#t-day').value), counterparty: w.querySelector('#t-cp').value }) },
      { label: '取消', value: null }] });
  if (!res) return;
  if (res.amount == null || res.amount <= 0) return alertDlg('金额不对', '请输入大于 0 的金额');
  await run(() => L.addTemplate(res), { okText: '模板已保存' });
  render();
}

boot().catch((e) => { $app.innerHTML = `<div class="boot">打开失败：${esc(e.message ?? e)}</div>`; });
