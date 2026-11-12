// IndexedDB 封装：kv（账本、PIN、设置）、photos（凭证图片）、snapshots（本地自动备份，轮转保留）
const DB_NAME = 'ledger-pwa';
const DB_VERSION = 1;
export const SNAPSHOT_KEEP = 20;
let dbp = null;

function open() {
  if (dbp) return dbp;
  dbp = new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains('kv')) db.createObjectStore('kv');
      if (!db.objectStoreNames.contains('photos')) db.createObjectStore('photos', { keyPath: 'id' });
      if (!db.objectStoreNames.contains('snapshots')) db.createObjectStore('snapshots', { keyPath: 'id' });
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
  return dbp;
}
function tx(stores, mode, fn) {
  return open().then((db) => new Promise((resolve, reject) => {
    const t = db.transaction(stores, mode);
    let result;
    Promise.resolve(fn(t)).then((r) => { result = r; });
    t.oncomplete = () => resolve(result);
    t.onerror = () => reject(t.error);
    t.onabort = () => reject(t.error ?? new Error('写入被中止'));
  }));
}
const req2p = (r) => new Promise((res, rej) => { r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); });

export const kvGet = (k) => tx(['kv'], 'readonly', (t) => req2p(t.objectStore('kv').get(k)));
export const kvSet = (k, v) => tx(['kv'], 'readwrite', (t) => { t.objectStore('kv').put(v, k); });

/** 保存账本：一次事务写入，写入后读回核对条数 */
export async function saveLedger(state, ops) {
  await tx(['kv'], 'readwrite', (t) => { t.objectStore('kv').put({ state, ops, savedAt: new Date().toISOString() }, 'ledger'); });
}
export const loadLedger = () => kvGet('ledger');

export const putPhoto = (p) => tx(['photos'], 'readwrite', (t) => { t.objectStore('photos').put(p); });
export const getPhoto = (id) => tx(['photos'], 'readonly', (t) => req2p(t.objectStore('photos').get(id)));
export const allPhotos = () => tx(['photos'], 'readonly', (t) => req2p(t.objectStore('photos').getAll()));
export async function replacePhotos(list) {
  await tx(['photos'], 'readwrite', (t) => { const s = t.objectStore('photos'); s.clear(); for (const p of list) s.put(p); });
}

export async function addSnapshot(reason, state, ops) {
  const id = new Date().toISOString() + '-' + Math.random().toString(36).slice(2, 6);
  await tx(['snapshots'], 'readwrite', async (t) => {
    const s = t.objectStore('snapshots');
    s.put({ id, createdAt: new Date().toISOString(), reason, state, ops, entries: Object.keys(state.entries).length });
  });
  const all = await listSnapshots();
  if (all.length > SNAPSHOT_KEEP) {
    const drop = all.slice(SNAPSHOT_KEEP);
    await tx(['snapshots'], 'readwrite', (t) => { for (const d of drop) t.objectStore('snapshots').delete(d.id); });
  }
  return id;
}
export async function listSnapshots() {
  const all = await tx(['snapshots'], 'readonly', (t) => req2p(t.objectStore('snapshots').getAll()));
  return all.sort((a, b) => b.id.localeCompare(a.id));
}
export const getSnapshot = (id) => tx(['snapshots'], 'readonly', (t) => req2p(t.objectStore('snapshots').get(id)));
