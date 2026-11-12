// 完整备份：ZIP（backup.json + photos/*）+ 校验；恢复前严格校验
import { canonical, sha256Hex, deepEqual } from './util.js';
import { zipStore, unzipStore } from './zip.js';
import { Ledger, replay } from './ledger.js';

export const BACKUP_FORMAT = 'ledger-pwa-backup';
export const BACKUP_VERSION = 1;

/** photos: [{ id, mime, bytes: Uint8Array, createdAt }] */
export async function makeBackup(ledger, photos = [], { now = new Date(), appVersion = '' } = {}) {
  const payload = { state: ledger.state, ops: ledger.ops };
  const photoMeta = [];
  const files = [];
  for (const ph of [...photos].sort((a, b) => a.id.localeCompare(b.id))) {
    const name = `photos/${ph.id}`;
    photoMeta.push({ id: ph.id, mime: ph.mime, createdAt: ph.createdAt, size: ph.bytes.length, sha256: await sha256Hex(ph.bytes), file: name });
    files.push({ name, data: ph.bytes });
  }
  const meta = { format: BACKUP_FORMAT, version: BACKUP_VERSION, exportedAt: now.toISOString(), appVersion,
    deviceId: ledger.deviceId, counts: { entries: Object.keys(ledger.state.entries).length, ops: ledger.ops.length, photos: photoMeta.length },
    checksum: await sha256Hex(canonical(payload)), photos: photoMeta };
  const json = new TextEncoder().encode(JSON.stringify({ meta, ...payload }));
  return zipStore([{ name: 'backup.json', data: json }, ...files]);
}

/** 解析并校验备份；成功返回 { meta, state, ops, photos }，失败抛错（不改动任何数据） */
export async function readBackup(bytes) {
  const files = unzipStore(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes));
  const jf = files.find((f) => f.name === 'backup.json');
  if (!jf) throw new Error('备份里没有 backup.json');
  let obj;
  try { obj = JSON.parse(new TextDecoder().decode(jf.data)); } catch { throw new Error('backup.json 无法解析'); }
  const { meta, state, ops } = obj;
  if (meta?.format !== BACKUP_FORMAT) throw new Error('不是本 App 的备份文件');
  if (meta.version > BACKUP_VERSION) throw new Error('备份来自更新的版本，请先更新 App');
  if (await sha256Hex(canonical({ state, ops })) !== meta.checksum) throw new Error('数据校验和不一致，备份可能已损坏');
  if (!deepEqual(replay(ops), state)) throw new Error('操作日志与数据不一致，备份可能已损坏');
  const l = new Ledger({ state, ops });
  const issues = l.verifyIntegrity();
  if (issues.length) throw new Error('数据自检失败：' + issues.slice(0, 3).join('；'));
  const photos = [];
  for (const pm of meta.photos ?? []) {
    const f = files.find((x) => x.name === pm.file);
    if (!f) throw new Error(`缺少凭证图片 ${pm.id}`);
    if (await sha256Hex(f.data) !== pm.sha256) throw new Error(`凭证图片校验失败 ${pm.id}`);
    photos.push({ id: pm.id, mime: pm.mime, createdAt: pm.createdAt, bytes: f.data });
  }
  return { meta, state, ops, photos };
}

/** 所有被引用的凭证 id（含已删除流水与历史版本里的，保证回退后凭证仍在） */
export function referencedPhotoIds(ledger) {
  const ids = new Set();
  for (const e of Object.values(ledger.state.entries)) for (const p of e.photoIds ?? []) ids.add(p);
  for (const op of ledger.ops) for (const c of op.changes) for (const x of [c.before, c.after]) for (const p of x?.photoIds ?? []) ids.add(p);
  return ids;
}
