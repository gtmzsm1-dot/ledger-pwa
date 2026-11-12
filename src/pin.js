// PIN：PBKDF2-SHA256 加盐哈希存本机，不存明文
const ITER = 150000;
const hex = (b) => [...new Uint8Array(b)].map((x) => x.toString(16).padStart(2, '0')).join('');
async function derive(pin, saltHex, iter = ITER) {
  const salt = new Uint8Array(saltHex.match(/../g).map((h) => parseInt(h, 16)));
  const key = await crypto.subtle.importKey('raw', new TextEncoder().encode(pin), 'PBKDF2', false, ['deriveBits']);
  return hex(await crypto.subtle.deriveBits({ name: 'PBKDF2', hash: 'SHA-256', salt, iterations: iter }, key, 256));
}
export async function makePinRecord(pin) {
  const salt = hex(crypto.getRandomValues(new Uint8Array(16)));
  return { salt, iter: ITER, hash: await derive(pin, salt) };
}
export async function checkPin(pin, rec) {
  if (!rec) return false;
  return (await derive(pin, rec.salt, rec.iter)) === rec.hash;
}
export const isValidPin = (p) => /^\d{6}$/.test(p);
