import { escapeHtml as esc } from './core/util.js';
export { esc };

const modalRoot = () => document.getElementById('modal-root');

/** 底部弹窗。buttons: [{ label, value, kind }]，返回所点按钮的 value；点遮罩返回 null */
export function sheet({ title = '', message = '', html = '', buttons = [{ label: '好', value: true, kind: 'primary' }], onMount = null }) {
  return new Promise((resolve) => {
    const wrap = document.createElement('div');
    wrap.className = 'mask';
    wrap.innerHTML = `<div class="sheet" role="dialog" aria-modal="true">
      ${title ? `<h2>${esc(title)}</h2>` : ''}${message ? `<p>${esc(message)}</p>` : ''}${html}
      <div class="sheet-btns">${buttons.map((b, i) => `<button class="btn ${b.kind ?? ''}" data-i="${i}">${esc(b.label)}</button>`).join('')}</div></div>`;
    const done = (v) => { wrap.remove(); resolve(v); };
    wrap.addEventListener('click', (e) => {
      if (e.target === wrap) return done(null);
      const b = e.target.closest('button[data-i]');
      if (b) {
        const btn = buttons[+b.dataset.i];
        const v = typeof btn.value === 'function' ? btn.value(wrap) : btn.value;
        if (v !== undefined) done(v);
      }
    });
    modalRoot().appendChild(wrap);
    onMount?.(wrap);
  });
}
export const confirmDlg = (title, message, okLabel = '确定', danger = false) =>
  sheet({ title, message, buttons: [{ label: okLabel, value: true, kind: danger ? 'danger' : 'primary' }, { label: '取消', value: false }] }).then((v) => v === true);
export const alertDlg = (title, message) => sheet({ title, message });
export function promptDlg(title, { value = '', placeholder = '', okLabel = '确定', type = 'text', inputmode = '' } = {}) {
  return sheet({ title, html: `<input class="input" id="dlg-input" type="${type}" ${inputmode ? `inputmode="${inputmode}"` : ''} value="${esc(value)}" placeholder="${esc(placeholder)}"><div class="hint"></div>`,
    buttons: [{ label: okLabel, kind: 'primary', value: (w) => w.querySelector('#dlg-input').value }, { label: '取消', value: null }],
    onMount: (w) => setTimeout(() => w.querySelector('#dlg-input')?.focus(), 50) });
}

let toastTimer = null;
export function toast(text, action = null) {
  const root = document.getElementById('toast-root');
  root.innerHTML = `<div class="toast"><span>${esc(text)}</span>${action ? `<button>${esc(action.label)}</button>` : ''}</div>`;
  if (action) root.querySelector('button').onclick = () => { root.innerHTML = ''; action.run(); };
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { root.innerHTML = ''; }, action ? 6000 : 2500);
}

/** PIN 键盘（渲染到容器），onDone(pin) 返回 Promise<string|null>：返回字符串表示错误提示，null 表示成功 */
export function pinPad(container, { title, sub = '', onDone, onCancel = null }) {
  let pin = '';
  const render = (msg = sub, err = false) => {
    container.innerHTML = `<div class="pin-wrap"><div class="pin-title">${esc(title)}</div>
      <div class="pin-sub ${err ? 'err-text' : ''}">${esc(msg)}</div>
      <div class="dots">${[0, 1, 2, 3, 4, 5].map((i) => `<span class="dot ${i < pin.length ? 'on' : ''}"></span>`).join('')}</div>
      <div class="pad">${[1, 2, 3, 4, 5, 6, 7, 8, 9].map((n) => `<button data-k="${n}">${n}</button>`).join('')}
      <button class="fn" data-k="cancel">${onCancel ? '取消' : ''}</button><button data-k="0">0</button><button class="fn" data-k="del">删除</button></div></div>`;
  };
  render();
  container.onclick = async (e) => {
    const b = e.target.closest('button[data-k]'); if (!b) return;
    const k = b.dataset.k;
    if (k === 'del') pin = pin.slice(0, -1);
    else if (k === 'cancel') { if (onCancel) onCancel(); return; }
    else if (pin.length < 6) pin += k;
    render();
    if (pin.length === 6) {
      const p = pin; pin = '';
      const err = await onDone(p);
      if (err) render(err, true);
    }
  };
}

/** 弹出 PIN 验证，返回是否通过 */
export function askPin(title, check) {
  return new Promise((resolve) => {
    const wrap = document.createElement('div');
    wrap.className = 'full';
    modalRoot().appendChild(wrap);
    pinPad(wrap, { title, sub: '输入 6 位密码', onCancel: () => { wrap.remove(); resolve(false); },
      onDone: async (p) => { if (await check(p)) { wrap.remove(); resolve(true); return null; } return '密码不对，请重试'; } });
  });
}
