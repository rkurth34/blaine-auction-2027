// Shared data layer: every page reads and saves through the Apps Script web app over the sheet.
// The team passcode is asked once per browser and kept in localStorage.
import { API_URL } from './config.js';

const KEY = 'blaine27.passcode';
let gate = null;

function stored() { try { return localStorage.getItem(KEY) || ''; } catch (e) { return ''; } }
function remember(v) { try { localStorage.setItem(KEY, v); } catch (e) {} }
function forget() { try { localStorage.removeItem(KEY); } catch (e) {} }

function askPasscode(wrong) {
  if (gate) return gate;
  gate = new Promise(resolve => {
    const wrap = document.createElement('div');
    wrap.className = 'gate';
    wrap.innerHTML = `<form class="gate-box"><label for="gatePass">Team passcode</label>
      <input id="gatePass" type="password" autocomplete="current-password" required>
      ${wrong ? '<p class="gate-err">That passcode didn’t work.</p>' : ''}
      <button type="submit">Open</button></form>`;
    document.body.append(wrap);
    const input = wrap.querySelector('input');
    input.focus();
    wrap.querySelector('form').addEventListener('submit', e => {
      e.preventDefault();
      const v = input.value.trim();
      if (!v) return;
      remember(v); wrap.remove(); gate = null; resolve(v);
    });
  });
  return gate;
}

export async function call(action, payload = {}) {
  let key = stored() || await askPasscode(false);
  for (let tries = 0; tries < 3; tries++) {
    const res = await fetch(API_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ ...payload, action, key }),
      redirect: 'follow',
    });
    if (!res.ok) throw new Error('Server error ' + res.status);
    const data = await res.json();
    if (data.error === 'passcode') { forget(); key = await askPasscode(true); continue; }
    if (data.error) throw new Error(data.error);
    return data;
  }
  throw new Error('passcode');
}

export const read = tabs => call('read', { tabs }).then(d => d.tabs);
export const upsert = (tab, row) => call('upsert', { tab, row }).then(d => d.row);
export const patch = (tab, id, data) => call('patch', { tab, id, data }).then(d => d.row);
export const remove = (tab, id) => call('remove', { tab, id });
export const claim = (tab, id, asker, by) => call('claim', { tab, id, asker, by });
export const setting = (name, value) => call('setting', { name, value });
export const readOrg = () => call('readOrg');
export const saveOrg = chart => call('saveOrg', { chart });

// Re-read every `ms` while the tab is visible, so edits from teammates show up.
export function poll(fn, ms = 45000) {
  let t = setInterval(() => { if (!document.hidden) fn(); }, ms);
  document.addEventListener('visibilitychange', () => { if (!document.hidden) fn(); });
  return () => clearInterval(t);
}
