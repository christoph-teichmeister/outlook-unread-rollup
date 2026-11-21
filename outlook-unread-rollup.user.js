// ==UserScript==
// @name         Outlook: Unread Folder Totals
// @namespace    https://github.com/christoph-teichmeister/outlook-unread-rollup
// @version      0.3
// @description  Collapsed parent folders show own + subfolder unread total (Outlook web hides subfolder counts).
// @match        https://outlook.office.com/mail/*
// @match        https://outlook.office365.com/mail/*
// @match        https://outlook.cloud.microsoft/mail/*
// @icon         https://raw.githubusercontent.com/christoph-teichmeister/outlook-unread-rollup/refs/heads/main/icon.svg
// @grant        none
// @updateURL    https://raw.githubusercontent.com/christoph-teichmeister/outlook-unread-rollup/main/outlook-unread-rollup.user.js
// @downloadURL  https://raw.githubusercontent.com/christoph-teichmeister/outlook-unread-rollup/main/outlook-unread-rollup.user.js
// ==/UserScript==
(() => {
  'use strict';
  const API = 'https://outlook.office.com/api/v2.0/me/MailFolders';
  let totals = {}; // "Inbox/ai-portal" -> {own, total}

  // MSAL cache in localStorage holds an outlook.office.com token with Mail scopes (OWA's own login).
  function token() {
    for (let i = 0; i < localStorage.length; i++) {
      const k = localStorage.key(i);
      if (!/accesstoken/i.test(k)) continue;
      try {
        const v = JSON.parse(localStorage.getItem(k));
        if ((v.target || '').includes('outlook.office.com/Mail.') && v.expiresOn * 1000 > Date.now()) return v.secret;
      } catch (e) { /* ignore */ }
    }
  }

  async function list(id, tok) {
    const u = `${API}/${id}/childfolders?$select=DisplayName,UnreadItemCount,ChildFolderCount&$top=250`;
    const r = await fetch(u, { headers: { Authorization: 'Bearer ' + tok } });
    return r.ok ? (await r.json()).value || [] : [];
  }

  async function walk(id, path, tok, out) {
    let sum = 0;
    for (const f of await list(id, tok)) {
      const p = path ? `${path}/${f.DisplayName}` : f.DisplayName;
      const sub = f.ChildFolderCount ? await walk(f.Id, p, tok, out) : 0;
      out[p] = { own: f.UnreadItemCount, total: f.UnreadItemCount + sub };
      sum += f.UnreadItemCount + sub;
    }
    return sum;
  }

  async function refresh() {
    const tok = token();
    if (!tok) return;
    const out = {};
    // ponytail: ChildFolderCount of msgfolderroot unknown, walk it directly
    await walk('msgfolderroot', '', tok, out);
    totals = out;
    apply();
  }

  const nameOf = (item) => item.querySelector(':scope > span')?.textContent.trim();
  const nativeBadge = (item) => item.querySelectorAll(':scope > span:not([data-rollup])')[1];
  const isNum = (el) => el && /^\d+/.test(el.textContent.trim());
  let ref = null; // copy of a native badge, so our badge gets exactly the same classes/look

  function apply() {
    const items = [...document.querySelectorAll('[role=treeitem]')];
    const stack = []; // names by level
    let inAccount = false;
    for (const it of items) {
      const lvl = +it.getAttribute('aria-level');
      const name = nameOf(it);
      if (lvl === 1) { inAccount = (it.innerText || '').includes('@'); stack.length = 0; continue; }
      if (!inAccount || !name) continue;
      const nb = nativeBadge(it);
      if (isNum(nb) && nb.style.display !== 'none') ref = nb.cloneNode(true);
      stack.length = lvl - 1; stack[lvl - 2] = name; // level 2 = top-level folder
      const t = totals[stack.join('/')];
      if (!t) continue;
      const collapsed = it.getAttribute('aria-expanded') === 'false';
      let mark = it.querySelector(':scope > span[data-rollup]');
      if (collapsed && t.total > t.own) {
        if (!mark) {
          const base = isNum(nb) ? nb : ref;
          if (base) {
            mark = base.cloneNode(true);
            mark.style.display = '';
          } else {
            mark = document.createElement('span');
            mark.style.cssText = 'margin-left:auto;margin-right:8px';
          }
          mark.dataset.rollup = '1';
          it.insertBefore(mark, it.querySelector(':scope > span').nextSibling);
        }
        const num = [...mark.querySelectorAll('span')].find((x) => !x.children.length && /^\d+$/.test(x.textContent.trim())) || mark;
        if (num.textContent !== String(t.total)) num.textContent = t.total;
        if (nb) nb.style.display = 'none';
      } else if (mark) {
        mark.remove();
        if (nb) nb.style.display = '';
      }
    }
  }

  let timer;
  new MutationObserver(() => { clearTimeout(timer); timer = setTimeout(apply, 250); })
    .observe(document.body, { childList: true, subtree: true, attributes: true, attributeFilter: ['aria-expanded'] });
  setInterval(refresh, 90000);
  document.addEventListener('visibilitychange', () => { if (!document.hidden) refresh(); });
  setTimeout(refresh, 3000);
})();
