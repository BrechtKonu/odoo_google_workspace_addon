/*
 * Gmail content script (phase 1, read/augment only).
 *
 * This is the thing the Apps Script add-on structurally CANNOT do: touch the
 * message body DOM. It scans the open message for Odoo-style references
 * (KOTASK-053, LATR.PS-002, LATR.HT-2095) and turns them into clickable chips
 * that resolve the live record via the service worker and open it in Odoo.
 *
 * All CSS selectors that depend on Gmail's (unstable) DOM are isolated in
 * SELECTORS below, so a Gmail redesign is a one-place fix.
 */
(function () {
  'use strict';

  const SELECTORS = {
    // Gmail renders the open message body in elements with role="listitem";
    // the actual body is the `.a3s` container. Kept here intentionally.
    messageBody: 'div.a3s',
  };

  // Matches PREFIX-123 and PREFIX.SUB-123 (e.g. KOTASK-053, LATR.PS-002).
  const REF_RE = /\b([A-Z][A-Z0-9]{1,9}(?:\.[A-Z]{1,4})?-\d{1,7})\b/g;
  const MARK = 'data-konu-odoo';

  function decorate(node) {
    if (node.getAttribute(MARK)) return;
    node.setAttribute(MARK, '1');
    const html = node.innerHTML;
    if (!REF_RE.test(html)) return;
    REF_RE.lastIndex = 0;
    // Only rewrite text nodes to avoid corrupting links/attributes.
    const walker = document.createTreeWalker(node, NodeFilter.SHOW_TEXT, null);
    const targets = [];
    let n;
    while ((n = walker.nextNode())) {
      if (REF_RE.test(n.nodeValue)) targets.push(n);
      REF_RE.lastIndex = 0;
    }
    targets.forEach((textNode) => {
      const frag = document.createDocumentFragment();
      let last = 0;
      const text = textNode.nodeValue;
      let m;
      REF_RE.lastIndex = 0;
      while ((m = REF_RE.exec(text))) {
        if (m.index > last) frag.appendChild(document.createTextNode(text.slice(last, m.index)));
        frag.appendChild(makeChip(m[1]));
        last = m.index + m[0].length;
      }
      if (last < text.length) frag.appendChild(document.createTextNode(text.slice(last)));
      textNode.parentNode.replaceChild(frag, textNode);
    });
  }

  function makeChip(ref) {
    const chip = document.createElement('span');
    chip.className = 'konu-odoo-chip';
    chip.textContent = ref;
    chip.title = 'Open ' + ref + ' in Odoo';
    chip.addEventListener('click', (ev) => {
      ev.preventDefault();
      ev.stopPropagation();
      chip.classList.add('konu-odoo-chip--loading');
      chrome.runtime.sendMessage({ type: 'resolveReference', reference: ref }, (resp) => {
        chip.classList.remove('konu-odoo-chip--loading');
        if (resp && resp.record && resp.record.url) {
          window.open(resp.record.url, '_blank', 'noopener');
        } else if (resp && resp.error) {
          chip.title = resp.error;
          chip.classList.add('konu-odoo-chip--error');
        } else {
          chip.title = 'No matching Odoo record for ' + ref;
          chip.classList.add('konu-odoo-chip--error');
        }
      });
    });
    return chip;
  }

  function scan() {
    document.querySelectorAll(SELECTORS.messageBody).forEach(decorate);
  }

  // Gmail is a SPA; observe DOM mutations and rescan (debounced).
  let pending = null;
  const observer = new MutationObserver(() => {
    if (pending) return;
    pending = setTimeout(() => { pending = null; scan(); }, 400);
  });
  observer.observe(document.body, { childList: true, subtree: true });
  scan();
})();
