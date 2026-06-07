/* Popup: configure the Odoo connection and look up a reference. */

const $ = (id) => document.getElementById(id);

function setStatus(text, ok) {
  const el = $('status');
  el.textContent = text || '';
  el.className = ok === true ? 'ok' : ok === false ? 'err' : '';
}

async function load() {
  const { odoo_url, odoo_token } = await chrome.storage.local.get(['odoo_url', 'odoo_token']);
  $('url').value = odoo_url || '';
  $('token').value = odoo_token || '';
}

$('save').addEventListener('click', async () => {
  const odoo_url = $('url').value.trim().replace(/\/+$/, '');
  const odoo_token = $('token').value.trim();
  if (!odoo_url || !odoo_token) { setStatus('URL and API key are required.', false); return; }
  await chrome.storage.local.set({ odoo_url, odoo_token });
  setStatus('Saved.', true);
});

$('test').addEventListener('click', () => {
  setStatus('Testing…');
  chrome.runtime.sendMessage({ type: 'ping' }, (resp) => {
    if (resp && resp.ok) setStatus('Connection OK.', true);
    else setStatus((resp && resp.error) || 'Connection failed.', false);
  });
});

$('lookup').addEventListener('click', () => {
  const reference = $('ref').value.trim();
  $('result').innerHTML = '';
  if (!reference) { setStatus('Enter a reference.', false); return; }
  setStatus('Searching…');
  chrome.runtime.sendMessage({ type: 'resolveReference', reference }, (resp) => {
    if (resp && resp.record) {
      const r = resp.record;
      const div = document.createElement('div');
      div.className = 'result';
      const a = document.createElement('a');
      a.href = r.url; a.target = '_blank'; a.rel = 'noopener';
      a.textContent = r.ref + ' — ' + r.name;
      div.appendChild(document.createTextNode((r.type || 'record').toUpperCase() + ': '));
      div.appendChild(a);
      $('result').appendChild(div);
      setStatus('', true);
    } else if (resp && resp.error) {
      setStatus(resp.error, false);
    } else {
      setStatus('No match for "' + reference + '".', false);
    }
  });
});

load();
