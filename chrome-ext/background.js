/*
 * MV3 service worker. All Odoo network calls happen here (host_permissions
 * bypass CORS in the worker), so content scripts and the popup talk to Odoo
 * only through chrome.runtime messages.
 *
 * Message protocol:
 *   { type: 'ping' }                              -> { ok } | { error }
 *   { type: 'resolveReference', reference }       -> { record } | { error }
 *   { type: 'linkExisting', resModel, resId, recordName, ids } -> { result } | { error }
 *   { type: 'linkedRecords', gmailMessageId, gmailThreadId }   -> { records } | { error }
 */

importScripts('lib/odoo.js');

chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
  (async () => {
    try {
      switch (msg && msg.type) {
        case 'ping':
          await KONU_ODOO.ping();
          sendResponse({ ok: true });
          break;
        case 'resolveReference':
          sendResponse({ record: await KONU_ODOO.resolveReference(msg.reference) });
          break;
        case 'linkExisting':
          sendResponse({ result: await KONU_ODOO.linkExisting(msg.resModel, msg.resId, msg.recordName, msg.ids) });
          break;
        case 'linkedRecords':
          sendResponse({ records: (await KONU_ODOO.linkedRecords(msg.gmailMessageId, msg.gmailThreadId)).records || [] });
          break;
        default:
          sendResponse({ error: 'Unknown message type' });
      }
    } catch (err) {
      sendResponse({ error: String((err && err.message) || err) });
    }
  })();
  return true; // keep the channel open for the async response
});
