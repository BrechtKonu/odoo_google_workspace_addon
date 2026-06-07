/*
 * Shared Odoo JSON-RPC client for the Chrome extension service worker.
 *
 * Mirrors the Apps Script add-on's transport: every endpoint is a
 * `type='jsonrpc', auth='outlook'` route authenticated with the user's Odoo
 * API key as a Bearer token (the standard mail_plugin flow). Loaded into the
 * MV3 service worker via importScripts(); never imported into a content script
 * (page-origin fetch would hit CORS — the worker has host_permissions instead).
 */

const KONU_ODOO = {
  async config() {
    const { odoo_url, odoo_token } = await chrome.storage.local.get(['odoo_url', 'odoo_token']);
    return { baseUrl: (odoo_url || '').replace(/\/+$/, ''), token: odoo_token || '' };
  },

  isConfigured(cfg) {
    return !!(cfg && cfg.baseUrl && cfg.token);
  },

  /** Call an add-on JSON-RPC route. Returns the `result` object or throws. */
  async call(path, params) {
    const cfg = await this.config();
    if (!this.isConfigured(cfg)) {
      throw new Error('Not configured — set the Odoo URL and API key in the extension popup.');
    }
    const url = cfg.baseUrl + path;
    const resp = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + cfg.token,
      },
      body: JSON.stringify({ jsonrpc: '2.0', method: 'call', params: params || {} }),
    });
    if (!resp.ok) {
      throw new Error('Odoo returned HTTP ' + resp.status);
    }
    const data = await resp.json();
    if (data.error) {
      throw new Error((data.error.data && data.error.data.message) || data.error.message || 'Odoo error');
    }
    return data.result;
  },

  // Convenience wrappers over the same endpoints the add-on uses.
  ping() { return this.call('/gmail_addon/ping', {}); },
  taskSearch(term, limit) { return this.call('/gmail_addon/task/search', { search_term: term, limit: limit || 5 }); },
  ticketSearch(term, limit) { return this.call('/gmail_addon/ticket/search', { search_term: term, limit: limit || 5 }); },
  leadSearch(term, limit) { return this.call('/gmail_addon/lead/search', { search_term: term, limit: limit || 5 }); },
  linkedRecords(gmailMessageId, gmailThreadId) {
    return this.call('/gmail_addon/email/linked_records', {
      gmail_message_id: gmailMessageId || '', gmail_thread_id: gmailThreadId || '',
    });
  },
  linkExisting(resModel, resId, recordName, ids) {
    return this.call('/gmail_addon/email/link_record', Object.assign(
      { res_model: resModel, res_id: resId, record_name: recordName || '' }, ids || {}));
  },

  /** Resolve a typed reference to the top task, then ticket, then lead. */
  async resolveReference(reference) {
    const term = String(reference || '').trim();
    if (!term) return null;
    const task = await this.taskSearch(term, 1);
    if (task && task.tasks && task.tasks.length) {
      const t = task.tasks[0];
      return { type: 'task', res_model: 'project.task', id: t.id, name: t.name, ref: t.reference || t.task_number || ('#' + t.id), url: t.url };
    }
    const ticket = await this.ticketSearch(term, 1);
    if (ticket && ticket.tickets && ticket.tickets.length) {
      const t = ticket.tickets[0];
      return { type: 'ticket', res_model: 'helpdesk.ticket', id: t.id, name: t.name, ref: t.reference || t.ticket_ref || ('#' + t.id), url: t.url };
    }
    const lead = await this.leadSearch(term, 1);
    if (lead && lead.leads && lead.leads.length) {
      const t = lead.leads[0];
      return { type: 'lead', res_model: 'crm.lead', id: t.id, name: t.name, ref: t.reference || t.lead_ref || ('#' + t.id), url: t.url };
    }
    return null;
  },
};
