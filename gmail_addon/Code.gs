// ============================================================
// Odoo Tasks & Tickets — Gmail Add-on
// Google Apps Script (V8)
// ============================================================

// ─── AUTH & CONFIG ───────────────────────────────────────────────────────────

var USER_PROPS = PropertiesService.getUserProperties();

// Batch-load all user properties once per execution and cache the result.
var _propsCache = null;
function getProps_() {
  if (!_propsCache) _propsCache = USER_PROPS.getProperties();
  return _propsCache;
}

function getOdooUrl_() {
  return (getProps_().odoo_url || '').replace(/\/$/, '');
}

function getToken_() {
  return getProps_().odoo_token || '';
}

function isConfigured_() {
  var p = getProps_();
  return !!(p.odoo_url && p.odoo_token);
}

function getDriveFolderId_() {
  return getProps_().drive_folder_id || '';
}

function saveAuth_(url, token) {
  USER_PROPS.setProperty('odoo_url', url.replace(/\/$/, ''));
  USER_PROPS.setProperty('odoo_token', token);
  _propsCache = null;
}

// ─── ODOO HTTP CLIENT ────────────────────────────────────────────────────────

function buildOdooRequestOptionsWithAuth_(baseUrl, token, path, params) {
  var normalizedBaseUrl = String(baseUrl || '').replace(/\/$/, '');
  return {
    url: normalizedBaseUrl + path,
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + String(token || '') },
    payload: JSON.stringify({ jsonrpc: '2.0', method: 'call', id: 1, params: params || {} }),
    muteHttpExceptions: true
  };
}

function buildOdooRequestOptions_(path, params) {
  return buildOdooRequestOptionsWithAuth_(getOdooUrl_(), getToken_(), path, params);
}

function parseOdooResponse_(response, path) {
  var code = response.getResponseCode();
  if (code !== 200) throw new Error('HTTP ' + code + ' from ' + path);
  var body = JSON.parse(response.getContentText());
  if (body.error) {
    var msg = (body.error.data && body.error.data.message) || body.error.message || JSON.stringify(body.error);
    throw new Error(msg);
  }
  return body.result;
}

/**
 * POST to an Odoo JSON-RPC endpoint with Bearer auth.
 * Returns the parsed `result` field, or throws on error.
 */
function odooPost_(path, params) {
  var req = buildOdooRequestOptions_(path, params);
  var response = UrlFetchApp.fetch(req.url, req);
  return parseOdooResponse_(response, path);
}

/**
 * Run multiple Odoo requests in parallel using UrlFetchApp.fetchAll.
 * Each item: { path, params }. Returns results in the same order; null on error.
 */
function odooFetchAll_(requests) {
  var options = requests.map(function(r) { return buildOdooRequestOptions_(r.path, r.params); });
  var responses = UrlFetchApp.fetchAll(options);
  return responses.map(function(resp, i) {
    try { return parseOdooResponse_(resp, requests[i].path); } catch (_) { return null; }
  });
}

function testOdooConnection_(url, token) {
  var path = '/gmail_addon/suggest_context';
  var req = buildOdooRequestOptionsWithAuth_(url, token, path, { sender_email: '' });
  var response = UrlFetchApp.fetch(req.url, req);
  parseOdooResponse_(response, path);
}

// ─── ODOO API WRAPPERS ───────────────────────────────────────────────────────

function apiSuggestContext_(senderEmail) {
  if (!senderEmail) return odooPost_('/gmail_addon/suggest_context', { sender_email: '' });
  var cache = CacheService.getUserCache();
  var key = 'ctx_' + senderEmail;
  var hit = cache.get(key);
  if (hit) { try { return JSON.parse(hit); } catch (_) {} }
  var result = odooPost_('/gmail_addon/suggest_context', { sender_email: senderEmail });
  try { cache.put(key, JSON.stringify(result), 300); } catch (_) {}
  return result;
}

function apiTaskSearch_(params) {
  return odooPost_('/gmail_addon/task/search', params);
}

function apiTicketSearch_(params) {
  return odooPost_('/gmail_addon/ticket/search', params);
}

function apiLeadSearch_(params) {
  return odooPost_('/gmail_addon/lead/search', params);
}

function apiProjectDropdown_() {
  var cache = CacheService.getUserCache();
  var key = 'dd_project_v1';
  var hit = cache.get(key);
  if (hit) { try { return JSON.parse(hit); } catch (_) {} }
  var result = odooPost_('/gmail_addon/project/dropdown', {});
  try { cache.put(key, JSON.stringify(result), 180); } catch (_) {}
  return result;
}

function apiStageDropdown_(projectId, teamId) {
  var pid = projectId || null;
  var tid = teamId || null;
  var cache = CacheService.getUserCache();
  var key = 'dd_stage_v1_' + String(pid || '') + '_' + String(tid || '');
  var hit = cache.get(key);
  if (hit) { try { return JSON.parse(hit); } catch (_) {} }
  var result = odooPost_('/gmail_addon/stage/dropdown', {
    project_id: projectId || null,
    team_id: teamId || null
  });
  try { cache.put(key, JSON.stringify(result), 120); } catch (_) {}
  return result;
}

function apiTeamDropdown_() {
  var cache = CacheService.getUserCache();
  var key = 'dd_team_v1';
  var hit = cache.get(key);
  if (hit) { try { return JSON.parse(hit); } catch (_) {} }
  var result = odooPost_('/gmail_addon/team/dropdown', {});
  try { cache.put(key, JSON.stringify(result), 180); } catch (_) {}
  return result;
}

function apiCrmTeamDropdown_() {
  var cache = CacheService.getUserCache();
  var key = 'dd_crm_team_v1';
  var hit = cache.get(key);
  if (hit) { try { return JSON.parse(hit); } catch (_) {} }
  var result = odooPost_('/gmail_addon/crm/team/dropdown', {});
  try { cache.put(key, JSON.stringify(result), 180); } catch (_) {}
  return result;
}

function apiCrmStageDropdown_(teamId) {
  var cache = CacheService.getUserCache();
  var key = 'dd_crm_stage_v1_' + String(teamId || '');
  var hit = cache.get(key);
  if (hit) { try { return JSON.parse(hit); } catch (_) {} }
  var result = odooPost_('/gmail_addon/crm/stage/dropdown', { team_id: teamId || null });
  try { cache.put(key, JSON.stringify(result), 120); } catch (_) {}
  return result;
}

function apiUserDropdown_() {
  var cache = CacheService.getUserCache();
  var key = 'dd_user_v1';
  var hit = cache.get(key);
  if (hit) { try { return JSON.parse(hit); } catch (_) {} }
  var result = odooPost_('/gmail_addon/user/dropdown', {});
  try { cache.put(key, JSON.stringify(result), 180); } catch (_) {}
  return result;
}

function apiPartnerAutocomplete_(term) {
  return odooPost_('/gmail_addon/partner/autocomplete', { search_term: term });
}

function apiCreatePartner_(name, email, companyName) {
  return odooPost_('/gmail_addon/partner/create', { name: name, email: email, company_name: companyName || '' });
}

function apiFormSchema_(recordType) {
  var cache = CacheService.getUserCache();
  var key = 'form_schema_' + String(recordType || '');
  var hit = cache.get(key);
  if (hit) { try { return JSON.parse(hit); } catch (_) {} }
  var result = odooPost_('/gmail_addon/form/schema', { record_type: recordType });
  try { cache.put(key, JSON.stringify(result), 180); } catch (_) {}
  return result;
}

function apiCreateTask_(params) {
  return odooPost_('/gmail_addon/task/create', params);
}

function apiCreateTicket_(params) {
  return odooPost_('/gmail_addon/ticket/create', params);
}

function apiCreateLead_(params) {
  return odooPost_('/gmail_addon/lead/create', params);
}

function apiLogEmail_(params) {
  return odooPost_('/gmail_addon/log_email', params);
}

function apiLinkedRecords_(rfcMessageId, gmailMessageId, gmailThreadId) {
  return odooPost_('/gmail_addon/email/linked_records', {
    rfc_message_id: rfcMessageId || '',
    gmail_message_id: gmailMessageId || '',
    gmail_thread_id: gmailThreadId || ''
  });
}

function apiDocumentLinkedRecords_(documentId, hostApp) {
  return odooPost_('/gmail_addon/document/linked_records', {
    document_id: documentId || '',
    host_app: hostApp || '',
    limit: 20
  });
}

function apiDocumentLinkRecord_(documentId, hostApp, resModel, resId, recordName) {
  return odooPost_('/gmail_addon/document/link_record', {
    document_id: documentId || '',
    host_app: hostApp || '',
    res_model: resModel || '',
    res_id: resId || 0,
    record_name: recordName || ''
  });
}

// ─── COMPOSE ACTION HELPERS ──────────────────────────────────────────────────

/** Minimal HTML escaping for inline text. */
function escapeHtml_(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function getSingleFormValue_(form, name) {
  var value = form ? form[name] : '';
  if (Array.isArray(value)) return value.length ? value[0] : '';
  return value || '';
}

function getMultiFormValues_(form, name) {
  var value = form ? form[name] : [];
  if (value === undefined || value === null || value === '') return [];
  return Array.isArray(value) ? value : [value];
}

function getRecordRef_(rec) {
  if (!rec) return '';
  return rec.reference || rec.task_number || rec.ticket_ref || rec.lead_ref || ('#' + rec.id);
}

function getRecordTypeLabel_(rec) {
  if (!rec) return 'Record';
  if (rec.type === 'task') return 'Task';
  if (rec.type === 'ticket') return 'Ticket';
  if (rec.type === 'lead') return rec.lead_type === 'opportunity' || rec.type_label === 'Opportunity'
    ? 'Opportunity'
    : 'Lead';
  return 'Record';
}

function getRecordKnownIcon_(rec) {
  if (rec && rec.type === 'task') return CardService.Icon.DESCRIPTION;
  if (rec && rec.type === 'ticket') return CardService.Icon.CONFIRMATION_NUMBER_ICON;
  return CardService.Icon.PERSON;
}

function addDynamicFieldWidgets_(section, schema) {
  ((schema && schema.extra_fields) || []).forEach(function(field) {
    var fieldName = 'extra__' + field.name;
    var requiredSuffix = field.required ? ' *' : '';
    if (field.type === 'selection' || field.type === 'many2one') {
      section.addWidget(CardService.newTextParagraph().setText((field.label || field.name) + requiredSuffix));
      var dropdown = CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.DROPDOWN)
        .setFieldName(fieldName)
        .addItem('Select...', '', !field.required);
      (field.options || []).forEach(function(option) {
        dropdown.addItem(option.label, String(option.value), false);
      });
      section.addWidget(dropdown);
      return;
    }

    if (field.type === 'many2many') {
      section.addWidget(CardService.newTextParagraph().setText((field.label || field.name) + requiredSuffix));
      var multi = CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.CHECK_BOX)
        .setFieldName(fieldName);
      (field.options || []).forEach(function(option) {
        multi.addItem(option.label, String(option.value), false);
      });
      section.addWidget(multi);
      return;
    }

    if (field.type === 'boolean') {
      section.addWidget(CardService.newTextParagraph().setText(field.label || field.name));
      section.addWidget(CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.CHECK_BOX)
        .setFieldName(fieldName)
        .addItem(field.label || field.name, 'true', false)
      );
      return;
    }

    section.addWidget(CardService.newTextInput()
      .setFieldName(fieldName)
      .setTitle((field.label || field.name) + requiredSuffix)
      .setMultiline(field.type === 'text' || field.type === 'html')
      .setHint(field.help || '')
    );
  });
}

function extractDynamicFieldValues_(form, schema) {
  var extraValues = {};
  ((schema && schema.extra_fields) || []).forEach(function(field) {
    var fieldName = 'extra__' + field.name;
    if (field.type === 'many2many') {
      var manyValues = getMultiFormValues_(form, fieldName);
      if (manyValues.length) extraValues[field.name] = manyValues;
      return;
    }
    if (field.type === 'boolean') {
      extraValues[field.name] = getMultiFormValues_(form, fieldName).indexOf('true') >= 0;
      return;
    }
    var value = getSingleFormValue_(form, fieldName);
    if (value !== '') extraValues[field.name] = value;
  });
  return extraValues;
}

/**
 * Returns a ButtonSet with a reply button and an insert-link button.
 * ref – display text, e.g. "KOTASK-001 · Task Name · Assignee"
 * url – Odoo record URL
 */
function buildComposeButtons_(ref, url) {
  var action = CardService.newAction()
    .setFunctionName('onInsertReferenceInEmail')
    .setParameters({ reference: ref, url: url });
  return CardService.newButtonSet()
    .addButton(CardService.newImageButton()
      .setIconUrl('https://www.gstatic.com/images/icons/material/system/1x/reply_black_18dp.png')
      .setAltText('Reply with reference')
      .setComposeAction(action, CardService.ComposedEmailType.REPLY_AS_NEW_EMAIL)
    );
}

/**
 * Compose action: inserts a styled smart chip linking to the Odoo record.
 * Used by both reply and standalone-draft compose buttons.
 */
function onInsertReferenceInEmail(e) {
  var params = e.parameters;
  var reference = params.reference;
  var url = params.url;
  var safeUrl = escapeHtml_(url);
  var chipHtml =
    '<p><a href="' + safeUrl + '" style="background:#e8f0fe;border-radius:12px;' +
    'color:#1a73e8;display:inline-block;font-family:Google Sans,Roboto,sans-serif;' +
    'font-size:13px;font-weight:500;padding:3px 10px;text-decoration:none;">' +
    escapeHtml_(reference) + '</a></p><p><br></p>';
  var messageId = e.gmail && e.gmail.messageId;
  var draft = messageId
    ? GmailApp.getMessageById(messageId).createDraftReplyAll('', { htmlBody: chipHtml })
    : GmailApp.createDraft('', '', '', { htmlBody: chipHtml });
  return CardService.newComposeActionResponseBuilder().setGmailDraft(draft).build();
}

/**
 * UpdateDraftAction: inserts the chip at cursor position (or end of draft).
 * Only valid inside a compose trigger context.
 */
function onInsertAtCursor(e) {
  var params = e.parameters;
  var safeUrl = escapeHtml_(params.url);
  var chipHtml =
    '<p><a href="' + safeUrl + '" style="background:#e8f0fe;border-radius:12px;' +
    'color:#1a73e8;display:inline-block;font-family:Google Sans,Roboto,sans-serif;' +
    'font-size:13px;font-weight:500;padding:3px 10px;text-decoration:none;">' +
    escapeHtml_(params.reference) + '</a></p><p><br></p>';
  var updateAction = CardService.newUpdateDraftBodyAction()
    .addUpdateContent(chipHtml, CardService.ContentType.MUTABLE_HTML);
  return CardService.newUpdateDraftActionResponseBuilder()
    .setUpdateDraftBodyAction(updateAction)
    .build();
}

// ─── EMAIL CONTEXT ───────────────────────────────────────────────────────────

/**
 * Extracts bare email addresses from an RFC 2822 address list string.
 * "John Smith <john@example.com>, jane@example.com" → "john@example.com, jane@example.com"
 */
function extractEmailsFromAddressList_(str) {
  if (!str) return '';
  var emails = [];
  var re = /[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/g;
  var m;
  while ((m = re.exec(str)) !== null) {
    emails.push(m[0].toLowerCase());
  }
  return emails.join(', ');
}

/**
 * Parses an RFC 2822 address list into [{email, label}] entries.
 * label is the display name if present, otherwise the bare email.
 */
function extractParticipantList_(str) {
  if (!str) return [];
  var results = [];
  var seen = {};
  // Match "Display Name <email>" or bare "email"
  var re = /(?:([^<,]+?)\s*<([a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})>|([a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}))/g;
  var m;
  while ((m = re.exec(str)) !== null) {
    var email = (m[2] || m[3] || '').toLowerCase();
    var name = (m[1] || '').trim().replace(/^"|"$/g, '');
    if (email && !seen[email]) {
      seen[email] = true;
      results.push({ email: email, name: name, label: name ? name + ' (' + email + ')' : email });
    }
  }
  return results;
}

function getEmailContext_(e) {
  try {
    var messageId = e && e.gmail && e.gmail.messageId;
    if (!messageId) return {};
    var msg = GmailApp.getMessageById(messageId);
    if (!msg) return {};
    var from = msg.getFrom() || '';
    var emailMatch = from.match(/<([^>]+)>/) || from.match(/(\S+@\S+)/);
    var senderEmail = emailMatch ? emailMatch[1] : from;
    var senderName = from.replace(/<[^>]+>/, '').trim().replace(/^"|"$/g, '') || senderEmail;
    var plainBody = msg.getPlainBody() || '';
    var rfcMessageId = '';
    try { rfcMessageId = msg.getHeader('Message-ID') || ''; } catch (_) {}
    return {
      senderEmail: senderEmail,
      senderName: senderName,
      subject: msg.getSubject() || '',
      plainBody: plainBody.substring(0, 2000),
      to: msg.getTo() || '',
      cc: msg.getCc() || '',
      messageId: messageId,
      threadId: (e && e.gmail && e.gmail.threadId) || '',
      rfcMessageId: rfcMessageId
    };
  } catch (err) {
    return {};
  }
}

// Clean email HTML for Odoo chatter: strip cid: images, keep external URLs
function cleanEmailHtml_(html) {
  if (!html) return '';
  // Remove <script> and <style> blocks
  html = html.replace(/<script[\s\S]*?<\/script>/gi, '');
  html = html.replace(/<style[\s\S]*?<\/style>/gi, '');
  // Replace cid: images with placeholder
  html = html.replace(/src=["']cid:[^"']*["']/gi, 'src="" alt="[embedded image]"');
  // Strip base64 inline images (too large for JSON)
  html = html.replace(/src=["']data:image[^"']*["']/gi, 'src="" alt="[inline image]"');
  // Strip tracking pixels: <img> tags that are 1x1 or have width/height=1
  html = html.replace(/<img[^>]*(width=["']?1["']?|height=["']?1["']?)[^>]*>/gi, '');
  return html;
}

// ─── LOGIN CARD ──────────────────────────────────────────────────────────────

function buildLoginCard_() {
  var section = CardService.newCardSection()
    .setHeader('Connect to Odoo')
    .addWidget(CardService.newTextParagraph().setText(
      'Enter your Odoo URL and API key to get started.'
    ))
    .addWidget(CardService.newTextInput()
      .setFieldName('odoo_url')
      .setTitle('Odoo URL')
      .setHint('https://yourcompany.odoo.com')
      .setValue(getOdooUrl_())
    )
    .addWidget(CardService.newTextInput()
      .setFieldName('odoo_token')
      .setTitle('API Key')
      .setHint('Your Odoo API key (from user preferences)')
      .setValue(getToken_())
    )
    .addWidget(CardService.newTextInput()
      .setFieldName('drive_folder_id')
      .setTitle('Drive Folder ID for Attachments')
      .setHint('Paste a Google Drive folder ID (find it in the folder URL)')
      .setValue(getDriveFolderId_())
    );

  if (!getDriveFolderId_()) {
    section.addWidget(CardService.newDecoratedText()
      .setText('⚠️ No Drive folder ID set — attachments will not be saved.')
    );
  }

  section
    .addWidget(CardService.newTextButton()
      .setText('Save & Connect')
      .setOnClickAction(CardService.newAction().setFunctionName('onSaveAuth'))
    );

  var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
  var authStatus = authInfo.getAuthorizationStatus();
  var authSection = CardService.newCardSection().setHeader('Google Permissions');

  if (authStatus === ScriptApp.AuthorizationStatus.REQUIRED) {
    authSection.addWidget(CardService.newDecoratedText()
      .setText('⚠️ Some Google permissions are missing or revoked.')
    );
  } else {
    authSection.addWidget(CardService.newDecoratedText()
      .setText('Google permissions are up to date.')
    );
  }

  var authUrl = authInfo.getAuthorizationUrl();
  if (authUrl) {
    authSection.addWidget(CardService.newTextButton()
      .setText('Re-authorize Google Permissions')
      .setOpenLink(CardService.newOpenLink()
        .setUrl(authUrl)
        .setOpenAs(CardService.OpenAs.FULL_SIZE)
        .setOnClose(CardService.OnClose.RELOAD_ADD_ON)
      )
    );
  }

  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Odoo Tasks & Tickets'))
    .addSection(section)
    .addSection(authSection)
    .build();
}

function onSaveAuth(e) {
  var url = (e.formInput.odoo_url || '').trim();
  var token = (e.formInput.odoo_token || '').trim();
  var folderId = (e.formInput.drive_folder_id || '').trim();
  if (!url || !token) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Please enter both URL and API key.'))
      .build();
  }

  // Validate Drive folder ID if provided
  if (folderId) {
    try {
      DriveApp.getFolderById(folderId);
    } catch (err) {
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText('Drive folder ID is invalid or you do not have access to it.'))
        .build();
    }
  }

  // Test Odoo connection
  try {
    testOdooConnection_(url, token);
  } catch (err) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Connection test failed: ' + err.message))
      .build();
  }

  // Persist settings only after all checks pass.
  saveAuth_(url, token);
  USER_PROPS.setProperty('drive_folder_id', folderId);
  _propsCache = null;

  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText('Connected to Odoo!'))
    .setNavigation(CardService.newNavigation().popToRoot())
    .build();
}

// ─── HOME CARD ───────────────────────────────────────────────────────────────

function onGmailMessageOpen(e) {
  if (!isConfigured_()) return buildLoginCard_();
  return buildHomeCard_(e);
}

// ─── COMPOSE HOME CARD ───────────────────────────────────────────────────────

/**
 * Strip common reply/forward prefixes so the same base subject matches across
 * the contextual trigger (reading) and the compose trigger (replying).
 */
function normalizeSubject_(subject) {
  return (subject || '')
    .replace(/^((Re|AW|FW|Fwd|TR|WG|SV|VS):\s*)+/gi, '')
    .trim()
    .substring(0, 200);
}

function onGmailCompose(e) {
  if (!isConfigured_()) return buildLoginCard_();
  return buildComposeHomeCard_(e);
}

function buildComposeHomeCard_(e) {
  // Compose trigger events never expose e.gmail.threadId, even with
  // draftAccess:"METADATA". Recover it from the cache seeded by the
  // contextual trigger (buildHomeCard_) using the normalised subject as key.
  var threadId = null;
  var composeSubject = (e && e.gmail && e.gmail.subject) || '';
  if (composeSubject) {
    var normalizedSubj = normalizeSubject_(composeSubject);
    if (normalizedSubj) {
      try { threadId = CacheService.getUserCache().get('tid_subj_' + normalizedSubj); } catch (_) {}
    }
  }

  // Try to extract sender email from the thread (needed for suggest_context).
  var senderEmail = '';
  if (threadId) {
    try {
      var thread = GmailApp.getThreadById(threadId);
      if (thread) {
        var messages = thread.getMessages();
        if (messages && messages.length > 0) {
          var from = messages[0].getFrom() || '';
          var emailMatch = from.match(/<([^>]+)>/) || from.match(/(\S+@\S+)/);
          senderEmail = emailMatch ? emailMatch[1] : from;
        }
      }
    } catch (_) {}
  }

  // Run suggest_context + linked_records in parallel (only what's needed).
  var context = null;
  var linkedRecords = [];

  if (threadId || senderEmail) {
    var pendingRequests = [];
    var idxContext = -1, idxLinked = -1;

    if (senderEmail) {
      var cache = CacheService.getUserCache();
      var ctxCacheKey = 'ctx_' + senderEmail;
      try { var hit = cache.get(ctxCacheKey); if (hit) context = JSON.parse(hit); } catch (_) {}
      if (!context) {
        idxContext = pendingRequests.length;
        pendingRequests.push({ path: '/gmail_addon/suggest_context', params: { sender_email: senderEmail } });
      }
    }

    if (threadId) {
      idxLinked = pendingRequests.length;
      pendingRequests.push({ path: '/gmail_addon/email/linked_records', params: {
        rfc_message_id: '',
        gmail_message_id: '',
        gmail_thread_id: threadId
      }});
    }

    if (pendingRequests.length > 0) {
      var fetchResults = odooFetchAll_(pendingRequests);
      if (idxContext >= 0 && fetchResults[idxContext]) {
        context = fetchResults[idxContext];
        try { CacheService.getUserCache().put('ctx_' + senderEmail, JSON.stringify(context), 300); } catch (_) {}
      }
      if (idxLinked >= 0 && fetchResults[idxLinked]) {
        linkedRecords = fetchResults[idxLinked].records || [];
      }
    }
  }

  var suggestedProjectId = String((context && context.suggested_project_id) || '');
  var suggestedTeamId    = String((context && context.suggested_team_id)    || '');

  var cardBuilder = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Odoo Tasks & Tickets'));

  // Linked records section — quick insert without searching
  if (linkedRecords.length > 0) {
    var linkedSection = CardService.newCardSection().setHeader('Linked to this thread');
    linkedRecords.forEach(function(rec) {
      var refId = getRecordRef_(rec);
      var assignee = rec.user_name ? ' · ' + rec.user_name : '';
      var refStr = refId + ' · ' + rec.name + assignee;
      linkedSection.addWidget(CardService.newDecoratedText()
        .setTopLabel(getRecordTypeLabel_(rec) + ' ' + refId)
        .setText(rec.name)
        .setBottomLabel(rec.stage || '')
        .setStartIcon(CardService.newIconImage().setIcon(getRecordKnownIcon_(rec)))
        .setOpenLink(CardService.newOpenLink().setUrl(rec.url))
        .setButton(CardService.newImageButton()
          .setIconUrl('https://www.gstatic.com/images/icons/material/system/1x/insert_link_black_18dp.png')
          .setAltText('Insert reference at cursor')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('onInsertAtCursor')
            .setParameters({ reference: refStr, url: rec.url })
          )
        )
      );
    });
    cardBuilder.addSection(linkedSection);
  }

  // Search / navigate section
  var navSection = CardService.newCardSection();
  if (linkedRecords.length === 0 && threadId) {
    navSection.addWidget(CardService.newTextParagraph()
      .setText('No linked tasks, tickets, or leads found for this thread.')
    );
  }
  navSection.addWidget(CardService.newButtonSet()
    .addButton(CardService.newTextButton()
      .setText('Search Tasks')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('onNavigateToTaskSearch')
        .setParameters({
          compose_ctx: 'true',
          suggested_project_id: suggestedProjectId,
          sender_email: senderEmail
        })
      )
    )
    .addButton(CardService.newTextButton()
      .setText('Search Tickets')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('onNavigateToTicketSearch')
        .setParameters({
          compose_ctx: 'true',
          suggested_team_id: suggestedTeamId,
          sender_email: senderEmail
        })
      )
    )
    .addButton(CardService.newTextButton()
      .setText('Search Leads')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('onNavigateToLeadSearch')
        .setParameters({
          compose_ctx: 'true',
          suggested_team_id: String((context && context.suggested_crm_team_id) || ''),
          sender_email: senderEmail
        })
      )
    )
  );
  cardBuilder.addSection(navSection);

  return cardBuilder.build();
}

// ─── HOME CARD (contextual) ──────────────────────────────────────────────────

function buildHomeCard_(e, overrideSenderEmail) {
  var ctx = getEmailContext_(e);
  var senderEmail = overrideSenderEmail || ctx.senderEmail || '';

  // Check suggest_context cache before deciding which requests to make.
  var context = null;
  var linkedRecords = [];
  var hasLinkedIds = !!(ctx.rfcMessageId || ctx.messageId || ctx.threadId);
  var cache = CacheService.getUserCache();

  // Cache threadId by normalised subject so the compose trigger can look it up
  // (compose events do not expose e.gmail.threadId).
  if (ctx.threadId && ctx.subject) {
    var normalizedSubj = normalizeSubject_(ctx.subject);
    if (normalizedSubj) {
      try { cache.put('tid_subj_' + normalizedSubj, ctx.threadId, 21600); } catch (_) {}
    }
  }

  var filterMine = getProps_().filter_mine_records === '1';
  var ctxCacheKey = senderEmail ? 'ctx_' + senderEmail + (filterMine ? '_mine' : '') : null;
  if (ctxCacheKey) {
    try { var hit = cache.get(ctxCacheKey); if (hit) context = JSON.parse(hit); } catch (_) {}
  }

  // Build only the requests that are still needed, then run them in parallel.
  var pendingRequests = [];
  var idxContext = -1, idxLinked = -1;
  if (!context && senderEmail) {
    idxContext = pendingRequests.length;
    pendingRequests.push({ path: '/gmail_addon/suggest_context', params: { sender_email: senderEmail, filter_mine: filterMine } });
  }
  if (hasLinkedIds) {
    idxLinked = pendingRequests.length;
    pendingRequests.push({ path: '/gmail_addon/email/linked_records', params: {
      rfc_message_id: ctx.rfcMessageId || '',
      gmail_message_id: ctx.messageId || '',
      gmail_thread_id: ctx.threadId || ''
    }});
  }
  if (pendingRequests.length > 0) {
    var fetchResults = odooFetchAll_(pendingRequests);
    if (idxContext >= 0 && fetchResults[idxContext]) {
      context = fetchResults[idxContext];
      try { cache.put(ctxCacheKey, JSON.stringify(context), 300); } catch (_) {}
    }
    if (idxLinked >= 0 && fetchResults[idxLinked]) {
      linkedRecords = fetchResults[idxLinked].records || [];
    }
  }

  var headerSection = CardService.newCardSection();

  // Build participant dropdown (From + To + CC), falling back to static display when < 2 contacts.
  var fromEntry = ctx.senderEmail
    ? [{ email: ctx.senderEmail, name: ctx.senderName || '', label: ctx.senderName && ctx.senderName !== ctx.senderEmail
        ? ctx.senderName + ' (' + ctx.senderEmail + ')' : ctx.senderEmail }]
    : [];
  var toEntries = extractParticipantList_(ctx.to || '');
  var ccEntries = extractParticipantList_(ctx.cc || '');

  // Merge, deduplicating by email; From always comes first.
  var seenEmails = {};
  var participants = [];
  fromEntry.concat(toEntries).concat(ccEntries).forEach(function(p) {
    if (!seenEmails[p.email]) {
      seenEmails[p.email] = true;
      participants.push(p);
    }
  });

  if (participants.length >= 2) {
    var contactInput = CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.DROPDOWN)
      .setFieldName('contact_email')
      .setTitle('Contact');
    participants.forEach(function(p) {
      contactInput.addItem(p.label, p.email, p.email === senderEmail);
    });
    contactInput.setOnChangeAction(CardService.newAction().setFunctionName('onChangeSender'));
    headerSection.addWidget(contactInput);
    if (context && context.partner_name) {
      var partnerUrl = getOdooUrl_() + '/odoo/contacts/' + context.partner_id;
      headerSection.addWidget(CardService.newDecoratedText()
        .setTopLabel('Odoo contact')
        .setText(context.partner_name)
        .setBottomLabel(context.partner_email || senderEmail)
        .setStartIcon(CardService.newIconImage()
          .setIconUrl('https://lh3.googleusercontent.com/d/1j4I8gHrJxKXsULly-7U16lZrBjJ6QRVM')
          .setAltText('Linked to Odoo')
        )
        .setOpenLink(CardService.newOpenLink().setUrl(partnerUrl))
      );
    }
  } else if (context && context.partner_name) {
    headerSection.addWidget(CardService.newDecoratedText()
      .setTopLabel('Sender')
      .setText(context.partner_name)
      .setBottomLabel(context.partner_email || senderEmail)
      .setStartIcon(CardService.newIconImage().setIcon(CardService.Icon.PERSON))
    );
  } else if (senderEmail) {
    headerSection.addWidget(CardService.newDecoratedText()
      .setTopLabel('Sender')
      .setText(ctx.senderName || senderEmail)
      .setBottomLabel(senderEmail)
    );
  } else {
    headerSection.addWidget(CardService.newTextParagraph().setText('Open an email to see sender details.'));
  }

  if (senderEmail && !(context && context.partner_name)) {
    var selectedParticipant = participants.filter(function(p) { return p.email === senderEmail; })[0];
    var selectedName = (selectedParticipant && selectedParticipant.name) || ctx.senderName || '';
    var emailDomain = senderEmail.indexOf('@') >= 0 ? senderEmail.split('@')[1] : '';
    headerSection.addWidget(CardService.newDecoratedText()
      .setText('Create contact in Odoo')
      .setStartIcon(CardService.newIconImage()
        .setIconUrl('https://lh3.googleusercontent.com/d/1r1-qK_eKIH3N1763KyV4ScHQ0l6z8hD1')
        .setAltText('Create contact')
      )
      .setOnClickAction(CardService.newAction()
        .setFunctionName('onNavigateToCreateContact')
        .setParameters({
          sender_email: senderEmail,
          sender_name: selectedName,
          email_domain: emailDomain
        })
      )
    );
  }

  // Linked records section (only shown when links exist)
  var linkedSection = null;
  if (linkedRecords.length > 0) {
    linkedSection = CardService.newCardSection().setHeader('Linked to this email');
    linkedRecords.forEach(function(rec) {
      var refId = getRecordRef_(rec);
      var topLabel = getRecordTypeLabel_(rec) + ' ' + refId;
      linkedSection.addWidget(CardService.newDecoratedText()
        .setTopLabel(topLabel)
        .setText(rec.name)
        .setBottomLabel(rec.stage || '')
        .setStartIcon(CardService.newIconImage().setIcon(getRecordKnownIcon_(rec)))
        .setOpenLink(CardService.newOpenLink().setUrl(rec.url))
      );
      var assignee = rec.user_name ? ' · ' + rec.user_name : '';
      var ref = refId + ' · ' + rec.name + assignee;
      linkedSection.addWidget(buildComposeButtons_(ref, rec.url));
    });
  }

  // Navigation buttons
  var navSection = CardService.newCardSection().setHeader('Actions');

  navSection.addWidget(CardService.newButtonSet()
    .addButton(CardService.newTextButton()
      .setText('Search Tasks')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('onNavigateToTaskSearch')
        .setParameters({
          suggested_project_id: String((context && context.suggested_project_id) || ''),
          sender_email: senderEmail,
          sender_name: ctx.senderName || ''
        })
      )
    )
    .addButton(CardService.newTextButton()
      .setText('Search Tickets')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('onNavigateToTicketSearch')
        .setParameters({
          suggested_team_id: String((context && context.suggested_team_id) || ''),
          sender_email: senderEmail,
          sender_name: ctx.senderName || ''
        })
      )
    )
    .addButton(CardService.newTextButton()
      .setText('Search Leads')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('onNavigateToLeadSearch')
        .setParameters({
          suggested_team_id: String((context && context.suggested_crm_team_id) || ''),
          sender_email: senderEmail,
          sender_name: ctx.senderName || ''
        })
      )
    )
  );

  navSection.addWidget(CardService.newButtonSet()
    .addButton(CardService.newTextButton()
      .setText('New Task')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('onNavigateToCreateTask')
        .setParameters({
          subject: ctx.subject || '',
          plain_body: ctx.plainBody || '',
          cc: ctx.cc || '',
          sender_email: senderEmail,
          sender_name: ctx.senderName || '',
          suggested_project_id: String((context && context.suggested_project_id) || ''),
          suggested_project_name: (context && context.suggested_project_name) || ''
        })
      )
    )
    .addButton(CardService.newTextButton()
      .setText('New Ticket')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('onNavigateToCreateTicket')
        .setParameters({
          subject: ctx.subject || '',
          plain_body: ctx.plainBody || '',
          cc: ctx.cc || '',
          sender_email: senderEmail,
          sender_name: ctx.senderName || '',
          suggested_team_id: String((context && context.suggested_team_id) || ''),
          suggested_team_name: (context && context.suggested_team_name) || ''
        })
      )
    )
    .addButton(CardService.newTextButton()
      .setText('New Lead')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('onNavigateToCreateLead')
        .setParameters({
          subject: ctx.subject || '',
          plain_body: ctx.plainBody || '',
          cc: ctx.cc || '',
          sender_email: senderEmail,
          sender_name: ctx.senderName || '',
          suggested_team_id: String((context && context.suggested_crm_team_id) || '')
        })
      )
    )
  );

  navSection.addWidget(CardService.newTextButton()
    .setText('Settings')
    .setOnClickAction(CardService.newAction().setFunctionName('onNavigateToSettings'))
  );

  // Recent records – merged and sorted by write_date descending.
  // When no email is open (no senderEmail), fetch recent records globally.
  var recentItems = [];
  var showCompany = getProps_().show_company_records === '1';

  if (context && (context.recent_tasks || context.recent_tickets)) {
    if (context.recent_tasks) {
      context.recent_tasks.forEach(function(task) {
        recentItems.push({ type: 'task', data: task, write_date: task.write_date || '' });
      });
    }
    if (context.recent_tickets) {
      context.recent_tickets.forEach(function(ticket) {
        recentItems.push({ type: 'ticket', data: ticket, write_date: ticket.write_date || '' });
      });
    }
    if (context.recent_leads) {
      context.recent_leads.forEach(function(lead) {
        recentItems.push({ type: 'lead', data: lead, write_date: lead.write_date || '' });
      });
    }
    if (showCompany) {
      if (context.company_tasks) {
        context.company_tasks.forEach(function(task) {
          recentItems.push({ type: 'task', data: task, write_date: task.write_date || '' });
        });
      }
      if (context.company_tickets) {
        context.company_tickets.forEach(function(ticket) {
          recentItems.push({ type: 'ticket', data: ticket, write_date: ticket.write_date || '' });
        });
      }
      if (context.company_leads) {
        context.company_leads.forEach(function(lead) {
          recentItems.push({ type: 'lead', data: lead, write_date: lead.write_date || '' });
        });
      }
    }
  } else if (!senderEmail) {
    // No email open — fetch recent tasks and tickets in parallel
    var globalFetch = odooFetchAll_([
      { path: '/gmail_addon/task/search',   params: { limit: 5, offset: 0 } },
      { path: '/gmail_addon/ticket/search', params: { limit: 5, offset: 0 } },
      { path: '/gmail_addon/lead/search',   params: { limit: 5, offset: 0 } }
    ]);
    ((globalFetch[0] && globalFetch[0].tasks)   || []).forEach(function(task) {
      recentItems.push({ type: 'task', data: task, write_date: task.write_date || '' });
    });
    ((globalFetch[1] && globalFetch[1].tickets) || []).forEach(function(ticket) {
      recentItems.push({ type: 'ticket', data: ticket, write_date: ticket.write_date || '' });
    });
    ((globalFetch[2] && globalFetch[2].leads) || []).forEach(function(lead) {
      recentItems.push({ type: 'lead', data: lead, write_date: lead.write_date || '' });
    });
  }

  recentItems.sort(function(a, b) {
    if (a.write_date > b.write_date) return -1;
    if (a.write_date < b.write_date) return 1;
    return 0;
  });

  var shownRecent = recentItems.slice(0, 10);
  var recentHeader = senderEmail ? 'Recent for this contact' : 'Recent records';
  var recentSection = CardService.newCardSection().setHeader(recentHeader);

  if (senderEmail) {
    var filterCheckboxes = CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.CHECK_BOX)
      .setFieldName('recent_filters')
      .setOnChangeAction(CardService.newAction()
        .setFunctionName('onToggleRecentFilters')
        .setParameters({ sender_email: senderEmail })
      );
    if (context && context.company_partner_id) {
      filterCheckboxes.addItem('Include ' + context.company_partner_name, 'company', showCompany);
    }
    filterCheckboxes.addItem('Only mine', 'mine', filterMine);
    recentSection.addWidget(filterCheckboxes);
  }

  if (shownRecent.length === 0) {
    recentSection.addWidget(CardService.newTextParagraph().setText(
      senderEmail ? 'No recent records for this contact.' : 'No recent records found.'
    ));
  } else {
    shownRecent.forEach(function(item) {
      if (item.type === 'task') {
        var task = item.data;
        var taskWidget = CardService.newDecoratedText()
          .setTopLabel('[Task] ' + (task.task_number || ('#' + task.id)) + ' · ' + (task.project_name || ''))
          .setText(task.name)
          .setBottomLabel(task.stage_name || '')
          .setStartIcon(CardService.newIconImage().setIcon(CardService.Icon.DESCRIPTION))
          .setOpenLink(CardService.newOpenLink().setUrl(task.url));
        if (ctx.messageId) {
          taskWidget.setButton(CardService.newImageButton()
            .setIconUrl('https://www.gstatic.com/images/icons/material/system/1x/comment_grey600_18dp.png')
            .setAltText('Log email')
            .setOnClickAction(CardService.newAction()
              .setFunctionName('onNavigateToLogEmail')
              .setParameters({
                res_model: 'project.task',
                res_id: String(task.id),
                record_name: task.name,
                sender_email: senderEmail,
                subject: ctx.subject || ''
              })
            )
          );
        }
        recentSection.addWidget(taskWidget);
      } else if (item.type === 'ticket') {
        var ticket = item.data;
        var ticketWidget = CardService.newDecoratedText()
          .setTopLabel('[Ticket] ' + (ticket.ticket_ref || ('#' + ticket.id)) + ' · ' + (ticket.team_name || ''))
          .setText(ticket.name)
          .setBottomLabel(ticket.stage_name || '')
          .setStartIcon(CardService.newIconImage().setIcon(CardService.Icon.CONFIRMATION_NUMBER_ICON))
          .setOpenLink(CardService.newOpenLink().setUrl(ticket.url));
        if (ctx.messageId) {
          ticketWidget.setButton(CardService.newImageButton()
            .setIconUrl('https://www.gstatic.com/images/icons/material/system/1x/comment_grey600_18dp.png')
            .setAltText('Log email')
            .setOnClickAction(CardService.newAction()
              .setFunctionName('onNavigateToLogEmail')
              .setParameters({
                res_model: 'helpdesk.ticket',
                res_id: String(ticket.id),
                record_name: ticket.name,
                sender_email: senderEmail,
                subject: ctx.subject || ''
              })
            )
          );
        }
        recentSection.addWidget(ticketWidget);
      } else {
        var lead = item.data;
        var leadWidget = CardService.newDecoratedText()
          .setTopLabel('[' + (lead.type_label || 'Lead') + '] ' + getRecordRef_(lead) + ' · ' + (lead.team_name || ''))
          .setText(lead.name)
          .setBottomLabel((lead.stage_name || '') + (lead.user_name ? ' · ' + lead.user_name : ''))
          .setStartIcon(CardService.newIconImage().setIcon(CardService.Icon.PERSON))
          .setOpenLink(CardService.newOpenLink().setUrl(lead.url));
        if (ctx.messageId) {
          leadWidget.setButton(CardService.newImageButton()
            .setIconUrl('https://www.gstatic.com/images/icons/material/system/1x/comment_grey600_18dp.png')
            .setAltText('Log email')
            .setOnClickAction(CardService.newAction()
              .setFunctionName('onNavigateToLogEmail')
              .setParameters({
                res_model: 'crm.lead',
                res_id: String(lead.id),
                record_name: lead.name,
                sender_email: senderEmail,
                subject: ctx.subject || ''
              })
            )
          );
        }
        recentSection.addWidget(leadWidget);
      }
    });
  }

  var cardBuilder = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Odoo Tasks & Tickets'))
    .addSection(headerSection);
  if (linkedSection) cardBuilder.addSection(linkedSection);
  cardBuilder.addSection(navSection).addSection(recentSection);
  return cardBuilder.build();
}

function onNavigateToSettings(e) {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(buildLoginCard_()))
    .build();
}

function onChangeSender(e) {
  var selectedEmail = (e.formInput && e.formInput.contact_email) || '';
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(buildHomeCard_(e, selectedEmail)))
    .build();
}

function onToggleRecentFilters(e) {
  var selected = (e.formInput && e.formInput.recent_filters) || [];
  var showCompany = selected.indexOf('company') >= 0;
  var filterMine = selected.indexOf('mine') >= 0;
  USER_PROPS.setProperties({
    show_company_records: showCompany ? '1' : '0',
    filter_mine_records: filterMine ? '1' : '0'
  });
  _propsCache = null;
  var senderEmail = (e.parameters && e.parameters.sender_email) || '';
  // Invalidate both mine/non-mine context cache entries so the card re-fetches
  try {
    var cache = CacheService.getUserCache();
    if (senderEmail) {
      cache.remove('ctx_' + senderEmail);
      cache.remove('ctx_' + senderEmail + '_mine');
    }
  } catch (_) {}
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(buildHomeCard_(e, senderEmail)))
    .build();
}

// ─── CREATE CONTACT CARD ─────────────────────────────────────────────────────

function buildCreateContactCard_(params, suggestedCompanyName, companyNames) {
  params = params || {};
  var section = CardService.newCardSection().setHeader('New Odoo Contact');
  section.addWidget(CardService.newTextInput()
    .setFieldName('contact_name')
    .setTitle('Name *')
    .setValue(params.sender_name || '')
  );
  section.addWidget(CardService.newTextInput()
    .setFieldName('contact_email')
    .setTitle('Email *')
    .setValue(params.sender_email || '')
  );
  var companyInput = CardService.newTextInput()
    .setFieldName('company_name')
    .setTitle('Company (optional)')
    .setHint('Type to filter, or leave empty')
    .setValue(suggestedCompanyName || '');
  if (companyNames && companyNames.length > 0) {
    var suggestions = CardService.newSuggestions();
    companyNames.forEach(function(name) { suggestions.addSuggestion(name); });
    companyInput.setSuggestions(suggestions);
  }
  section.addWidget(companyInput);
  section.addWidget(CardService.newTextButton()
    .setText('Create Contact')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('onCreateContact')
      .setParameters({ original_sender_email: params.sender_email || '' })
    )
  );
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Create Contact'))
    .addSection(section)
    .build();
}


var _PERSONAL_DOMAINS = { 'gmail.com': 1, 'outlook.com': 1, 'hotmail.com': 1, 'yahoo.com': 1,
  'live.com': 1, 'icloud.com': 1, 'protonmail.com': 1, 'me.com': 1, 'msn.com': 1 };

function onNavigateToCreateContact(e) {
  var p = e.parameters || {};
  var domain = p.email_domain || '';
  var suggestedCompanyName = '';
  var companyNames = [];
  try {
    // Load top 30 companies by recent activity for the static suggestion list.
    var result = odooPost_('/gmail_addon/partner/autocomplete', { search_term: '', companies_only: true, limit: 30 });
    companyNames = (result.partners || []).map(function(c) { return c.name; });
    // Prefill with best domain match if the domain is not a personal provider.
    if (domain && !_PERSONAL_DOMAINS[domain]) {
      var domainMatch = (result.partners || []).filter(function(c) {
        return c.name.toLowerCase().indexOf(domain.split('.')[0].toLowerCase()) >= 0;
      })[0];
      if (domainMatch) suggestedCompanyName = domainMatch.name;
    }
  } catch (_) {}
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(
      buildCreateContactCard_(p, suggestedCompanyName, companyNames)
    ))
    .build();
}

function onCreateContact(e) {
  var form = e.formInput || {};
  var params = e.parameters || {};
  var name        = (form.contact_name || '').trim();
  var email       = (form.contact_email || '').trim();
  var companyName = (form.company_name || '').trim();
  if (!name || !email) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Name and email are required.'))
      .build();
  }
  try {
    var result = apiCreatePartner_(name, email, companyName);
    if (result.error) {
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText(result.error))
        .build();
    }
    var cacheEmail = params.original_sender_email || email;
    try { CacheService.getUserCache().remove('ctx_' + cacheEmail); } catch (_) {}
    var msg = result.already_exists
      ? 'Contact already exists: ' + result.partner_name
      : 'Contact created: ' + result.partner_name;
    // Rebuild the home card immediately so the green icon appears without a manual refresh.
    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().updateCard(buildHomeCard_(e, cacheEmail)))
      .setNotification(CardService.newNotification().setText(msg))
      .build();
  } catch (err) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Error: ' + err.message))
      .build();
  }
}

// ─── TASK SEARCH CARD ────────────────────────────────────────────────────────

function onNavigateToTaskSearch(e) {
  var p = e.parameters || {};
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(
      buildTaskSearchCard_({
        suggested_project_id: p.suggested_project_id || '',
        sender_email: p.sender_email || '',
        sender_name: p.sender_name || '',
        search_term: '',
        offset: 0,
        compose_ctx: p.compose_ctx || ''
      })
    ))
    .build();
}

function buildTaskSearchCard_(params) {
  params = params || {};
  var offset = parseInt(params.offset) || 0;
  var searchTerm = params.search_term || '';
  var projectId = params.project_id || params.suggested_project_id || '';
  var stageId = params.stage_id || '';
  var userId = params.user_id || '';

  // Load dropdowns in parallel
  var dropdownResults = odooFetchAll_([
    { path: '/gmail_addon/project/dropdown', params: {} },
    { path: '/gmail_addon/stage/dropdown', params: { project_id: projectId || null, team_id: null, record_type: 'task' } },
    { path: '/gmail_addon/user/dropdown', params: {} }
  ]);
  var projects = (dropdownResults[0] && dropdownResults[0].projects) || [];
  var stages = (dropdownResults[1] && dropdownResults[1].stages) || [];
  var users = (dropdownResults[2] && dropdownResults[2].users) || [];

  var formSection = CardService.newCardSection().setHeader('Search Tasks');

  formSection.addWidget(CardService.newTextInput()
    .setFieldName('search_term')
    .setTitle('Search')
    .setHint('Task name or number (e.g. PROJ-001)')
    .setValue(searchTerm)
  );

  // Project dropdown
  formSection.addWidget(CardService.newTextParagraph().setText('Project'));
  var projectInput = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setFieldName('project_id')
    .setOnChangeAction(CardService.newAction()
      .setFunctionName('onTaskProjectChanged')
      .setParameters({ offset: '0', search_term: searchTerm, stage_id: stageId,
                       sender_email: params.sender_email || '',
                       sender_name: params.sender_name || '',
                       compose_ctx: params.compose_ctx || '' })
    )
    .addItem('All projects', '', !projectId);
  projects.forEach(function(p) {
    projectInput.addItem(p.name, String(p.id), String(p.id) === String(projectId));
  });
  formSection.addWidget(projectInput);

  // Stage dropdown
  formSection.addWidget(CardService.newTextParagraph().setText('Stage'));
  var stageInput = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setFieldName('stage_id')
    .addItem('All stages', '', !stageId);
  stages.forEach(function(s) {
    stageInput.addItem(s.name, String(s.id), String(s.id) === String(stageId));
  });
  formSection.addWidget(stageInput);

  // Assigned user dropdown
  formSection.addWidget(CardService.newTextParagraph().setText('Assigned to'));
  var userInput = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setFieldName('user_id')
    .addItem('Anyone', '', !userId);
  users.forEach(function(u) {
    userInput.addItem(u.name, String(u.id), String(u.id) === String(userId));
  });
  formSection.addWidget(userInput);

  formSection.addWidget(CardService.newButtonSet()
    .addButton(CardService.newTextButton()
      .setText('Search')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('onTaskSearch')
        .setParameters({ offset: '0', sender_email: params.sender_email || '',
                         sender_name: params.sender_name || '',
                         compose_ctx: params.compose_ctx || '' })
      )
    )
    .addButton(CardService.newTextButton()
      .setText('+ New Task')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('onNavigateToCreateTask')
        .setParameters({
          subject: '',
          plain_body: '',
          cc: '',
          sender_email: params.sender_email || '',
          sender_name: params.sender_name || '',
          suggested_project_id: projectId,
          suggested_project_name: ''
        })
      )
    )
  );

  var cardBuilder = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Search Tasks'))
    .addSection(formSection);

  // Results section (if search was performed)
  if (params.results) {
    var results = params.results;
    var total = params.total || 0;
    var resultsSection = CardService.newCardSection()
      .setHeader('Results (' + total + ' found)');

    if (results.length === 0) {
      resultsSection.addWidget(CardService.newTextParagraph().setText('No tasks found.'));
    } else {
      results.forEach(function(task) {
        var taskRef = task.task_number || ('#' + task.id);
        var taskAssignee = task.user_name ? ' · ' + task.user_name : '';
        var taskRefStr = taskRef + ' · ' + task.name + taskAssignee;
        var taskWidget = CardService.newDecoratedText()
          .setTopLabel('[Task] ' + (task.task_number || ('#' + task.id)) + ' · ' + (task.project_name || ''))
          .setText(task.name)
          .setBottomLabel((task.stage_name || '') + (task.user_name ? ' · ' + task.user_name : ''))
          .setOpenLink(CardService.newOpenLink().setUrl(task.url));
        var taskButtonSet = CardService.newButtonSet();
        if (!params.compose_ctx) {
          var taskComposeAction = CardService.newAction()
            .setFunctionName('onInsertReferenceInEmail')
            .setParameters({ reference: taskRefStr, url: task.url });
          taskWidget.setButton(CardService.newImageButton()
            .setIconUrl('https://www.gstatic.com/images/icons/material/system/1x/comment_grey600_18dp.png')
            .setAltText('Log email')
            .setOnClickAction(CardService.newAction()
              .setFunctionName('onNavigateToLogEmail')
              .setParameters({
                res_model: 'project.task',
                res_id: String(task.id),
                record_name: task.name,
                sender_email: params.sender_email || '',
                subject: params.subject || ''
              })
            )
          );
          taskButtonSet.addButton(CardService.newImageButton()
            .setIconUrl('https://www.gstatic.com/images/icons/material/system/1x/reply_grey600_18dp.png')
            .setAltText('Reply with reference')
            .setComposeAction(taskComposeAction, CardService.ComposedEmailType.REPLY_AS_NEW_EMAIL)
          );
        } else {
          taskButtonSet.addButton(
            CardService.newImageButton()
              .setIconUrl('https://www.gstatic.com/images/icons/material/system/1x/insert_link_black_18dp.png')
              .setAltText('Insert reference at cursor')
              .setOnClickAction(CardService.newAction()
                .setFunctionName('onInsertAtCursor')
                .setParameters({ reference: taskRefStr, url: task.url })
              )
          );
        }
        resultsSection.addWidget(taskWidget);
        resultsSection.addWidget(taskButtonSet);
        resultsSection.addWidget(CardService.newDivider());
      });

      // Pagination
      var pageSize = 10;
      var paginationSet = CardService.newButtonSet();
      if (offset > 0) {
        paginationSet.addButton(CardService.newTextButton()
          .setText('◄ Prev')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('onTaskSearch')
            .setParameters({
              offset: String(Math.max(0, offset - pageSize)),
              search_term: searchTerm,
              project_id: projectId,
              stage_id: stageId,
              user_id: userId,
              sender_email: params.sender_email || '',
              sender_name: params.sender_name || '',
              compose_ctx: params.compose_ctx || ''
            })
          )
        );
      }
      if (offset + results.length < total) {
        paginationSet.addButton(CardService.newTextButton()
          .setText('Next ►')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('onTaskSearch')
            .setParameters({
              offset: String(offset + pageSize),
              search_term: searchTerm,
              project_id: projectId,
              stage_id: stageId,
              user_id: userId,
              sender_email: params.sender_email || '',
              sender_name: params.sender_name || '',
              compose_ctx: params.compose_ctx || ''
            })
          )
        );
      }
      resultsSection.addWidget(paginationSet);
    }
    cardBuilder.addSection(resultsSection);
  }

  return cardBuilder.build();
}

function onTaskSearch(e) {
  var form = e.formInput || {};
  var params = e.parameters || {};
  var searchTerm = form.search_term || params.search_term || '';
  var projectId = form.project_id || params.project_id || '';
  var stageId = form.stage_id || params.stage_id || '';
  var userId = form.user_id || params.user_id || '';
  var offset = parseInt(params.offset) || 0;

  var results = [];
  var total = 0;
  var errorMsg = '';
  try {
    var data = apiTaskSearch_({
      search_term: searchTerm,
      project_id: projectId ? parseInt(projectId) : null,
      stage_id: stageId ? parseInt(stageId) : null,
      user_id: userId ? parseInt(userId) : null,
      limit: 10,
      offset: offset
    });
    results = data.tasks || [];
    total = data.total || 0;
  } catch (err) {
    errorMsg = err.message;
  }

  var cardParams = {
    search_term: searchTerm,
    project_id: projectId,
    stage_id: stageId,
    user_id: userId,
    offset: offset,
    results: results,
    total: total,
    sender_email: params.sender_email || '',
    sender_name: params.sender_name || '',
    subject: params.subject || '',
    compose_ctx: params.compose_ctx || ''
  };

  var card = buildTaskSearchCard_(cardParams);
  var response = CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(card));
  if (errorMsg) {
    response.setNotification(CardService.newNotification().setText('Error: ' + errorMsg));
  }
  return response.build();
}

function onTaskProjectChanged(e) {
  var form = e.formInput || {};
  var params = e.parameters || {};
  var projectId = form.project_id || '';
  // Reload stages for the selected project, then update card
  return onTaskSearch(Object.assign({}, e, {
    formInput: form,
    parameters: Object.assign({}, params, { offset: '0', project_id: projectId, stage_id: '' })
  }));
}

// ─── CREATE TASK CARD ────────────────────────────────────────────────────────

function onNavigateToCreateTask(e) {
  var p = e.parameters || {};
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(buildTaskCreateCard_(p)))
    .build();
}

function buildTaskCreateCard_(params) {
  params = params || {};
  var projects = [];
  var schema = { extra_fields: [] };
  try { projects = (apiProjectDropdown_().projects) || []; } catch (e) {}
  try { schema = apiFormSchema_('task') || schema; } catch (e) {}

  var section = CardService.newCardSection().setHeader('New Task');

  // Project
  section.addWidget(CardService.newTextParagraph().setText('Project *'));
  var projectInput = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setFieldName('project_id')
    .addItem('Select a project...', '', false);
  projects.forEach(function(p) {
    var selected = String(p.id) === String(params.suggested_project_id || '');
    projectInput.addItem(p.name, String(p.id), selected);
  });
  section.addWidget(projectInput);

  // Task name
  section.addWidget(CardService.newTextInput()
    .setFieldName('task_name')
    .setTitle('Task Name *')
    .setValue(params.subject || '')
  );

  // Partner
  section.addWidget(CardService.newTextInput()
    .setFieldName('partner_email')
    .setTitle('Customer email')
    .setHint('email of the customer')
    .setValue(params.sender_email || '')
  );

  // Description
  section.addWidget(CardService.newTextInput()
    .setFieldName('description')
    .setTitle('Description')
    .setMultiline(true)
    .setValue(params.plain_body || '')
  );

  // CC / Followers
  section.addWidget(CardService.newTextInput()
    .setFieldName('cc_addresses')
    .setTitle('CC / Followers')
    .setHint('Comma-separated emails')
    .setValue(params.cc || '')
  );

  addDynamicFieldWidgets_(section, schema);

  section.addWidget(CardService.newTextButton()
    .setText('Create Task')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('onCreateTask')
      .setParameters({
        sender_email: params.sender_email || '',
        sender_name: params.sender_name || ''
      })
    )
  );

  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('New Task'))
    .addSection(section)
    .build();
}

function onCreateTask(e) {
  var form = e.formInput || {};
  var params = e.parameters || {};
  var projectId = form.project_id;
  var name = (form.task_name || '').trim();
  var ccAddresses = extractEmailsFromAddressList_(form.cc_addresses || '');

  if (!projectId || !name) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Project and Task Name are required.'))
      .build();
  }

  var schema = { extra_fields: [] };
  try { schema = apiFormSchema_('task') || schema; } catch (err) {}

  // Resolve partner by email
  var partnerId = null;
  var partnerEmail = (form.partner_email || '').trim();
  if (partnerEmail) {
    try {
      var matches = apiPartnerAutocomplete_(partnerEmail);
      if (matches.partners && matches.partners.length > 0) {
        partnerId = matches.partners[0].id;
      }
    } catch (err) {}
  }

  // Fetch email body and RFC Message-ID from the current email
  var emailBody = '';
  var rfcMessageId = '';
  var gmailMessageId = (e.gmail && e.gmail.messageId) || '';
  var gmailThreadId = (e.gmail && e.gmail.threadId) || '';
  if (gmailMessageId) {
    try {
      var msg = GmailApp.getMessageById(gmailMessageId);
      emailBody = cleanEmailHtml_(msg.getBody());
      try { rfcMessageId = msg.getHeader('Message-ID') || ''; } catch (_) {}
    } catch (err) {}
  }

  try {
    var result = apiCreateTask_({
      project_id: parseInt(projectId),
      name: name,
      partner_id: partnerId,
      description: form.description || '',
      cc_addresses: ccAddresses,
      extra_values: extractDynamicFieldValues_(form, schema),
      email_body: emailBody,
      email_subject: name,
      author_email: partnerEmail || (params.sender_email || ''),
      rfc_message_id: rfcMessageId,
      gmail_message_id: gmailMessageId,
      gmail_thread_id: gmailThreadId
    });

    var successSection = CardService.newCardSection()
      .addWidget(CardService.newTextParagraph().setText('Task created successfully!'))
      .addWidget(CardService.newTextButton()
        .setText('Open Task in Odoo')
        .setOpenLink(CardService.newOpenLink().setUrl(result.task_url))
      );

    // Offer Drive attachment upload as a separate action when a Drive folder is configured
    if (gmailMessageId && getDriveFolderId_()) {
      successSection.addWidget(CardService.newTextButton()
        .setText('Log with Drive Attachments')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('onShowDriveUploadCard_')
          .setParameters({
            res_model: 'project.task',
            res_id: String(result.task_id),
            author_email: partnerEmail || (params.sender_email || ''),
            subject: name
          })
        )
      );
    }

    var card = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle('Task Created'))
      .addSection(successSection)
      .build();

    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().pushCard(card))
      .setNotification(CardService.newNotification().setText('Task created!'))
      .build();
  } catch (err) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Error: ' + err.message))
      .build();
  }
}

// ─── TICKET SEARCH CARD ──────────────────────────────────────────────────────

function onNavigateToTicketSearch(e) {
  var p = e.parameters || {};
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(
      buildTicketSearchCard_({
        suggested_team_id: p.suggested_team_id || '',
        sender_email: p.sender_email || '',
        sender_name: p.sender_name || '',
        search_term: '',
        offset: 0,
        compose_ctx: p.compose_ctx || ''
      })
    ))
    .build();
}

function buildTicketSearchCard_(params) {
  params = params || {};
  var offset = parseInt(params.offset) || 0;
  var searchTerm = params.search_term || '';
  var teamId = params.team_id || params.suggested_team_id || '';
  var stageId = params.stage_id || '';
  var userId = params.user_id || '';

  // Load dropdowns in parallel
  var dropdownResults = odooFetchAll_([
    { path: '/gmail_addon/team/dropdown', params: {} },
    { path: '/gmail_addon/stage/dropdown', params: { project_id: null, team_id: teamId || null, record_type: 'ticket' } },
    { path: '/gmail_addon/user/dropdown', params: {} }
  ]);
  var teams = (dropdownResults[0] && dropdownResults[0].teams) || [];
  var stages = (dropdownResults[1] && dropdownResults[1].stages) || [];
  var users = (dropdownResults[2] && dropdownResults[2].users) || [];

  var formSection = CardService.newCardSection().setHeader('Search Tickets');

  formSection.addWidget(CardService.newTextInput()
    .setFieldName('search_term')
    .setTitle('Search')
    .setHint('Ticket name or #id')
    .setValue(searchTerm)
  );

  formSection.addWidget(CardService.newTextParagraph().setText('Team'));
  var teamInput = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setFieldName('team_id')
    .setOnChangeAction(CardService.newAction()
      .setFunctionName('onTicketTeamChanged')
      .setParameters({ offset: '0', search_term: searchTerm, stage_id: '',
                       sender_email: params.sender_email || '',
                       sender_name: params.sender_name || '',
                       compose_ctx: params.compose_ctx || '' })
    )
    .addItem('All teams', '', !teamId);
  teams.forEach(function(t) {
    teamInput.addItem(t.name, String(t.id), String(t.id) === String(teamId));
  });
  formSection.addWidget(teamInput);

  formSection.addWidget(CardService.newTextParagraph().setText('Stage'));
  var stageInput = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setFieldName('stage_id')
    .addItem('All stages', '', !stageId);
  stages.forEach(function(s) {
    stageInput.addItem(s.name, String(s.id), String(s.id) === String(stageId));
  });
  formSection.addWidget(stageInput);

  // Assigned user dropdown
  formSection.addWidget(CardService.newTextParagraph().setText('Assigned to'));
  var userInput = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setFieldName('user_id')
    .addItem('Anyone', '', !userId);
  users.forEach(function(u) {
    userInput.addItem(u.name, String(u.id), String(u.id) === String(userId));
  });
  formSection.addWidget(userInput);

  formSection.addWidget(CardService.newButtonSet()
    .addButton(CardService.newTextButton()
      .setText('Search')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('onTicketSearch')
        .setParameters({ offset: '0', sender_email: params.sender_email || '',
                         sender_name: params.sender_name || '',
                         compose_ctx: params.compose_ctx || '' })
      )
    )
    .addButton(CardService.newTextButton()
      .setText('+ New Ticket')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('onNavigateToCreateTicket')
        .setParameters({
          subject: '',
          plain_body: '',
          cc: '',
          sender_email: params.sender_email || '',
          sender_name: params.sender_name || '',
          suggested_team_id: teamId,
          suggested_team_name: ''
        })
      )
    )
  );

  var cardBuilder = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Search Tickets'))
    .addSection(formSection);

  if (params.error) {
    cardBuilder.addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph().setText(params.error))
    );
  } else if (params.results) {
    var results = params.results;
    var total = params.total || 0;
    var resultsSection = CardService.newCardSection()
      .setHeader('Results (' + total + ' found)');

    if (results.length === 0) {
      resultsSection.addWidget(CardService.newTextParagraph().setText('No tickets found.'));
    } else {
      results.forEach(function(ticket) {
        var ticketRef = ticket.ticket_ref || ('#' + ticket.id);
        var ticketAssignee = ticket.user_name ? ' · ' + ticket.user_name : '';
        var ticketRefStr = ticketRef + ' · ' + ticket.name + ticketAssignee;
        var ticketWidget = CardService.newDecoratedText()
          .setTopLabel('[Ticket] ' + ticketRef + ' · ' + (ticket.team_name || ''))
          .setText(ticket.name)
          .setBottomLabel(ticket.stage_name || '')
          .setOpenLink(CardService.newOpenLink().setUrl(ticket.url));
        var ticketButtonSet = CardService.newButtonSet();
        if (!params.compose_ctx) {
          var ticketComposeAction = CardService.newAction()
            .setFunctionName('onInsertReferenceInEmail')
            .setParameters({ reference: ticketRefStr, url: ticket.url });
          ticketWidget.setButton(CardService.newImageButton()
            .setIconUrl('https://www.gstatic.com/images/icons/material/system/1x/comment_grey600_18dp.png')
            .setAltText('Log email')
            .setOnClickAction(CardService.newAction()
              .setFunctionName('onNavigateToLogEmail')
              .setParameters({
                res_model: 'helpdesk.ticket',
                res_id: String(ticket.id),
                record_name: ticket.name,
                sender_email: params.sender_email || '',
                subject: params.subject || ''
              })
            )
          );
          ticketButtonSet.addButton(CardService.newImageButton()
            .setIconUrl('https://www.gstatic.com/images/icons/material/system/1x/reply_grey600_18dp.png')
            .setAltText('Reply with reference')
            .setComposeAction(ticketComposeAction, CardService.ComposedEmailType.REPLY_AS_NEW_EMAIL)
          );
        } else {
          ticketButtonSet.addButton(
            CardService.newImageButton()
              .setIconUrl('https://www.gstatic.com/images/icons/material/system/1x/insert_link_black_18dp.png')
              .setAltText('Insert reference at cursor')
              .setOnClickAction(CardService.newAction()
                .setFunctionName('onInsertAtCursor')
                .setParameters({ reference: ticketRefStr, url: ticket.url })
              )
          );
        }
        resultsSection.addWidget(ticketWidget);
        resultsSection.addWidget(ticketButtonSet);
        resultsSection.addWidget(CardService.newDivider());
      });

      var paginationSet = CardService.newButtonSet();
      if (offset > 0) {
        paginationSet.addButton(CardService.newTextButton()
          .setText('◄ Prev')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('onTicketSearch')
            .setParameters({
              offset: String(Math.max(0, offset - 10)),
              search_term: searchTerm, team_id: teamId, stage_id: stageId, user_id: userId,
              sender_email: params.sender_email || '',
              sender_name: params.sender_name || '',
              compose_ctx: params.compose_ctx || ''
            })
          )
        );
      }
      if (offset + results.length < total) {
        paginationSet.addButton(CardService.newTextButton()
          .setText('Next ►')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('onTicketSearch')
            .setParameters({
              offset: String(offset + 10),
              search_term: searchTerm, team_id: teamId, stage_id: stageId, user_id: userId,
              sender_email: params.sender_email || '',
              sender_name: params.sender_name || '',
              compose_ctx: params.compose_ctx || ''
            })
          )
        );
      }
      resultsSection.addWidget(paginationSet);
    }
    cardBuilder.addSection(resultsSection);
  }

  return cardBuilder.build();
}

function onTicketSearch(e) {
  var form = e.formInput || {};
  var params = e.parameters || {};
  var searchTerm = form.search_term || params.search_term || '';
  var teamId = form.team_id || params.team_id || '';
  var stageId = form.stage_id || params.stage_id || '';
  var userId = form.user_id || params.user_id || '';
  var offset = parseInt(params.offset) || 0;

  var results = [];
  var total = 0;
  var errorMsg = '';
  try {
    var data = apiTicketSearch_({
      search_term: searchTerm,
      team_id: teamId ? parseInt(teamId) : null,
      stage_id: stageId ? parseInt(stageId) : null,
      user_id: userId ? parseInt(userId) : null,
      limit: 10,
      offset: offset
    });
    if (data.error) {
      errorMsg = data.error;
    } else {
      results = data.tickets || [];
      total = data.total || 0;
    }
  } catch (err) {
    errorMsg = err.message;
  }

  var cardParams = {
    search_term: searchTerm,
    team_id: teamId,
    stage_id: stageId,
    user_id: userId,
    offset: offset,
    results: results,
    total: total,
    error: errorMsg || null,
    sender_email: params.sender_email || '',
    sender_name: params.sender_name || '',
    subject: params.subject || '',
    compose_ctx: params.compose_ctx || ''
  };

  var card = buildTicketSearchCard_(cardParams);
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(card))
    .build();
}

function onTicketTeamChanged(e) {
  var form = e.formInput || {};
  var params = e.parameters || {};
  var teamId = form.team_id || '';
  return onTicketSearch(Object.assign({}, e, {
    formInput: form,
    parameters: Object.assign({}, params, { offset: '0', team_id: teamId, stage_id: '' })
  }));
}

// ─── CREATE TICKET CARD ──────────────────────────────────────────────────────

function onNavigateToCreateTicket(e) {
  var p = e.parameters || {};
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(buildTicketCreateCard_(p)))
    .build();
}

function buildTicketCreateCard_(params) {
  params = params || {};
  var teams = [];
  var schema = { extra_fields: [] };
  try { teams = (apiTeamDropdown_().teams) || []; } catch (e) {}
  try { schema = apiFormSchema_('ticket') || schema; } catch (e) {}

  var section = CardService.newCardSection().setHeader('New Ticket');

  section.addWidget(CardService.newTextParagraph().setText('Helpdesk Team *'));
  var teamInput = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setFieldName('team_id')
    .addItem('Select a team...', '', false);
  teams.forEach(function(t) {
    var selected = String(t.id) === String(params.suggested_team_id || '');
    teamInput.addItem(t.name, String(t.id), selected);
  });
  section.addWidget(teamInput);

  section.addWidget(CardService.newTextInput()
    .setFieldName('ticket_name')
    .setTitle('Ticket Name *')
    .setValue(params.subject || '')
  );

  section.addWidget(CardService.newTextInput()
    .setFieldName('partner_email')
    .setTitle('Customer email')
    .setHint('email of the customer')
    .setValue(params.sender_email || '')
  );

  section.addWidget(CardService.newTextParagraph().setText('Priority'));
  var priorityInput = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setFieldName('priority')
    .addItem('Low', '0', false)
    .addItem('Normal', '1', true)
    .addItem('High', '2', false)
    .addItem('Urgent', '3', false);
  section.addWidget(priorityInput);

  section.addWidget(CardService.newTextInput()
    .setFieldName('description')
    .setTitle('Description')
    .setMultiline(true)
    .setValue(params.plain_body || '')
  );

  section.addWidget(CardService.newTextInput()
    .setFieldName('cc_addresses')
    .setTitle('CC / Followers')
    .setHint('Comma-separated emails')
    .setValue(params.cc || '')
  );

  addDynamicFieldWidgets_(section, schema);

  section.addWidget(CardService.newTextButton()
    .setText('Create Ticket')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('onCreateTicket')
      .setParameters({
        sender_email: params.sender_email || '',
        sender_name: params.sender_name || ''
      })
    )
  );

  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('New Ticket'))
    .addSection(section)
    .build();
}

function onCreateTicket(e) {
  var form = e.formInput || {};
  var params = e.parameters || {};
  var teamId = form.team_id;
  var name = (form.ticket_name || '').trim();
  var ccAddresses = extractEmailsFromAddressList_(form.cc_addresses || '');

  if (!teamId || !name) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Team and Ticket Name are required.'))
      .build();
  }

  var schema = { extra_fields: [] };
  try { schema = apiFormSchema_('ticket') || schema; } catch (err) {}

  var partnerId = null;
  var partnerEmail = (form.partner_email || '').trim();
  if (partnerEmail) {
    try {
      var matches = apiPartnerAutocomplete_(partnerEmail);
      if (matches.partners && matches.partners.length > 0) {
        partnerId = matches.partners[0].id;
      }
    } catch (err) {}
  }

  // Fetch email body and RFC Message-ID from the current email
  var emailBody = '';
  var rfcMessageId = '';
  var gmailMessageId = (e.gmail && e.gmail.messageId) || '';
  var gmailThreadId = (e.gmail && e.gmail.threadId) || '';
  if (gmailMessageId) {
    try {
      var msg = GmailApp.getMessageById(gmailMessageId);
      emailBody = cleanEmailHtml_(msg.getBody());
      try { rfcMessageId = msg.getHeader('Message-ID') || ''; } catch (_) {}
    } catch (err) {}
  }

  try {
    var result = apiCreateTicket_({
      team_id: parseInt(teamId),
      name: name,
      partner_id: partnerId,
      priority: form.priority || '1',
      description: form.description || '',
      cc_addresses: ccAddresses,
      extra_values: extractDynamicFieldValues_(form, schema),
      email_body: emailBody,
      email_subject: name,
      author_email: partnerEmail || (params.sender_email || ''),
      rfc_message_id: rfcMessageId,
      gmail_message_id: gmailMessageId,
      gmail_thread_id: gmailThreadId
    });

    if (result.error) {
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText('Error: ' + result.error))
        .build();
    }

    var successSection = CardService.newCardSection()
      .addWidget(CardService.newTextParagraph().setText('Ticket created successfully!'))
      .addWidget(CardService.newTextButton()
        .setText('Open Ticket in Odoo')
        .setOpenLink(CardService.newOpenLink().setUrl(result.ticket_url))
      );

    // Offer Drive attachment upload as a separate action when a Drive folder is configured
    if (gmailMessageId && getDriveFolderId_()) {
      successSection.addWidget(CardService.newTextButton()
        .setText('Log with Drive Attachments')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('onShowDriveUploadCard_')
          .setParameters({
            res_model: 'helpdesk.ticket',
            res_id: String(result.ticket_id),
            author_email: partnerEmail || (params.sender_email || ''),
            subject: name
          })
        )
      );
    }

    var card = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle('Ticket Created'))
      .addSection(successSection)
      .build();

    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().pushCard(card))
      .setNotification(CardService.newNotification().setText('Ticket created!'))
      .build();
  } catch (err) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Error: ' + err.message))
      .build();
  }
}

// ─── LEAD SEARCH CARD ───────────────────────────────────────────────────────

function onNavigateToLeadSearch(e) {
  var p = e.parameters || {};
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(
      buildLeadSearchCard_({
        suggested_team_id: p.suggested_team_id || '',
        sender_email: p.sender_email || '',
        sender_name: p.sender_name || '',
        search_term: '',
        offset: 0,
        compose_ctx: p.compose_ctx || ''
      })
    ))
    .build();
}

function buildLeadSearchCard_(params) {
  params = params || {};
  var offset = parseInt(params.offset) || 0;
  var searchTerm = params.search_term || '';
  var teamId = params.team_id || params.suggested_team_id || '';
  var stageId = params.stage_id || '';
  var userId = params.user_id || '';
  var leadType = params.lead_type || 'all';

  var dropdownResults = odooFetchAll_([
    { path: '/gmail_addon/crm/team/dropdown', params: {} },
    { path: '/gmail_addon/crm/stage/dropdown', params: { team_id: teamId || null } },
    { path: '/gmail_addon/user/dropdown', params: {} }
  ]);
  var teams = (dropdownResults[0] && dropdownResults[0].teams) || [];
  var stages = (dropdownResults[1] && dropdownResults[1].stages) || [];
  var users = (dropdownResults[2] && dropdownResults[2].users) || [];

  var formSection = CardService.newCardSection().setHeader('Search Leads');
  formSection.addWidget(CardService.newTextInput()
    .setFieldName('search_term')
    .setTitle('Search')
    .setHint('Lead, opportunity, company, email, or reference')
    .setValue(searchTerm)
  );

  formSection.addWidget(CardService.newTextParagraph().setText('Type'));
  formSection.addWidget(CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setFieldName('lead_type')
    .addItem('Both', 'all', leadType === 'all')
    .addItem('Leads', 'lead', leadType === 'lead')
    .addItem('Opportunities', 'opportunity', leadType === 'opportunity')
  );

  formSection.addWidget(CardService.newTextParagraph().setText('Sales Team'));
  var teamInput = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setFieldName('team_id')
    .setOnChangeAction(CardService.newAction()
      .setFunctionName('onLeadTeamChanged')
      .setParameters({ offset: '0', search_term: searchTerm, stage_id: '',
                       user_id: userId, lead_type: leadType,
                       sender_email: params.sender_email || '',
                       sender_name: params.sender_name || '',
                       compose_ctx: params.compose_ctx || '' })
    )
    .addItem('All teams', '', !teamId);
  teams.forEach(function(team) {
    teamInput.addItem(team.name, String(team.id), String(team.id) === String(teamId));
  });
  formSection.addWidget(teamInput);

  formSection.addWidget(CardService.newTextParagraph().setText('Stage'));
  var stageInput = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setFieldName('stage_id')
    .addItem('All stages', '', !stageId);
  stages.forEach(function(stage) {
    stageInput.addItem(stage.name, String(stage.id), String(stage.id) === String(stageId));
  });
  formSection.addWidget(stageInput);

  formSection.addWidget(CardService.newTextParagraph().setText('Assigned to'));
  var userInput = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setFieldName('user_id')
    .addItem('Anyone', '', !userId);
  users.forEach(function(user) {
    userInput.addItem(user.name, String(user.id), String(user.id) === String(userId));
  });
  formSection.addWidget(userInput);

  formSection.addWidget(CardService.newButtonSet()
    .addButton(CardService.newTextButton()
      .setText('Search')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('onLeadSearch')
        .setParameters({ offset: '0', sender_email: params.sender_email || '',
                         sender_name: params.sender_name || '',
                         compose_ctx: params.compose_ctx || '' })
      )
    )
    .addButton(CardService.newTextButton()
      .setText('+ New Lead')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('onNavigateToCreateLead')
        .setParameters({
          subject: '',
          plain_body: '',
          cc: '',
          sender_email: params.sender_email || '',
          sender_name: params.sender_name || '',
          suggested_team_id: teamId
        })
      )
    )
  );

  var cardBuilder = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Search Leads'))
    .addSection(formSection);

  if (params.error) {
    cardBuilder.addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph().setText(params.error))
    );
  } else if (params.results) {
    var results = params.results;
    var total = params.total || 0;
    var resultsSection = CardService.newCardSection().setHeader('Results (' + total + ' found)');

    if (results.length === 0) {
      resultsSection.addWidget(CardService.newTextParagraph().setText('No leads or opportunities found.'));
    } else {
      results.forEach(function(lead) {
        var leadRef = getRecordRef_(lead);
        var assignee = lead.user_name ? ' · ' + lead.user_name : '';
        var leadRefStr = leadRef + ' · ' + lead.name + assignee;
        var leadWidget = CardService.newDecoratedText()
          .setTopLabel('[' + (lead.type_label || 'Lead') + '] ' + leadRef + ' · ' + (lead.team_name || ''))
          .setText(lead.name)
          .setBottomLabel((lead.stage_name || '') + assignee)
          .setOpenLink(CardService.newOpenLink().setUrl(lead.url));
        var leadButtonSet = CardService.newButtonSet();
        if (!params.compose_ctx) {
          var composeAction = CardService.newAction()
            .setFunctionName('onInsertReferenceInEmail')
            .setParameters({ reference: leadRefStr, url: lead.url });
          leadWidget.setButton(CardService.newImageButton()
            .setIconUrl('https://www.gstatic.com/images/icons/material/system/1x/comment_grey600_18dp.png')
            .setAltText('Log email')
            .setOnClickAction(CardService.newAction()
              .setFunctionName('onNavigateToLogEmail')
              .setParameters({
                res_model: 'crm.lead',
                res_id: String(lead.id),
                record_name: lead.name,
                sender_email: params.sender_email || '',
                subject: params.subject || ''
              })
            )
          );
          leadButtonSet.addButton(CardService.newImageButton()
            .setIconUrl('https://www.gstatic.com/images/icons/material/system/1x/reply_grey600_18dp.png')
            .setAltText('Reply with reference')
            .setComposeAction(composeAction, CardService.ComposedEmailType.REPLY_AS_NEW_EMAIL)
          );
        } else {
          leadButtonSet.addButton(CardService.newImageButton()
            .setIconUrl('https://www.gstatic.com/images/icons/material/system/1x/insert_link_black_18dp.png')
            .setAltText('Insert reference at cursor')
            .setOnClickAction(CardService.newAction()
              .setFunctionName('onInsertAtCursor')
              .setParameters({ reference: leadRefStr, url: lead.url })
            )
          );
        }
        resultsSection.addWidget(leadWidget);
        resultsSection.addWidget(leadButtonSet);
        resultsSection.addWidget(CardService.newDivider());
      });

      var paginationSet = CardService.newButtonSet();
      if (offset > 0) {
        paginationSet.addButton(CardService.newTextButton()
          .setText('◄ Prev')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('onLeadSearch')
            .setParameters({
              offset: String(Math.max(0, offset - 10)),
              search_term: searchTerm, team_id: teamId, stage_id: stageId, user_id: userId,
              lead_type: leadType,
              sender_email: params.sender_email || '',
              sender_name: params.sender_name || '',
              compose_ctx: params.compose_ctx || ''
            })
          )
        );
      }
      if (offset + results.length < total) {
        paginationSet.addButton(CardService.newTextButton()
          .setText('Next ►')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('onLeadSearch')
            .setParameters({
              offset: String(offset + 10),
              search_term: searchTerm, team_id: teamId, stage_id: stageId, user_id: userId,
              lead_type: leadType,
              sender_email: params.sender_email || '',
              sender_name: params.sender_name || '',
              compose_ctx: params.compose_ctx || ''
            })
          )
        );
      }
      resultsSection.addWidget(paginationSet);
    }
    cardBuilder.addSection(resultsSection);
  }

  return cardBuilder.build();
}

function onLeadSearch(e) {
  var form = e.formInput || {};
  var params = e.parameters || {};
  var searchTerm = getSingleFormValue_(form, 'search_term') || params.search_term || '';
  var teamId = getSingleFormValue_(form, 'team_id') || params.team_id || '';
  var stageId = getSingleFormValue_(form, 'stage_id') || params.stage_id || '';
  var userId = getSingleFormValue_(form, 'user_id') || params.user_id || '';
  var leadType = getSingleFormValue_(form, 'lead_type') || params.lead_type || 'all';
  var offset = parseInt(params.offset) || 0;

  var results = [];
  var total = 0;
  var errorMsg = '';
  try {
    var data = apiLeadSearch_({
      search_term: searchTerm,
      lead_type: leadType,
      team_id: teamId ? parseInt(teamId, 10) : null,
      stage_id: stageId ? parseInt(stageId, 10) : null,
      user_id: userId ? parseInt(userId, 10) : null,
      limit: 10,
      offset: offset
    });
    if (data.error) {
      errorMsg = data.error;
    } else {
      results = data.leads || [];
      total = data.total || 0;
    }
  } catch (err) {
    errorMsg = err.message;
  }

  var card = buildLeadSearchCard_({
    search_term: searchTerm,
    team_id: teamId,
    stage_id: stageId,
    user_id: userId,
    lead_type: leadType,
    offset: offset,
    results: results,
    total: total,
    error: errorMsg || null,
    sender_email: params.sender_email || '',
    sender_name: params.sender_name || '',
    subject: params.subject || '',
    compose_ctx: params.compose_ctx || ''
  });

  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(card))
    .build();
}

function onLeadTeamChanged(e) {
  var form = e.formInput || {};
  var params = e.parameters || {};
  var teamId = getSingleFormValue_(form, 'team_id') || '';
  return onLeadSearch(Object.assign({}, e, {
    formInput: form,
    parameters: Object.assign({}, params, { offset: '0', team_id: teamId, stage_id: '' })
  }));
}

// ─── CREATE LEAD CARD ───────────────────────────────────────────────────────

function onNavigateToCreateLead(e) {
  var p = e.parameters || {};
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(buildLeadCreateCard_(p)))
    .build();
}

function buildLeadCreateCard_(params) {
  params = params || {};
  var teams = [];
  var schema = { extra_fields: [] };
  try { teams = (apiCrmTeamDropdown_().teams) || []; } catch (e) {}
  try { schema = apiFormSchema_('lead') || schema; } catch (e) {}

  var section = CardService.newCardSection().setHeader('New Lead');

  section.addWidget(CardService.newTextParagraph().setText('Type *'));
  section.addWidget(CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setFieldName('lead_type')
    .addItem('Lead', 'lead', true)
    .addItem('Opportunity', 'opportunity', false)
  );

  section.addWidget(CardService.newTextParagraph().setText('Sales Team'));
  var teamInput = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setFieldName('team_id')
    .addItem('Use Odoo default', '', !(params.suggested_team_id || ''));
  teams.forEach(function(team) {
    teamInput.addItem(team.name, String(team.id), String(team.id) === String(params.suggested_team_id || ''));
  });
  section.addWidget(teamInput);

  section.addWidget(CardService.newTextInput()
    .setFieldName('lead_name')
    .setTitle('Lead Name *')
    .setValue(params.subject || '')
  );

  section.addWidget(CardService.newTextInput()
    .setFieldName('partner_email')
    .setTitle('Customer email')
    .setHint('email of the customer')
    .setValue(params.sender_email || '')
  );

  section.addWidget(CardService.newTextInput()
    .setFieldName('contact_name')
    .setTitle('Contact name')
    .setValue(params.sender_name || '')
  );

  section.addWidget(CardService.newTextInput()
    .setFieldName('company_name')
    .setTitle('Company')
  );

  section.addWidget(CardService.newTextInput()
    .setFieldName('description')
    .setTitle('Description')
    .setMultiline(true)
    .setValue(params.plain_body || '')
  );

  section.addWidget(CardService.newTextInput()
    .setFieldName('cc_addresses')
    .setTitle('CC / Followers')
    .setHint('Comma-separated emails')
    .setValue(params.cc || '')
  );

  addDynamicFieldWidgets_(section, schema);

  section.addWidget(CardService.newTextButton()
    .setText('Create Lead')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('onCreateLead')
      .setParameters({
        sender_email: params.sender_email || '',
        sender_name: params.sender_name || ''
      })
    )
  );

  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('New Lead'))
    .addSection(section)
    .build();
}

function onCreateLead(e) {
  var form = e.formInput || {};
  var params = e.parameters || {};
  var name = (getSingleFormValue_(form, 'lead_name') || '').trim();
  var leadType = getSingleFormValue_(form, 'lead_type') || 'lead';
  var teamId = getSingleFormValue_(form, 'team_id') || '';
  var partnerEmail = (getSingleFormValue_(form, 'partner_email') || '').trim();
  var contactName = (getSingleFormValue_(form, 'contact_name') || params.sender_name || '').trim();
  var companyName = (getSingleFormValue_(form, 'company_name') || '').trim();
  var ccAddresses = extractEmailsFromAddressList_(getSingleFormValue_(form, 'cc_addresses') || '');

  if (!name) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Lead Name is required.'))
      .build();
  }

  var schema = { extra_fields: [] };
  try { schema = apiFormSchema_('lead') || schema; } catch (err) {}

  var partnerId = null;
  if (partnerEmail) {
    try {
      var matches = apiPartnerAutocomplete_(partnerEmail);
      if (matches.partners && matches.partners.length > 0) {
        partnerId = matches.partners[0].id;
      }
    } catch (err) {}
  }

  var emailBody = '';
  var rfcMessageId = '';
  var gmailMessageId = (e.gmail && e.gmail.messageId) || '';
  var gmailThreadId = (e.gmail && e.gmail.threadId) || '';
  if (gmailMessageId) {
    try {
      var msg = GmailApp.getMessageById(gmailMessageId);
      emailBody = cleanEmailHtml_(msg.getBody());
      try { rfcMessageId = msg.getHeader('Message-ID') || ''; } catch (_) {}
    } catch (err) {}
  }

  try {
    var result = apiCreateLead_({
      name: name,
      lead_type: leadType,
      team_id: teamId ? parseInt(teamId, 10) : null,
      partner_id: partnerId,
      contact_name: contactName,
      partner_name: companyName,
      email_from: partnerEmail,
      description: getSingleFormValue_(form, 'description') || '',
      cc_addresses: ccAddresses,
      extra_values: extractDynamicFieldValues_(form, schema),
      email_body: emailBody,
      email_subject: name,
      author_email: partnerEmail || (params.sender_email || ''),
      rfc_message_id: rfcMessageId,
      gmail_message_id: gmailMessageId,
      gmail_thread_id: gmailThreadId
    });

    if (result.error) {
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText('Error: ' + result.error))
        .build();
    }

    var recordLabel = leadType === 'opportunity' ? 'Opportunity' : 'Lead';
    var successSection = CardService.newCardSection()
      .addWidget(CardService.newTextParagraph().setText(recordLabel + ' created successfully!'))
      .addWidget(CardService.newTextButton()
        .setText('Open ' + recordLabel + ' in Odoo')
        .setOpenLink(CardService.newOpenLink().setUrl(result.lead_url))
      );

    if (gmailMessageId && getDriveFolderId_()) {
      successSection.addWidget(CardService.newTextButton()
        .setText('Log with Drive Attachments')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('onShowDriveUploadCard_')
          .setParameters({
            res_model: 'crm.lead',
            res_id: String(result.lead_id),
            author_email: partnerEmail || (params.sender_email || ''),
            subject: name
          })
        )
      );
    }

    var card = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle(recordLabel + ' Created'))
      .addSection(successSection)
      .build();

    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().pushCard(card))
      .setNotification(CardService.newNotification().setText(recordLabel + ' created!'))
      .build();
  } catch (err) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Error: ' + err.message))
      .build();
  }
}

// ─── LOG EMAIL CARD ──────────────────────────────────────────────────────────

function onNavigateToLogEmail(e) {
  var p = e.parameters || {};
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(buildLogEmailCard_(p)))
    .build();
}

function buildLogEmailCard_(params) {
  params = params || {};
  var section = CardService.newCardSection()
    .setHeader('Log Email')
    .addWidget(CardService.newTextParagraph()
      .setText('Log this email to:')
    )
    .addWidget(CardService.newDecoratedText()
      .setText(params.record_name || 'this record')
      .setBottomLabel((params.res_model || '') + ' #' + (params.res_id || ''))
    )
    .addWidget(CardService.newDivider())
    .addWidget(CardService.newTextParagraph()
      .setText('Subject: ' + (params.subject || '(no subject)'))
    )
    .addWidget(CardService.newTextParagraph()
      .setText('From: ' + (params.sender_email || ''))
    )
    .addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
        .setText('Confirm & Log')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('onConfirmLogEmail')
          .setParameters({
            res_model: params.res_model || '',
            res_id: params.res_id || '',
            sender_email: params.sender_email || '',
            subject: params.subject || ''
          })
        )
      )
      .addButton(CardService.newTextButton()
        .setText('Cancel')
        .setOnClickAction(CardService.newAction().setFunctionName('onCancelLogEmail'))
      )
    );

  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Log Email to Odoo'))
    .addSection(section)
    .build();
}

function onConfirmLogEmail(e) {
  var params = e.parameters || {};
  var resModel = params.res_model;
  var resId = params.res_id;
  var senderEmail = params.sender_email || '';
  var subject = params.subject || '';

  if (!resModel || !resId) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Missing record information.'))
      .build();
  }

  // Get HTML body and RFC Message-ID from current email
  var emailBody = '';
  var rfcMessageId = '';
  var gmailMessageId = (e.gmail && e.gmail.messageId) || '';
  var gmailThreadId = (e.gmail && e.gmail.threadId) || '';
  if (gmailMessageId) {
    try {
      var msg = GmailApp.getMessageById(gmailMessageId);
      emailBody = cleanEmailHtml_(msg.getBody());
      try { rfcMessageId = msg.getHeader('Message-ID') || ''; } catch (_) {}
    } catch (err) {
      emailBody = '<p>(email body unavailable)</p>';
    }
  }

  try {
    var result = apiLogEmail_({
      res_model: resModel,
      res_id: parseInt(resId),
      email_body: emailBody,
      email_subject: subject,
      author_email: senderEmail,
      rfc_message_id: rfcMessageId,
      gmail_message_id: gmailMessageId,
      gmail_thread_id: gmailThreadId
    });

    if (result.error) {
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText('Error: ' + result.error))
        .build();
    }

    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().popCard())
      .setNotification(CardService.newNotification().setText('Email logged to Odoo!'))
      .build();
  } catch (err) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Error: ' + err.message))
      .build();
  }
}

function onCancelLogEmail(e) {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().popCard())
    .build();
}

/**
 * Card action triggered from the task/ticket success card.
 * Runs Drive uploads via processEmail, then posts the enriched email body as a chatter note.
 * Falls back to plain HTML if Drive uploads fail.
 */
function onLogEmailWithDrive_(e) {
  var params = e.parameters || {};
  var resModel = params.res_model;
  var resId = params.res_id;
  var authorEmail = params.author_email || '';
  var subject = params.subject || '';

  var gmailMessageId = (e.gmail && e.gmail.messageId) || '';
  var gmailThreadId = (e.gmail && e.gmail.threadId) || '';

  if (!resModel || !resId || !gmailMessageId) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Missing required information.'))
      .build();
  }

  var msg = null;
  var rfcMessageId = '';
  try {
    msg = GmailApp.getMessageById(gmailMessageId);
    try { rfcMessageId = msg.getHeader('Message-ID') || ''; } catch (_) {}
  } catch (err) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Could not access email.'))
      .build();
  }

  // Read selection from form inputs (set by onShowDriveUploadCard_)
  var formInputs = (e.commonEventObject && e.commonEventObject.formInputs) || {};
  var hasFormData = Object.keys(formInputs).length > 0;
  // When no form data (fallback/direct call): include everything
  var includeImages = !hasFormData || !!(
    formInputs.include_images &&
    formInputs.include_images.stringInputs &&
    formInputs.include_images.stringInputs.value &&
    formInputs.include_images.stringInputs.value.indexOf('true') >= 0
  );
  // When form data present: upload selected attachment indices; empty selection → upload none
  var attachmentIndices = null;
  if (hasFormData) {
    var attValues = (
      formInputs.attachment_indices &&
      formInputs.attachment_indices.stringInputs &&
      formInputs.attachment_indices.stringInputs.value
    ) || [];
    attachmentIndices = attValues.map(function(v) { return parseInt(v, 10); });
  }

  // Try Drive uploads with user's selection; fall back to plain HTML on any error
  var noteBody = '';
  var fileIds = [], originalNames = [];
  try {
    var processed = processEmail(msg, { includeImages: includeImages, attachmentIndices: attachmentIndices });
    noteBody = processed.emailBody;
    fileIds = processed.fileIds;
    originalNames = processed.originalNames;
  } catch (driveErr) {
    console.warn('onLogEmailWithDrive_: Drive failed, using plain HTML', driveErr);
    noteBody = cleanEmailHtml_(msg.getBody());
  }

  if (!noteBody) noteBody = '<p>(email body unavailable)</p>';

  try {
    var result = apiLogEmail_({
      res_model: resModel,
      res_id: parseInt(resId),
      email_body: noteBody,
      email_subject: subject,
      author_email: authorEmail,
      rfc_message_id: rfcMessageId,
      gmail_message_id: gmailMessageId,
      gmail_thread_id: gmailThreadId
    });

    if (result && result.error) {
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText('Error: ' + result.error))
        .build();
    }

    if (fileIds.length) {
      try {
        var now = new Date();
        renameAndMoveFiles_(fileIds, originalNames, {
          year: String(now.getFullYear()),
          month: String(now.getMonth() + 1).padStart(2, '0'),
          recordId: String(parseInt(resId)),
          recordType: resModel === 'project.task' ? 'task' : (resModel === 'helpdesk.ticket' ? 'ticket' : 'lead')
        });
      } catch (_) {}
    }

    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Email logged with Drive attachments!'))
      .setNavigation(CardService.newNavigation().popCard())
      .build();
  } catch (err) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Error: ' + err.message))
      .build();
  }
}

/**
 * Shows a selection card listing available inline images and attachments.
 * The user can choose which items to upload to Drive before logging the email.
 * Called when the user clicks "Log with Drive Attachments" on the task/ticket success card.
 */
function onShowDriveUploadCard_(e) {
  var params = e.parameters || {};
  var gmailMessageId = (e.gmail && e.gmail.messageId) || '';

  if (!gmailMessageId) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('No email context.'))
      .build();
  }

  var msg;
  try {
    msg = GmailApp.getMessageById(gmailMessageId);
  } catch (err) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Could not access email.'))
      .build();
  }

  var html = msg.getBody() || '';
  var hasInlineImages = /src="cid:/i.test(html) || /src="https:\/\/mail\.google\.com\//i.test(html);

  var attachments = [];
  try { attachments = msg.getAttachments(); } catch (_) {}

  if (!hasInlineImages && attachments.length === 0) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('No images or attachments found.'))
      .build();
  }

  var section = CardService.newCardSection();

  if (hasInlineImages) {
    section.addWidget(
      CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.CHECK_BOX)
        .setFieldName('include_images')
        .addItem('Inline images (embedded in note)', 'true', true)
    );
  }

  if (attachments.length > 0) {
    var attInput = CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.CHECK_BOX)
      .setFieldName('attachment_indices')
      .setTitle('Attachments');
    attachments.forEach(function(att, idx) {
      attInput.addItem(att.getName() || ('Attachment ' + (idx + 1)), String(idx), true);
    });
    section.addWidget(attInput);
  }

  section.addWidget(
    CardService.newTextButton()
      .setText('Upload & Log Email')
      .setOnClickAction(
        CardService.newAction()
          .setFunctionName('onLogEmailWithDrive_')
          .setParameters(params)
      )
  );

  var card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Select items to upload'))
    .addSection(section)
    .build();

  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(card))
    .build();
}
