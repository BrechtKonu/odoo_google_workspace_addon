// ============================================================
// Odoo Tasks & Tickets — Docs & Sheets Add-on
// Google Apps Script (V8)
// ============================================================

// ─── ENTRY POINTS ─────────────────────────────────────────────────────────────

function onDocsOpen(e) {
  if (!isConfigured_()) return buildLoginCard_();
  return buildDocsSheetsHomeCard_(e, 'docs');
}

function onSheetsOpen(e) {
  if (!isConfigured_()) return buildLoginCard_();
  return buildDocsSheetsHomeCard_(e, 'sheets');
}

// ─── EDITOR CONTEXT & MEMORY ──────────────────────────────────────────────────

function getHostAppFromEvent_(e, fallbackHost) {
  var host = (fallbackHost || '').toLowerCase();
  var evHost = (e && e.commonEventObject && e.commonEventObject.hostApp) || '';
  if (evHost) host = String(evHost).toLowerCase();
  if (host === 'google_docs') host = 'docs';
  if (host === 'google_sheets') host = 'sheets';
  return host;
}

function normalizeSelectedText_(text, maxLen) {
  var s = String(text || '').trim();
  if (!s) return '';
  return s.substring(0, maxLen || 2000);
}

function getDocSelectedText_() {
  try {
    var doc = DocumentApp.getActiveDocument();
    if (!doc) return '';
    var selection = doc.getSelection();
    if (!selection) return '';
    var parts = [];
    selection.getRangeElements().forEach(function(rangeElem) {
      var el = rangeElem.getElement();
      if (!el || el.getType() !== DocumentApp.ElementType.TEXT) return;
      var text = el.asText().getText() || '';
      var start = rangeElem.getStartOffset();
      var end = rangeElem.getEndOffsetInclusive();
      if (start >= 0 && end >= start) {
        parts.push(text.substring(start, end + 1));
      } else {
        parts.push(text);
      }
    });
    return normalizeSelectedText_(parts.join('\n'), 2000);
  } catch (_) {
    return '';
  }
}

function getSheetSelectedText_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return '';
    var range = ss.getActiveRange();
    if (!range) return '';
    var values = range.getDisplayValues();
    var lines = values.map(function(row) { return row.join('\t'); });
    return normalizeSelectedText_(lines.join('\n'), 2000);
  } catch (_) {
    return '';
  }
}

function getEditorContext_(e, fallbackHost) {
  var hostApp = getHostAppFromEvent_(e, fallbackHost);
  var documentId = '';
  var documentTitle = '';
  var selectedText = '';

  if (hostApp === 'docs') {
    try {
      var doc = DocumentApp.getActiveDocument();
      if (doc) {
        documentId = doc.getId();
        documentTitle = doc.getName() || '';
        selectedText = getDocSelectedText_();
      }
    } catch (_) {}
  } else if (hostApp === 'sheets') {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (ss) {
        documentId = ss.getId();
        documentTitle = ss.getName() || '';
        selectedText = getSheetSelectedText_();
      }
    } catch (_) {}
  } else {
    // Fallback detection (when event host app is not provided)
    try {
      var d = DocumentApp.getActiveDocument();
      if (d) {
        hostApp = 'docs';
        documentId = d.getId();
        documentTitle = d.getName() || '';
        selectedText = getDocSelectedText_();
      }
    } catch (_) {}
    if (!hostApp || !documentId) {
      try {
        var s = SpreadsheetApp.getActiveSpreadsheet();
        if (s) {
          hostApp = 'sheets';
          documentId = s.getId();
          documentTitle = s.getName() || '';
          selectedText = getSheetSelectedText_();
        }
      } catch (_) {}
    }
  }

  return {
    hostApp: hostApp || '',
    documentId: documentId || '',
    documentTitle: documentTitle || '',
    selectedText: selectedText || ''
  };
}

function buildEditorCtxParams_(ctx) {
  return {
    host_app: ctx.hostApp || '',
    document_id: ctx.documentId || '',
    document_title: ctx.documentTitle || '',
    selected_text: ctx.selectedText || ''
  };
}

function getEditorCtxFromParams_(params) {
  params = params || {};
  return {
    hostApp: params.host_app || '',
    documentId: params.document_id || '',
    documentTitle: params.document_title || '',
    selectedText: params.selected_text || ''
  };
}

function getDocumentMemoryKey_(ctx) {
  if (!ctx || !ctx.hostApp || !ctx.documentId) return '';
  return 'docmem_v1_' + ctx.hostApp + '_' + ctx.documentId;
}

function getDocumentMemory_(ctx) {
  var key = getDocumentMemoryKey_(ctx);
  if (!key) return {};
  var raw = getProps_()[key];
  if (!raw) return {};
  try {
    var parsed = JSON.parse(raw);
    return parsed && typeof parsed === 'object' ? parsed : {};
  } catch (_) {
    return {};
  }
}

function saveDocumentMemory_(ctx, patch) {
  var key = getDocumentMemoryKey_(ctx);
  if (!key || !patch) return;
  var current = getDocumentMemory_(ctx);
  Object.keys(patch).forEach(function(k) { current[k] = patch[k]; });
  USER_PROPS.setProperty(key, JSON.stringify(current));
  _propsCache = null;
}

function previewText_(text, maxLen) {
  var s = normalizeSelectedText_(text, maxLen || 240);
  if (!s) return '';
  return s.length > (maxLen || 240) ? s.substring(0, (maxLen || 240)) + '…' : s;
}

function linkRecordToCurrentDocument_(ctx, resModel, resId, recordName) {
  if (!ctx || !ctx.documentId || !ctx.hostApp || !resModel || !resId) return;
  try {
    apiDocumentLinkRecord_(ctx.documentId, ctx.hostApp, resModel, parseInt(resId, 10), recordName || '');
  } catch (_) {}
}

// ─── HOME CARD ────────────────────────────────────────────────────────────────

function buildDocsSheetsHomeCard_(e, hostHint) {
  var ctx = getEditorContext_(e, hostHint);
  var memory = getDocumentMemory_(ctx);
  var linkedRecords = [];

  if (ctx.documentId && ctx.hostApp) {
    try {
      linkedRecords = (apiDocumentLinkedRecords_(ctx.documentId, ctx.hostApp).records) || [];
    } catch (_) {}
  }

  var suggestedProjectId = String(memory.last_project_id || '');
  var suggestedTeamId = String(memory.last_team_id || '');
  var taskFilterProjectId = String(memory.task_filter_project_id || suggestedProjectId || '');
  var ticketFilterTeamId = String(memory.ticket_filter_team_id || suggestedTeamId || '');

  var contextSection = CardService.newCardSection().setHeader('Context');
  contextSection.addWidget(CardService.newDecoratedText()
    .setTopLabel(ctx.hostApp === 'sheets' ? 'Google Sheets' : 'Google Docs')
    .setText(ctx.documentTitle || '(Untitled)')
    .setBottomLabel(ctx.documentId || '(No document id)')
    .setStartIcon(CardService.newIconImage().setIcon(CardService.Icon.DESCRIPTION))
  );
  if (ctx.selectedText) {
    contextSection.addWidget(CardService.newTextParagraph()
      .setText('<b>Selected text</b><br>' + escapeHtml_(previewText_(ctx.selectedText, 400)))
    );
  } else {
    contextSection.addWidget(CardService.newTextParagraph()
      .setText('No selected text detected. You can still search/create records.')
    );
  }

  var linkedSection = CardService.newCardSection().setHeader('Recent linked to this document');
  if (linkedRecords.length === 0) {
    linkedSection.addWidget(CardService.newTextParagraph().setText('No linked tasks or tickets yet.'));
  } else {
    linkedRecords.forEach(function(rec) {
      var refId = rec.type === 'task'
        ? (rec.task_number || ('#' + rec.id))
        : (rec.ticket_ref || ('#' + rec.id));
      var assignee = rec.user_name ? ' · ' + rec.user_name : '';
      var refStr = refId + ' · ' + rec.name + assignee;
      var icon = rec.type === 'task'
        ? CardService.Icon.DESCRIPTION
        : CardService.Icon.CONFIRMATION_NUMBER_ICON;

      linkedSection.addWidget(CardService.newDecoratedText()
        .setTopLabel((rec.type === 'task' ? 'Task' : 'Ticket') + ' ' + refId)
        .setText(rec.name)
        .setBottomLabel(rec.stage || '')
        .setStartIcon(CardService.newIconImage().setIcon(icon))
        .setOpenLink(CardService.newOpenLink().setUrl(rec.url))
        .setButton(CardService.newImageButton()
          .setIconUrl('https://www.gstatic.com/images/icons/material/system/1x/insert_link_black_18dp.png')
          .setAltText('Insert link')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('onInsertReferenceInEditor')
            .setParameters(Object.assign({
              reference: refStr,
              url: rec.url,
              res_model: rec.type === 'task' ? 'project.task' : 'helpdesk.ticket',
              res_id: String(rec.id),
              record_name: rec.name
            }, buildEditorCtxParams_(ctx)))
          )
        )
      );
      linkedSection.addWidget(CardService.newDivider());
    });
  }

  var actionsSection = CardService.newCardSection().setHeader('Actions');
  actionsSection.addWidget(CardService.newButtonSet()
    .addButton(CardService.newTextButton()
      .setText('Search Tasks')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('onNavigateToDocTaskSearch')
        .setParameters(Object.assign({
          project_id: taskFilterProjectId,
          search_term: '',
          offset: '0'
        }, buildEditorCtxParams_(ctx)))
      )
    )
    .addButton(CardService.newTextButton()
      .setText('Search Tickets')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('onNavigateToDocTicketSearch')
        .setParameters(Object.assign({
          team_id: ticketFilterTeamId,
          search_term: '',
          offset: '0'
        }, buildEditorCtxParams_(ctx)))
      )
    )
  );

  actionsSection.addWidget(CardService.newButtonSet()
    .addButton(CardService.newTextButton()
      .setText('New Task')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('onNavigateToDocCreateTask')
        .setParameters(Object.assign({
          suggested_project_id: suggestedProjectId
        }, buildEditorCtxParams_(ctx)))
      )
    )
    .addButton(CardService.newTextButton()
      .setText('New Ticket')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('onNavigateToDocCreateTicket')
        .setParameters(Object.assign({
          suggested_team_id: suggestedTeamId
        }, buildEditorCtxParams_(ctx)))
      )
    )
  );

  actionsSection.addWidget(CardService.newTextButton()
    .setText('Settings')
    .setOnClickAction(CardService.newAction().setFunctionName('onNavigateToSettings'))
  );

  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Odoo Tasks & Tickets'))
    .addSection(contextSection)
    .addSection(linkedSection)
    .addSection(actionsSection)
    .build();
}

// ─── INSERT LINK INTO DOC/SHEET ───────────────────────────────────────────────

function insertReferenceIntoDoc_(reference, url) {
  var doc = DocumentApp.getActiveDocument();
  if (!doc) throw new Error('No active document.');
  var cursor = doc.getCursor();
  if (cursor) {
    var inserted = cursor.insertText(reference);
    if (inserted) inserted.setLinkUrl(url);
    cursor.insertText('\n');
    return;
  }
  var body = doc.getBody();
  var p = body.appendParagraph(reference);
  var t = p.editAsText();
  t.setLinkUrl(0, reference.length - 1, url);
}

function insertReferenceIntoSheet_(reference, url) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('No active spreadsheet.');
  var range = ss.getActiveRange();
  if (!range) throw new Error('Select a target cell first.');
  var rich = SpreadsheetApp.newRichTextValue()
    .setText(reference)
    .setLinkUrl(url)
    .build();
  range.offset(0, 0, 1, 1).setRichTextValue(rich);
}

function onInsertReferenceInEditor(e) {
  var params = e.parameters || {};
  var reference = params.reference || '';
  var url = params.url || '';
  var ctx = getEditorCtxFromParams_(params);
  var host = (params.host_app || ctx.hostApp || '').toLowerCase();

  if (!reference || !url) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Missing reference URL or text.'))
      .build();
  }

  try {
    if (host === 'docs') {
      insertReferenceIntoDoc_(reference, url);
    } else if (host === 'sheets') {
      insertReferenceIntoSheet_(reference, url);
    } else {
      throw new Error('Unsupported host app context.');
    }
    if (ctx.documentId && host && params.res_model && params.res_id) {
      linkRecordToCurrentDocument_(ctx, params.res_model, params.res_id, params.record_name || reference);
    }
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Inserted link into ' + host + '.'))
      .build();
  } catch (err) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Insert failed: ' + err.message))
      .build();
  }
}

// ─── TASK SEARCH (DOCS/SHEETS) ────────────────────────────────────────────────

function onNavigateToDocTaskSearch(e) {
  var p = e.parameters || {};
  var params = Object.assign({
    search_term: '',
    offset: '0',
    project_id: p.project_id || p.suggested_project_id || '',
    stage_id: ''
  }, p);
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(buildDocTaskSearchCard_(params)))
    .build();
}

function buildDocTaskSearchCard_(params) {
  params = params || {};
  var offset = parseInt(params.offset, 10) || 0;
  var searchTerm = params.search_term || '';
  var projectId = params.project_id || params.suggested_project_id || '';
  var stageId = params.stage_id || '';
  var ctx = getEditorCtxFromParams_(params);

  var dropdownResults = odooFetchAll_([
    { path: '/gmail_addon/project/dropdown', params: {} },
    { path: '/gmail_addon/stage/dropdown', params: { project_id: projectId || null, team_id: null, record_type: 'task' } }
  ]);
  var projects = (dropdownResults[0] && dropdownResults[0].projects) || [];
  var stages = (dropdownResults[1] && dropdownResults[1].stages) || [];

  var formSection = CardService.newCardSection().setHeader('Search Tasks');
  formSection.addWidget(CardService.newTextInput()
    .setFieldName('search_term')
    .setTitle('Search')
    .setHint('Task name or number (e.g. PROJ-001)')
    .setValue(searchTerm)
  );

  formSection.addWidget(CardService.newTextParagraph().setText('Project'));
  var projectInput = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setFieldName('project_id')
    .setOnChangeAction(CardService.newAction()
      .setFunctionName('onDocTaskProjectChanged')
      .setParameters(Object.assign({
        search_term: searchTerm,
        stage_id: '',
        offset: '0'
      }, buildEditorCtxParams_(ctx)))
    )
    .addItem('All projects', '', !projectId);
  projects.forEach(function(p) {
    projectInput.addItem(p.name, String(p.id), String(p.id) === String(projectId));
  });
  formSection.addWidget(projectInput);

  formSection.addWidget(CardService.newTextParagraph().setText('Stage'));
  var stageInput = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setFieldName('stage_id')
    .addItem('All stages', '', !stageId);
  stages.forEach(function(s) {
    stageInput.addItem(s.name, String(s.id), String(s.id) === String(stageId));
  });
  formSection.addWidget(stageInput);

  formSection.addWidget(CardService.newButtonSet()
    .addButton(CardService.newTextButton()
      .setText('Search')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('onDocTaskSearch')
        .setParameters(Object.assign({ offset: '0' }, buildEditorCtxParams_(ctx)))
      )
    )
    .addButton(CardService.newTextButton()
      .setText('+ New Task')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('onNavigateToDocCreateTask')
        .setParameters(Object.assign({
          suggested_project_id: projectId,
          selected_text: ctx.selectedText || ''
        }, buildEditorCtxParams_(ctx)))
      )
    )
  );

  var cardBuilder = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Search Tasks'))
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
      resultsSection.addWidget(CardService.newTextParagraph().setText('No tasks found.'));
    } else {
      results.forEach(function(task) {
        resultsSection.addWidget(CardService.newDecoratedText()
          .setTopLabel('[Task] ' + (task.task_number || ('#' + task.id)) + ' · ' + (task.project_name || ''))
          .setText(task.name)
          .setBottomLabel((task.stage_name || '') + (task.user_name ? ' · ' + task.user_name : ''))
          .setOpenLink(CardService.newOpenLink().setUrl(task.url))
        );
        var taskRef = task.task_number || ('#' + task.id);
        var assignee = task.user_name ? ' · ' + task.user_name : '';
        var refStr = taskRef + ' · ' + task.name + assignee;
        resultsSection.addWidget(CardService.newButtonSet().addButton(
          CardService.newImageButton()
            .setIconUrl('https://www.gstatic.com/images/icons/material/system/1x/insert_link_black_18dp.png')
            .setAltText('Insert link')
            .setOnClickAction(CardService.newAction()
              .setFunctionName('onInsertReferenceInEditor')
              .setParameters(Object.assign({
                reference: refStr,
                url: task.url,
                res_model: 'project.task',
                res_id: String(task.id),
                record_name: task.name
              }, buildEditorCtxParams_(ctx)))
            )
        ));
        resultsSection.addWidget(CardService.newDivider());
      });

      var pageSize = 10;
      var pagination = CardService.newButtonSet();
      if (offset > 0) {
        pagination.addButton(CardService.newTextButton()
          .setText('◄ Prev')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('onDocTaskSearch')
            .setParameters(Object.assign({
              offset: String(Math.max(0, offset - pageSize)),
              search_term: searchTerm,
              project_id: projectId,
              stage_id: stageId
            }, buildEditorCtxParams_(ctx)))
          )
        );
      }
      if (offset + results.length < total) {
        pagination.addButton(CardService.newTextButton()
          .setText('Next ►')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('onDocTaskSearch')
            .setParameters(Object.assign({
              offset: String(offset + pageSize),
              search_term: searchTerm,
              project_id: projectId,
              stage_id: stageId
            }, buildEditorCtxParams_(ctx)))
          )
        );
      }
      resultsSection.addWidget(pagination);
    }
    cardBuilder.addSection(resultsSection);
  }

  return cardBuilder.build();
}

function onDocTaskSearch(e) {
  var form = e.formInput || {};
  var params = e.parameters || {};
  var ctx = getEditorCtxFromParams_(params);
  var searchTerm = form.search_term || params.search_term || '';
  var projectId = form.project_id || params.project_id || '';
  var stageId = form.stage_id || params.stage_id || '';
  var offset = parseInt(params.offset, 10) || 0;

  saveDocumentMemory_(ctx, { task_filter_project_id: projectId || '' });

  var results = [];
  var total = 0;
  var errorMsg = '';
  try {
    var data = apiTaskSearch_({
      search_term: searchTerm,
      project_id: projectId ? parseInt(projectId, 10) : null,
      stage_id: stageId ? parseInt(stageId, 10) : null,
      limit: 10,
      offset: offset
    });
    results = data.tasks || [];
    total = data.total || 0;
  } catch (err) {
    errorMsg = err.message;
  }

  var card = buildDocTaskSearchCard_(Object.assign({
    search_term: searchTerm,
    project_id: projectId,
    stage_id: stageId,
    offset: String(offset),
    results: results,
    total: total,
    error: errorMsg || null
  }, buildEditorCtxParams_(ctx)));

  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(card))
    .build();
}

function onDocTaskProjectChanged(e) {
  var form = e.formInput || {};
  var params = e.parameters || {};
  var projectId = form.project_id || '';
  return onDocTaskSearch(Object.assign({}, e, {
    formInput: form,
    parameters: Object.assign({}, params, { offset: '0', project_id: projectId, stage_id: '' })
  }));
}

// ─── TICKET SEARCH (DOCS/SHEETS) ──────────────────────────────────────────────

function onNavigateToDocTicketSearch(e) {
  var p = e.parameters || {};
  var params = Object.assign({
    search_term: '',
    offset: '0',
    team_id: p.team_id || p.suggested_team_id || '',
    stage_id: ''
  }, p);
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(buildDocTicketSearchCard_(params)))
    .build();
}

function buildDocTicketSearchCard_(params) {
  params = params || {};
  var offset = parseInt(params.offset, 10) || 0;
  var searchTerm = params.search_term || '';
  var teamId = params.team_id || params.suggested_team_id || '';
  var stageId = params.stage_id || '';
  var ctx = getEditorCtxFromParams_(params);

  var dropdownResults = odooFetchAll_([
    { path: '/gmail_addon/team/dropdown', params: {} },
    { path: '/gmail_addon/stage/dropdown', params: { project_id: null, team_id: teamId || null, record_type: 'ticket' } }
  ]);
  var teams = (dropdownResults[0] && dropdownResults[0].teams) || [];
  var stages = (dropdownResults[1] && dropdownResults[1].stages) || [];

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
      .setFunctionName('onDocTicketTeamChanged')
      .setParameters(Object.assign({
        search_term: searchTerm,
        stage_id: '',
        offset: '0'
      }, buildEditorCtxParams_(ctx)))
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

  formSection.addWidget(CardService.newButtonSet()
    .addButton(CardService.newTextButton()
      .setText('Search')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('onDocTicketSearch')
        .setParameters(Object.assign({ offset: '0' }, buildEditorCtxParams_(ctx)))
      )
    )
    .addButton(CardService.newTextButton()
      .setText('+ New Ticket')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('onNavigateToDocCreateTicket')
        .setParameters(Object.assign({
          suggested_team_id: teamId,
          selected_text: ctx.selectedText || ''
        }, buildEditorCtxParams_(ctx)))
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
    var resultsSection = CardService.newCardSection().setHeader('Results (' + total + ' found)');
    if (results.length === 0) {
      resultsSection.addWidget(CardService.newTextParagraph().setText('No tickets found.'));
    } else {
      results.forEach(function(ticket) {
        var ticketRef = ticket.ticket_ref || ('#' + ticket.id);
        resultsSection.addWidget(CardService.newDecoratedText()
          .setTopLabel('[Ticket] ' + ticketRef + ' · ' + (ticket.team_name || ''))
          .setText(ticket.name)
          .setBottomLabel((ticket.stage_name || '') + (ticket.user_name ? ' · ' + ticket.user_name : ''))
          .setOpenLink(CardService.newOpenLink().setUrl(ticket.url))
        );
        var assignee = ticket.user_name ? ' · ' + ticket.user_name : '';
        var refStr = ticketRef + ' · ' + ticket.name + assignee;
        resultsSection.addWidget(CardService.newButtonSet().addButton(
          CardService.newImageButton()
            .setIconUrl('https://www.gstatic.com/images/icons/material/system/1x/insert_link_black_18dp.png')
            .setAltText('Insert link')
            .setOnClickAction(CardService.newAction()
              .setFunctionName('onInsertReferenceInEditor')
              .setParameters(Object.assign({
                reference: refStr,
                url: ticket.url,
                res_model: 'helpdesk.ticket',
                res_id: String(ticket.id),
                record_name: ticket.name
              }, buildEditorCtxParams_(ctx)))
            )
        ));
        resultsSection.addWidget(CardService.newDivider());
      });

      var pageSize = 10;
      var pagination = CardService.newButtonSet();
      if (offset > 0) {
        pagination.addButton(CardService.newTextButton()
          .setText('◄ Prev')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('onDocTicketSearch')
            .setParameters(Object.assign({
              offset: String(Math.max(0, offset - pageSize)),
              search_term: searchTerm,
              team_id: teamId,
              stage_id: stageId
            }, buildEditorCtxParams_(ctx)))
          )
        );
      }
      if (offset + results.length < total) {
        pagination.addButton(CardService.newTextButton()
          .setText('Next ►')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('onDocTicketSearch')
            .setParameters(Object.assign({
              offset: String(offset + pageSize),
              search_term: searchTerm,
              team_id: teamId,
              stage_id: stageId
            }, buildEditorCtxParams_(ctx)))
          )
        );
      }
      resultsSection.addWidget(pagination);
    }
    cardBuilder.addSection(resultsSection);
  }

  return cardBuilder.build();
}

function onDocTicketSearch(e) {
  var form = e.formInput || {};
  var params = e.parameters || {};
  var ctx = getEditorCtxFromParams_(params);
  var searchTerm = form.search_term || params.search_term || '';
  var teamId = form.team_id || params.team_id || '';
  var stageId = form.stage_id || params.stage_id || '';
  var offset = parseInt(params.offset, 10) || 0;

  saveDocumentMemory_(ctx, { ticket_filter_team_id: teamId || '' });

  var results = [];
  var total = 0;
  var errorMsg = '';
  try {
    var data = apiTicketSearch_({
      search_term: searchTerm,
      team_id: teamId ? parseInt(teamId, 10) : null,
      stage_id: stageId ? parseInt(stageId, 10) : null,
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

  var card = buildDocTicketSearchCard_(Object.assign({
    search_term: searchTerm,
    team_id: teamId,
    stage_id: stageId,
    offset: String(offset),
    results: results,
    total: total,
    error: errorMsg || null
  }, buildEditorCtxParams_(ctx)));

  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(card))
    .build();
}

function onDocTicketTeamChanged(e) {
  var form = e.formInput || {};
  var params = e.parameters || {};
  var teamId = form.team_id || '';
  return onDocTicketSearch(Object.assign({}, e, {
    formInput: form,
    parameters: Object.assign({}, params, { offset: '0', team_id: teamId, stage_id: '' })
  }));
}

// ─── CREATE TASK (DOCS/SHEETS) ────────────────────────────────────────────────

function onNavigateToDocCreateTask(e) {
  var p = e.parameters || {};
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(buildDocTaskCreateCard_(p)))
    .build();
}

function buildDocTaskCreateCard_(params) {
  params = params || {};
  var projects = [];
  try { projects = (apiProjectDropdown_().projects) || []; } catch (_) {}

  var section = CardService.newCardSection().setHeader('New Task');
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

  section.addWidget(CardService.newTextInput()
    .setFieldName('task_name')
    .setTitle('Task Name *')
    .setValue(params.task_name || '')
  );

  section.addWidget(CardService.newTextInput()
    .setFieldName('partner_email')
    .setTitle('Customer email')
    .setHint('email of the customer')
    .setValue(params.partner_email || '')
  );

  section.addWidget(CardService.newTextInput()
    .setFieldName('description')
    .setTitle('Description')
    .setMultiline(true)
    .setValue(params.selected_text || params.description || '')
  );

  section.addWidget(CardService.newTextInput()
    .setFieldName('cc_addresses')
    .setTitle('CC / Followers')
    .setHint('Comma-separated emails')
    .setValue(params.cc_addresses || '')
  );

  section.addWidget(CardService.newTextButton()
    .setText('Create Task')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('onDocCreateTask')
      .setParameters(buildEditorCtxParams_(getEditorCtxFromParams_(params)))
    )
  );

  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('New Task'))
    .addSection(section)
    .build();
}

function onDocCreateTask(e) {
  var form = e.formInput || {};
  var params = e.parameters || {};
  var ctx = getEditorCtxFromParams_(params);
  var projectId = form.project_id;
  var name = (form.task_name || '').trim();
  var ccAddresses = extractEmailsFromAddressList_(form.cc_addresses || '');

  if (!projectId || !name) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Project and Task Name are required.'))
      .build();
  }

  var partnerId = null;
  var partnerEmail = (form.partner_email || '').trim();
  if (partnerEmail) {
    try {
      var matches = apiPartnerAutocomplete_(partnerEmail);
      if (matches.partners && matches.partners.length > 0) partnerId = matches.partners[0].id;
    } catch (_) {}
  }

  try {
    var result = apiCreateTask_({
      project_id: parseInt(projectId, 10),
      name: name,
      partner_id: partnerId,
      description: form.description || '',
      cc_addresses: ccAddresses
    });
    if (result.error) {
      throw new Error(result.error);
    }

    saveDocumentMemory_(ctx, {
      last_project_id: String(projectId),
      task_filter_project_id: String(projectId)
    });
    linkRecordToCurrentDocument_(ctx, 'project.task', result.task_id, name);

    var taskRef = '#' + result.task_id;
    var refStr = taskRef + ' · ' + name;
    var successSection = CardService.newCardSection()
      .addWidget(CardService.newTextParagraph().setText('Task created successfully!'))
      .addWidget(CardService.newTextButton()
        .setText('Open Task in Odoo')
        .setOpenLink(CardService.newOpenLink().setUrl(result.task_url))
      )
      .addWidget(CardService.newImageButton()
        .setIconUrl('https://www.gstatic.com/images/icons/material/system/1x/insert_link_black_18dp.png')
        .setAltText('Insert link')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('onInsertReferenceInEditor')
          .setParameters(Object.assign({
            reference: refStr,
            url: result.task_url,
            res_model: 'project.task',
            res_id: String(result.task_id),
            record_name: name
          }, buildEditorCtxParams_(ctx)))
        )
      );

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

// ─── CREATE TICKET (DOCS/SHEETS) ──────────────────────────────────────────────

function onNavigateToDocCreateTicket(e) {
  var p = e.parameters || {};
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(buildDocTicketCreateCard_(p)))
    .build();
}

function buildDocTicketCreateCard_(params) {
  params = params || {};
  var teams = [];
  try { teams = (apiTeamDropdown_().teams) || []; } catch (_) {}

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
    .setValue(params.ticket_name || '')
  );

  section.addWidget(CardService.newTextInput()
    .setFieldName('partner_email')
    .setTitle('Customer email')
    .setHint('email of the customer')
    .setValue(params.partner_email || '')
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
    .setValue(params.selected_text || params.description || '')
  );

  section.addWidget(CardService.newTextInput()
    .setFieldName('cc_addresses')
    .setTitle('CC / Followers')
    .setHint('Comma-separated emails')
    .setValue(params.cc_addresses || '')
  );

  section.addWidget(CardService.newTextButton()
    .setText('Create Ticket')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('onDocCreateTicket')
      .setParameters(buildEditorCtxParams_(getEditorCtxFromParams_(params)))
    )
  );

  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('New Ticket'))
    .addSection(section)
    .build();
}

function onDocCreateTicket(e) {
  var form = e.formInput || {};
  var params = e.parameters || {};
  var ctx = getEditorCtxFromParams_(params);
  var teamId = form.team_id;
  var name = (form.ticket_name || '').trim();
  var ccAddresses = extractEmailsFromAddressList_(form.cc_addresses || '');

  if (!teamId || !name) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Team and Ticket Name are required.'))
      .build();
  }

  var partnerId = null;
  var partnerEmail = (form.partner_email || '').trim();
  if (partnerEmail) {
    try {
      var matches = apiPartnerAutocomplete_(partnerEmail);
      if (matches.partners && matches.partners.length > 0) partnerId = matches.partners[0].id;
    } catch (_) {}
  }

  try {
    var result = apiCreateTicket_({
      team_id: parseInt(teamId, 10),
      name: name,
      partner_id: partnerId,
      priority: form.priority || '1',
      description: form.description || '',
      cc_addresses: ccAddresses
    });
    if (result.error) {
      throw new Error(result.error);
    }

    saveDocumentMemory_(ctx, {
      last_team_id: String(teamId),
      ticket_filter_team_id: String(teamId)
    });
    linkRecordToCurrentDocument_(ctx, 'helpdesk.ticket', result.ticket_id, name);

    var ticketRef = '#' + result.ticket_id;
    var refStr = ticketRef + ' · ' + name;
    var successSection = CardService.newCardSection()
      .addWidget(CardService.newTextParagraph().setText('Ticket created successfully!'))
      .addWidget(CardService.newTextButton()
        .setText('Open Ticket in Odoo')
        .setOpenLink(CardService.newOpenLink().setUrl(result.ticket_url))
      )
      .addWidget(CardService.newImageButton()
        .setIconUrl('https://www.gstatic.com/images/icons/material/system/1x/insert_link_black_18dp.png')
        .setAltText('Insert link')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('onInsertReferenceInEditor')
          .setParameters(Object.assign({
            reference: refStr,
            url: result.ticket_url,
            res_model: 'helpdesk.ticket',
            res_id: String(result.ticket_id),
            record_name: name
          }, buildEditorCtxParams_(ctx)))
        )
      );

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
