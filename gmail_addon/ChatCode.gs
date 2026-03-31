// ============================================================
// Odoo Tasks & Tickets — Google Chat Add-on
// Google Apps Script (V8)
//
// Shared utilities in Code.gs (same GAS project = same global scope):
// buildOdooRequestOptions_, odooPost_, odooFetchAll_, isConfigured_,
// buildLoginCard_, apiTaskSearch_, apiTicketSearch_, apiCreateTask_,
// apiCreateTicket_, extractEmailsFromAddressList_
// ============================================================

// ─── SLASH COMMAND IDs ───────────────────────────────────────────────────────
// Must match Cloud Console → Chat API → Commands.
var CHAT_CMD_TASK          = 1;
var CHAT_CMD_TICKET        = 2;
var CHAT_CMD_CONFIG        = 4;
var CHAT_CMD_TASK_CREATE   = 7;
var CHAT_CMD_TICKET_CREATE = 9;

// ─── SPACE MEMORY ────────────────────────────────────────────────────────────
// Per-space project/team config stored in ScriptProperties (shared across
// all users — space is the correct shared unit).

function getSpaceId_(e) {
  var spaceName = (e && e.chat && e.chat.space && e.chat.space.name) ||
                  (e && e.chat && e.chat.appCommandPayload && e.chat.appCommandPayload.space && e.chat.appCommandPayload.space.name) || '';
  if (spaceName.indexOf('spaces/') === 0) return spaceName.substring(7) || 'default';
  return spaceName || 'default';
}

function getSpaceMemory_(spaceId) {
  var props = PropertiesService.getScriptProperties();
  return {
    project_id:        props.getProperty('space_' + spaceId + '_project_id')        || '',
    project_name:      props.getProperty('space_' + spaceId + '_project_name')      || '',
    project_ids_csv:   props.getProperty('space_' + spaceId + '_project_ids_csv')   || '',
    team_id:           props.getProperty('space_' + spaceId + '_team_id')           || '',
    team_name:         props.getProperty('space_' + spaceId + '_team_name')         || '',
    team_ids_csv:      props.getProperty('space_' + spaceId + '_team_ids_csv')      || '',
    space_name:        props.getProperty('space_' + spaceId + '_space_name')        || ''
  };
}

function setSpaceMemory_(spaceId, data) {
  var props = PropertiesService.getScriptProperties();
  if (data.project_id        !== undefined) props.setProperty('space_' + spaceId + '_project_id',        String(data.project_id        || ''));
  if (data.project_name      !== undefined) props.setProperty('space_' + spaceId + '_project_name',      String(data.project_name      || ''));
  if (data.project_ids_csv   !== undefined) props.setProperty('space_' + spaceId + '_project_ids_csv',   String(data.project_ids_csv   || ''));
  if (data.team_id           !== undefined) props.setProperty('space_' + spaceId + '_team_id',           String(data.team_id           || ''));
  if (data.team_name         !== undefined) props.setProperty('space_' + spaceId + '_team_name',         String(data.team_name         || ''));
  if (data.team_ids_csv      !== undefined) props.setProperty('space_' + spaceId + '_team_ids_csv',      String(data.team_ids_csv      || ''));
  if (data.space_name        !== undefined) props.setProperty('space_' + spaceId + '_space_name',        String(data.space_name        || ''));
}

function parseCsvIds_(csv) {
  if (!csv) return [];
  return String(csv).split(',').map(function(s) { return s.trim(); }).filter(Boolean);
}

function toCsvIds_(ids) {
  return (ids || []).map(function(s) { return String(s || '').trim(); }).filter(Boolean).join(',');
}

// ─── ENTRY POINTS ────────────────────────────────────────────────────────────

function onChatMessageOpen(e) {
  if (!isConfigured_()) return buildLoginCard_();
  return buildChatHomeCard_(e);
}

function onAddedToSpace(e) {
  var spaceId = getSpaceId_(e);
  var spaceName = (e && e.chat && e.chat.space && e.chat.space.name) || '';
  if (spaceId !== 'default' && spaceName) setSpaceMemory_(spaceId, { space_name: spaceName });
  if (!isConfigured_()) return buildLoginCard_();
  return buildChatHomeCard_(e);
}

/**
 * Called for plain messages (non-slash-command). Saves the message text for
 * description prefill in task/ticket create dialogs.
 */
function onMessage(e) {
  if (!isConfigured_()) return buildLoginCard_();
  var spaceId = getSpaceId_(e);
  var spaceName = (e && e.chat && e.chat.space && e.chat.space.name) || '';
  if (spaceId !== 'default' && spaceName) setSpaceMemory_(spaceId, { space_name: spaceName });
  return buildChatHomeCard_(e);
}

/**
 * Handles all slash commands. All commands are configured as dialog commands
 * in Cloud Console (triggersDialog: true).
 * Uses dialogPush_() — correct for REQUEST_DIALOG (initial open).
 */
function onAppCommand(e) {
  console.log('onAppCommand event:', JSON.stringify(e));

  if (!isConfigured_()) {
    return dialogPush_({ sections: [{ widgets: [{ textParagraph: {
      text: 'Konu is not configured. Open the add-on sidebar and go to Settings.'
    }}]}]});
  }

  var spaceId = getSpaceId_(e);
  var payload = (e && e.chat && e.chat.appCommandPayload) || {};
  var spaceName = (payload.space && payload.space.name) || '';
  if (spaceId !== 'default' && spaceName) setSpaceMemory_(spaceId, { space_name: spaceName });

  var cmdId  = parseInt(payload.appCommandMetadata && payload.appCommandMetadata.appCommandId, 10);
  var memory = getSpaceMemory_(spaceId);
  var messageText  = (payload.message && payload.message.text) || '';
  var argumentText = messageText.replace(/^\/\S+\s*/, '').trim();
  console.log('onAppCommand: cmdId=' + cmdId + ' spaceId=' + spaceId + ' arg=' + argumentText);

  try {
    switch (cmdId) {
      case CHAT_CMD_CONFIG:        return dialogPush_(buildConfigCard_(spaceId, memory));
      case CHAT_CMD_TASK: {
        if (argumentText) {
          var scopedIds = parseCsvIds_(memory.project_ids_csv);
          var projectId = memory.project_id && (!scopedIds.length || scopedIds.indexOf(memory.project_id) >= 0)
                          ? memory.project_id : '';
          var taskData = { tasks: [], total: 0 };
          try {
            taskData = apiTaskSearch_({
              search_term: argumentText,
              project_id: projectId ? parseInt(projectId, 10) : null,
              limit: 10, offset: 0
            });
            rememberSpaceRecentTasks_(spaceId, taskData.tasks || []);
          } catch (_) {}
          return dialogPush_(buildChatTaskSearchCard_(spaceId, memory, {
            search_term: argumentText, project_id: projectId,
            results: taskData.tasks || [], total: taskData.total || 0,
            offset: 0, searched: true
          }));
        }
        return dialogPush_(buildChatTaskSearchCard_(spaceId, memory, {}));
      }
      case CHAT_CMD_TICKET: {
        if (argumentText) {
          var scopedTeamIds = parseCsvIds_(memory.team_ids_csv);
          var teamId = memory.team_id && (!scopedTeamIds.length || scopedTeamIds.indexOf(memory.team_id) >= 0)
                       ? memory.team_id : '';
          var ticketData = { tickets: [], total: 0 };
          try {
            ticketData = apiTicketSearch_({
              search_term: argumentText,
              team_id: teamId ? parseInt(teamId, 10) : null,
              limit: 10, offset: 0
            });
            rememberSpaceRecentTickets_(spaceId, ticketData.tickets || []);
          } catch (_) {}
          return dialogPush_(buildChatTicketSearchCard_(spaceId, memory, {
            search_term: argumentText, team_id: teamId,
            results: ticketData.tickets || [], total: ticketData.total || 0,
            offset: 0, searched: true
          }));
        }
        return dialogPush_(buildChatTicketSearchCard_(spaceId, memory, {}));
      }
      case CHAT_CMD_TASK_CREATE:
        return dialogPush_(buildChatTaskCreateCard_(spaceId, memory,
          argumentText ? { description: argumentText } : {}));
      case CHAT_CMD_TICKET_CREATE:
        return dialogPush_(buildChatTicketCreateCard_(spaceId, memory,
          argumentText ? { description: argumentText } : {}));
      default:
        console.log('onAppCommand: unrecognized cmdId=' + cmdId);
        return dialogPush_({ sections: [{ widgets: [{ textParagraph: { text: 'Unknown command.' } }] }] });
    }
  } catch (err) {
    console.log('onAppCommand error:', err && err.stack ? err.stack : String(err));
    return dialogPush_({ sections: [{ widgets: [{ textParagraph: { text: 'Error: ' + (err.message || String(err)) } }] }] });
  }
}

// ─── DIALOG RESPONSE HELPERS ─────────────────────────────────────────────────
// pushCard  → use for initial dialog open (REQUEST_DIALOG from onAppCommand)
// updateCard → use for all in-dialog button/action responses

function dialogPush_(card) {
  return { renderActions: { action: { navigations: [{ pushCard: card }] } } };
}

function dialogUpdate_(card) {
  return { renderActions: { action: { navigations: [{ updateCard: card }] } } };
}

function dialogClose_() {
  return { renderActions: { action: { navigations: [{ endNavigation: { action: 'CLOSE_DIALOG' } }] } } };
}

function dialogNotify_(text) {
  return { renderActions: { action: { notification: { text: text } } } };
}

function onChatCloseDialog() {
  console.log('onChatCloseDialog called');
  return dialogClose_();
}

// ─── ACTION PARAMETER HELPERS ────────────────────────────────────────────────

function toActionParameters_(params) {
  var out = [];
  Object.keys(params || {}).forEach(function(key) {
    var value = params[key];
    if (value !== undefined && value !== null) out.push({ key: key, value: String(value) });
  });
  return out;
}

function getActionParameter_(e, name) {
  var p = (e && e.commonEventObject && e.commonEventObject.parameters) || {};
  if (p[name] !== undefined && p[name] !== null) return String(p[name]);
  var p2 = (e && e.parameters) || {};
  if (p2[name] !== undefined && p2[name] !== null) return String(p2[name]);
  return '';
}

function getFormInputValue_(e, name) {
  var fi = e && e.commonEventObject && e.commonEventObject.formInputs && e.commonEventObject.formInputs[name];
  if (fi && fi.stringInputs && fi.stringInputs.value && fi.stringInputs.value.length) {
    return String(fi.stringInputs.value[0] || '');
  }
  var fi2 = e && e.formInput && e.formInput[name];
  if (Array.isArray(fi2)) return String(fi2[0] || '');
  return fi2 ? String(fi2) : '';
}

function getFormInputValues_(e, name) {
  var fi = e && e.commonEventObject && e.commonEventObject.formInputs && e.commonEventObject.formInputs[name];
  if (fi && fi.stringInputs && fi.stringInputs.value) {
    return fi.stringInputs.value.map(function(v) { return String(v); });
  }
  var fi2 = e && e.formInput && e.formInput[name];
  if (Array.isArray(fi2)) return fi2.map(function(v) { return String(v); });
  if (fi2) return [String(fi2)];
  return [];
}

// ─── DROPDOWN LOADERS (with UserCache) ───────────────────────────────────────

function loadProjects_() {
  var cache = CacheService.getUserCache();
  try { var h = cache.get('dd_project_v1'); if (h) return JSON.parse(h).projects || []; } catch (_) {}
  try {
    var r = odooPost_('/gmail_addon/project/dropdown', {});
    try { cache.put('dd_project_v1', JSON.stringify(r), 180); } catch (_) {}
    return (r && r.projects) || [];
  } catch (_) { return []; }
}

function loadTeams_() {
  var cache = CacheService.getUserCache();
  try { var h = cache.get('dd_team_v1'); if (h) return JSON.parse(h).teams || []; } catch (_) {}
  try {
    var r = odooPost_('/gmail_addon/team/dropdown', {});
    try { cache.put('dd_team_v1', JSON.stringify(r), 180); } catch (_) {}
    return (r && r.teams) || [];
  } catch (_) { return []; }
}

function loadProjectsAndTeams_() {
  var cache = CacheService.getUserCache();
  var pc = null, tc = null;
  try { var ph = cache.get('dd_project_v1'); if (ph) pc = JSON.parse(ph); } catch (_) {}
  try { var th = cache.get('dd_team_v1');    if (th) tc = JSON.parse(th); } catch (_) {}

  var toFetch = [];
  if (!pc) toFetch.push({ path: '/gmail_addon/project/dropdown', params: {} });
  if (!tc) toFetch.push({ path: '/gmail_addon/team/dropdown',    params: {} });

  if (toFetch.length) {
    var results = odooFetchAll_(toFetch);
    var ri = 0;
    if (!pc) { var pr = results[ri++]; if (pr) { pc = pr; try { cache.put('dd_project_v1', JSON.stringify(pr), 180); } catch (_) {} } }
    if (!tc) { var tr = results[ri];   if (tr) { tc = tr; try { cache.put('dd_team_v1',    JSON.stringify(tr), 180); } catch (_) {} } }
  }
  return { projects: (pc && pc.projects) || [], teams: (tc && tc.teams) || [] };
}

// ─── /config ─────────────────────────────────────────────────────────────────

function buildConfigCard_(spaceId, memory, opts) {
  opts = opts || {};
  var data = loadProjectsAndTeams_();

  var selectedProjectIds = opts.project_ids || parseCsvIds_(memory.project_ids_csv);
  var selectedTeamIds    = opts.team_ids    || parseCsvIds_(memory.team_ids_csv);

  var projectItems = data.projects.map(function(p) {
    return { text: p.name, value: String(p.id), selected: selectedProjectIds.indexOf(String(p.id)) >= 0 };
  });
  var teamItems = data.teams.map(function(t) {
    return { text: t.name, value: String(t.id), selected: selectedTeamIds.indexOf(String(t.id)) >= 0 };
  });

  var widgets = [
    { textParagraph: { text: 'Select projects and teams for this space. They scope task/ticket searches and creation.' } }
  ];
  if (projectItems.length) {
    widgets.push({ selectionInput: { name: 'project_ids', label: 'Projects', type: 'MULTI_SELECT', items: projectItems } });
  } else {
    widgets.push({ textParagraph: { text: 'No projects found.' } });
  }
  if (teamItems.length) {
    widgets.push({ selectionInput: { name: 'team_ids', label: 'Teams', type: 'MULTI_SELECT', items: teamItems } });
  } else {
    widgets.push({ textParagraph: { text: 'No helpdesk teams found.' } });
  }
  widgets.push({ buttonList: { buttons: [
    { text: 'Save',  onClick: { action: { function: 'onChatConfigSave', parameters: toActionParameters_({ space_id: spaceId }) } } },
    { text: 'Close', onClick: { action: { function: 'onChatCloseDialog' } } }
  ]}});
  if (opts.info)  widgets.push({ textParagraph: { text: opts.info } });
  if (opts.error) widgets.push({ textParagraph: { text: '<b>Error:</b> ' + opts.error } });

  return { header: { title: 'Space Configuration' }, sections: [{ widgets: widgets }] };
}

function onChatConfigSave(e) {
  var spaceId    = getActionParameter_(e, 'space_id') || getSpaceId_(e);
  var projectIds = getFormInputValues_(e, 'project_ids');
  var teamIds    = getFormInputValues_(e, 'team_ids');
  setSpaceMemory_(spaceId, {
    project_ids_csv: toCsvIds_(projectIds),
    team_ids_csv:    toCsvIds_(teamIds),
    project_id: projectIds.length ? String(projectIds[0]) : '',
    team_id:    teamIds.length    ? String(teamIds[0])    : ''
  });
  return dialogUpdate_(buildConfigCard_(spaceId, getSpaceMemory_(spaceId), {
    project_ids: projectIds, team_ids: teamIds, info: 'Saved.'
  }));
}

// ─── /task (search) ──────────────────────────────────────────────────────────

function buildChatTaskSearchCard_(spaceId, memory, state) {
  state = state || {};
  var searchTerm = state.search_term || '';
  var projectId  = state.project_id !== undefined ? state.project_id : (memory.project_id || '');
  var offset     = parseInt(state.offset || 0, 10) || 0;
  var results    = state.results || [];
  var total      = state.total   || 0;

  var scopedIds = parseCsvIds_(memory.project_ids_csv);
  var projects  = loadProjects_();
  if (scopedIds.length) projects = projects.filter(function(p) { return scopedIds.indexOf(String(p.id)) >= 0; });
  if (projectId && scopedIds.length && scopedIds.indexOf(String(projectId)) < 0) projectId = '';

  var projectItems = [{ text: 'All projects', value: '', selected: !projectId }];
  projects.forEach(function(p) {
    projectItems.push({ text: p.name, value: String(p.id), selected: String(p.id) === String(projectId) });
  });

  var widgets = [
    { textInput: { name: 'search_term', label: 'Task name or number', type: 'SINGLE_LINE', value: searchTerm } },
    { selectionInput: { name: 'project_id', label: 'Project', type: 'DROPDOWN', items: projectItems } },
    { buttonList: { buttons: [
      { text: 'Search', onClick: { action: { function: 'onChatTaskSearch', parameters: toActionParameters_({ space_id: spaceId, offset: '0' }) } } },
      { text: 'Close',  onClick: { action: { function: 'onChatCloseDialog' } } }
    ]}}
  ];

  if (state.error) widgets.push({ textParagraph: { text: '<b>Error:</b> ' + state.error } });

  if (results.length) {
    widgets.push({ divider: {} });
    widgets.push({ textParagraph: { text: '<b>Results (' + total + ')</b>' } });
    results.forEach(function(task) {
      var ref = task.task_number || ('#' + task.id);
      var sub = (task.project_name || '') + (task.stage_name ? ' · ' + task.stage_name : '');
      widgets.push({ decoratedText: {
        startIcon: { knownIcon: 'DESCRIPTION' },
        topLabel: ref,
        text: task.name || '(no title)',
        bottomLabel: sub
      }});
      widgets.push({ buttonList: { buttons: [
        { text: 'Open', onClick: { openLink: { url: task.url || '' } } },
        { text: 'Post link', onClick: { action: {
          function: 'onChatPostLink',
          parameters: toActionParameters_({ space_id: spaceId, url: task.url || '', ref: ref, name: task.name || '', type: 'task' })
        }}}
      ]}});
    });
    var pagerBtns = [];
    if (offset > 0) pagerBtns.push({ text: 'Prev', onClick: { action: { function: 'onChatTaskSearch', parameters: toActionParameters_({
      space_id: spaceId, offset: String(Math.max(0, offset - 10)), search_term: searchTerm, project_id: projectId
    })}}});
    if (offset + results.length < total) pagerBtns.push({ text: 'Next', onClick: { action: { function: 'onChatTaskSearch', parameters: toActionParameters_({
      space_id: spaceId, offset: String(offset + 10), search_term: searchTerm, project_id: projectId
    })}}});
    if (pagerBtns.length) widgets.push({ buttonList: { buttons: pagerBtns } });
  } else if (state.searched) {
    widgets.push({ textParagraph: { text: 'No results found.' } });
  }

  return { header: { title: 'Search Tasks' }, sections: [{ widgets: widgets }] };
}

function onChatTaskSearch(e) {
  var spaceId    = getActionParameter_(e, 'space_id') || getSpaceId_(e);
  var searchTerm = getFormInputValue_(e, 'search_term') || getActionParameter_(e, 'search_term');
  var projectId  = getFormInputValue_(e, 'project_id')  || getActionParameter_(e, 'project_id');
  var offset     = parseInt(getActionParameter_(e, 'offset') || '0', 10) || 0;

  var results = [], total = 0, error = '';
  try {
    var data = apiTaskSearch_({ search_term: searchTerm || '', project_id: projectId ? parseInt(projectId, 10) : null, limit: 10, offset: offset });
    results = (data && data.tasks)  || [];
    total   = (data && data.total)  || 0;
    rememberSpaceRecentTasks_(spaceId, results);
  } catch (err) {
    error = err && err.message ? err.message : String(err);
  }

  if (projectId) setSpaceMemory_(spaceId, { project_id: projectId });

  return dialogUpdate_(buildChatTaskSearchCard_(spaceId, getSpaceMemory_(spaceId), {
    search_term: searchTerm, project_id: projectId,
    offset: offset, results: results, total: total, searched: true, error: error
  }));
}

// ─── /ticket (search) ────────────────────────────────────────────────────────

function buildChatTicketSearchCard_(spaceId, memory, state) {
  state = state || {};
  var searchTerm = state.search_term || '';
  var teamId     = state.team_id !== undefined ? state.team_id : (memory.team_id || '');
  var offset     = parseInt(state.offset || 0, 10) || 0;
  var results    = state.results || [];
  var total      = state.total   || 0;

  var scopedIds = parseCsvIds_(memory.team_ids_csv);
  var teams     = loadTeams_();
  if (scopedIds.length) teams = teams.filter(function(t) { return scopedIds.indexOf(String(t.id)) >= 0; });
  if (teamId && scopedIds.length && scopedIds.indexOf(String(teamId)) < 0) teamId = '';

  var teamItems = [{ text: 'All teams', value: '', selected: !teamId }];
  teams.forEach(function(t) {
    teamItems.push({ text: t.name, value: String(t.id), selected: String(t.id) === String(teamId) });
  });

  var widgets = [
    { textInput: { name: 'search_term', label: 'Ticket subject or ID', type: 'SINGLE_LINE', value: searchTerm } },
    { selectionInput: { name: 'team_id', label: 'Team', type: 'DROPDOWN', items: teamItems } },
    { buttonList: { buttons: [
      { text: 'Search', onClick: { action: { function: 'onChatTicketSearch', parameters: toActionParameters_({ space_id: spaceId, offset: '0' }) } } },
      { text: 'Close',  onClick: { action: { function: 'onChatCloseDialog' } } }
    ]}}
  ];

  if (state.error) widgets.push({ textParagraph: { text: '<b>Error:</b> ' + state.error } });

  if (results.length) {
    widgets.push({ divider: {} });
    widgets.push({ textParagraph: { text: '<b>Results (' + total + ')</b>' } });
    results.forEach(function(ticket) {
      var ref = ticket.ticket_ref || ('#' + ticket.id);
      var sub = (ticket.team_name || '') + (ticket.stage_name ? ' · ' + ticket.stage_name : '');
      widgets.push({ decoratedText: {
        startIcon: { knownIcon: 'CONFIRMATION_NUMBER_ICON' },
        topLabel: ref,
        text: ticket.name || '(no title)',
        bottomLabel: sub
      }});
      widgets.push({ buttonList: { buttons: [
        { text: 'Open', onClick: { openLink: { url: ticket.url || '' } } },
        { text: 'Post link', onClick: { action: {
          function: 'onChatPostLink',
          parameters: toActionParameters_({ space_id: spaceId, url: ticket.url || '', ref: ref, name: ticket.name || '', type: 'ticket' })
        }}}
      ]}});
    });
    var pagerBtns = [];
    if (offset > 0) pagerBtns.push({ text: 'Prev', onClick: { action: { function: 'onChatTicketSearch', parameters: toActionParameters_({
      space_id: spaceId, offset: String(Math.max(0, offset - 10)), search_term: searchTerm, team_id: teamId
    })}}});
    if (offset + results.length < total) pagerBtns.push({ text: 'Next', onClick: { action: { function: 'onChatTicketSearch', parameters: toActionParameters_({
      space_id: spaceId, offset: String(offset + 10), search_term: searchTerm, team_id: teamId
    })}}});
    if (pagerBtns.length) widgets.push({ buttonList: { buttons: pagerBtns } });
  } else if (state.searched) {
    widgets.push({ textParagraph: { text: 'No results found.' } });
  }

  return { header: { title: 'Search Tickets' }, sections: [{ widgets: widgets }] };
}

function onChatTicketSearch(e) {
  var spaceId    = getActionParameter_(e, 'space_id') || getSpaceId_(e);
  var searchTerm = getFormInputValue_(e, 'search_term') || getActionParameter_(e, 'search_term');
  var teamId     = getFormInputValue_(e, 'team_id')     || getActionParameter_(e, 'team_id');
  var offset     = parseInt(getActionParameter_(e, 'offset') || '0', 10) || 0;

  var results = [], total = 0, error = '';
  try {
    var data = apiTicketSearch_({ search_term: searchTerm || '', team_id: teamId ? parseInt(teamId, 10) : null, limit: 10, offset: offset });
    results = (data && data.tickets) || [];
    total   = (data && data.total)   || 0;
    rememberSpaceRecentTickets_(spaceId, results);
  } catch (err) {
    error = err && err.message ? err.message : String(err);
  }

  if (teamId) setSpaceMemory_(spaceId, { team_id: teamId });

  return dialogUpdate_(buildChatTicketSearchCard_(spaceId, getSpaceMemory_(spaceId), {
    search_term: searchTerm, team_id: teamId,
    offset: offset, results: results, total: total, searched: true, error: error
  }));
}

// ─── /task.create ─────────────────────────────────────────────────────────────

function buildChatTaskCreateCard_(spaceId, memory, state) {
  state = state || {};
  var projectId   = state.project_id !== undefined ? state.project_id : (memory.project_id || '');
  var taskName    = state.task_name    || '';
  var description = state.description || '';
  var ccRaw       = state.cc_addresses || '';

  var scopedIds = parseCsvIds_(memory.project_ids_csv);
  var projects  = loadProjects_();
  if (scopedIds.length) {
    projects = projects.filter(function(p) { return scopedIds.indexOf(String(p.id)) >= 0; });
    if (projectId && scopedIds.indexOf(String(projectId)) < 0) projectId = '';
    if (!projectId && projects.length) projectId = String(projects[0].id);
  }

  var projectItems = [{ text: 'Select a project…', value: '', selected: !projectId }];
  projects.forEach(function(p) {
    projectItems.push({ text: p.name, value: String(p.id), selected: String(p.id) === String(projectId) });
  });

  var widgets = [
    { selectionInput: { name: 'project_id',  label: 'Project *',          type: 'DROPDOWN',      items: projectItems } },
    { textInput:      { name: 'task_name',    label: 'Task Name *',        type: 'SINGLE_LINE',   value: taskName } },
    { textInput:      { name: 'description',  label: 'Description',        type: 'MULTIPLE_LINE', value: description } },
    { textInput:      { name: 'cc_addresses', label: 'Followers (emails)', type: 'SINGLE_LINE',   value: ccRaw, hintText: 'Comma-separated' } },
    { buttonList: { buttons: [
      { text: 'Create Task', onClick: { action: { function: 'onChatTaskCreateSubmit', parameters: toActionParameters_({ space_id: spaceId }) } } },
      { text: 'Close',       onClick: { action: { function: 'onChatCloseDialog' } } }
    ]}}
  ];
  if (state.error) widgets.push({ textParagraph: { text: '<b>Error:</b> ' + state.error } });

  return { header: { title: 'Create Task' }, sections: [{ widgets: widgets }] };
}

function onChatTaskCreateSubmit(e) {
  var spaceId     = getActionParameter_(e, 'space_id') || getSpaceId_(e);
  var projectId   = getFormInputValue_(e, 'project_id');
  var taskName    = getFormInputValue_(e, 'task_name');
  var description = getFormInputValue_(e, 'description');
  var ccRaw       = getFormInputValue_(e, 'cc_addresses');
  var memory      = getSpaceMemory_(spaceId);

  if (!projectId || !taskName) {
    return dialogUpdate_(buildChatTaskCreateCard_(spaceId, memory, {
      project_id: projectId, task_name: taskName, description: description, cc_addresses: ccRaw,
      error: 'Project and Task Name are required.'
    }));
  }

  try {
    var result = apiCreateTask_({
      project_id:   parseInt(projectId, 10),
      name:         taskName,
      description:  description || '',
      cc_addresses: extractEmailsFromAddressList_(ccRaw || '')
    });

    setSpaceMemory_(spaceId, { project_id: projectId });
    var task = {
      id: result.task_id || '', task_number: result.task_number || ('#' + (result.task_id || '')),
      name: taskName, project_name: '', stage_name: '',
      url: result.task_url || '', write_date: new Date().toISOString()
    };
    rememberSpaceRecentTasks_(spaceId, [task]);

    return dialogUpdate_({ header: { title: 'Task Created' }, sections: [{ widgets: [
      { decoratedText: {
        startIcon: { knownIcon: 'DESCRIPTION' }, topLabel: task.task_number,
        text: taskName, bottomLabel: 'Task created successfully'
      }},
      { buttonList: { buttons: [
        { text: 'Open',      onClick: { openLink: { url: task.url || '' } } },
        { text: 'Post link', onClick: { action: { function: 'onChatPostLink', parameters: toActionParameters_({ space_id: spaceId, url: task.url || '', ref: task.task_number || '', name: taskName, type: 'task' }) } } },
        { text: 'Close',     onClick: { action: { function: 'onChatCloseDialog' } } }
      ]}}
    ]}]});
  } catch (err) {
    return dialogUpdate_(buildChatTaskCreateCard_(spaceId, memory, {
      project_id: projectId, task_name: taskName, description: description, cc_addresses: ccRaw,
      error: err && err.message ? err.message : String(err)
    }));
  }
}

// ─── /ticket.create ──────────────────────────────────────────────────────────

function buildChatTicketCreateCard_(spaceId, memory, state) {
  state = state || {};
  var teamId      = state.team_id !== undefined ? state.team_id : (memory.team_id || '');
  var ticketName  = state.ticket_name  || '';
  var description = state.description || '';
  var ccRaw       = state.cc_addresses || '';

  var scopedIds = parseCsvIds_(memory.team_ids_csv);
  var teams     = loadTeams_();
  if (scopedIds.length) {
    teams = teams.filter(function(t) { return scopedIds.indexOf(String(t.id)) >= 0; });
    if (teamId && scopedIds.indexOf(String(teamId)) < 0) teamId = '';
    if (!teamId && teams.length) teamId = String(teams[0].id);
  }

  var teamItems = [{ text: 'Select a team…', value: '', selected: !teamId }];
  teams.forEach(function(t) {
    teamItems.push({ text: t.name, value: String(t.id), selected: String(t.id) === String(teamId) });
  });

  var widgets = [
    { selectionInput: { name: 'team_id',      label: 'Team *',             type: 'DROPDOWN',      items: teamItems } },
    { textInput:      { name: 'ticket_name',   label: 'Subject *',          type: 'SINGLE_LINE',   value: ticketName } },
    { textInput:      { name: 'description',   label: 'Description',        type: 'MULTIPLE_LINE', value: description } },
    { textInput:      { name: 'cc_addresses',  label: 'Followers (emails)', type: 'SINGLE_LINE',   value: ccRaw, hintText: 'Comma-separated' } },
    { buttonList: { buttons: [
      { text: 'Create Ticket', onClick: { action: { function: 'onChatTicketCreateSubmit', parameters: toActionParameters_({ space_id: spaceId }) } } },
      { text: 'Close',         onClick: { action: { function: 'onChatCloseDialog' } } }
    ]}}
  ];
  if (state.error) widgets.push({ textParagraph: { text: '<b>Error:</b> ' + state.error } });

  return { header: { title: 'Create Ticket' }, sections: [{ widgets: widgets }] };
}

function onChatTicketCreateSubmit(e) {
  var spaceId     = getActionParameter_(e, 'space_id') || getSpaceId_(e);
  var teamId      = getFormInputValue_(e, 'team_id');
  var ticketName  = getFormInputValue_(e, 'ticket_name');
  var description = getFormInputValue_(e, 'description');
  var ccRaw       = getFormInputValue_(e, 'cc_addresses');
  var memory      = getSpaceMemory_(spaceId);

  if (!teamId || !ticketName) {
    return dialogUpdate_(buildChatTicketCreateCard_(spaceId, memory, {
      team_id: teamId, ticket_name: ticketName, description: description, cc_addresses: ccRaw,
      error: 'Team and Subject are required.'
    }));
  }

  try {
    var result = apiCreateTicket_({
      team_id:      parseInt(teamId, 10),
      name:         ticketName,
      description:  description || '',
      cc_addresses: extractEmailsFromAddressList_(ccRaw || '')
    });

    setSpaceMemory_(spaceId, { team_id: teamId });
    var ticket = {
      id: result.ticket_id || '', ticket_ref: result.ticket_ref || ('#' + (result.ticket_id || '')),
      name: ticketName, team_name: '', stage_name: '',
      url: result.ticket_url || '', write_date: new Date().toISOString()
    };
    rememberSpaceRecentTickets_(spaceId, [ticket]);

    return dialogUpdate_({ header: { title: 'Ticket Created' }, sections: [{ widgets: [
      { decoratedText: {
        startIcon: { knownIcon: 'CONFIRMATION_NUMBER_ICON' }, topLabel: ticket.ticket_ref,
        text: ticketName, bottomLabel: 'Ticket created successfully'
      }},
      { buttonList: { buttons: [
        { text: 'Open',      onClick: { openLink: { url: ticket.url || '' } } },
        { text: 'Post link', onClick: { action: { function: 'onChatPostLink', parameters: toActionParameters_({ space_id: spaceId, url: ticket.url || '', ref: ticket.ticket_ref || '', name: ticketName, type: 'ticket' }) } } },
        { text: 'Close',     onClick: { action: { function: 'onChatCloseDialog' } } }
      ]}}
    ]}]});
  } catch (err) {
    return dialogUpdate_(buildChatTicketCreateCard_(spaceId, memory, {
      team_id: teamId, ticket_name: ticketName, description: description, cc_addresses: ccRaw,
      error: err && err.message ? err.message : String(err)
    }));
  }
}

// ─── POST LINK TO CHAT ───────────────────────────────────────────────────────

/**
 * Posts a message body to a Chat space via REST API (user credentials).
 * Supports text messages. Cards require service account credentials
 * and are not supported in Workspace Add-on context.
 * @param {string} spaceName  e.g. 'spaces/hTD8mSAAAAE'
 * @param {Object} body       Chat message body (text, etc.)
 */
function postChatMessage_(spaceName, body) {
  var token = ScriptApp.getOAuthToken();
  var resp = UrlFetchApp.fetch(
    'https://chat.googleapis.com/v1/' + spaceName + '/messages',
    {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    }
  );
  var code = resp.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('HTTP ' + code + ': ' + resp.getContentText().substring(0, 200));
  }
}

/**
 * Builds a Chat message body for sharing a task or ticket link.
 * Uses plain text with Chat markdown — cards require service account credentials
 * which are not available in Workspace Add-on context.
 * @param {string} type  'task' or 'ticket'
 * @param {string} ref   e.g. 'TASK-42' or '#2641'
 * @param {string} name  record name
 * @param {string} url   Odoo URL
 */
function buildShareCardPayload_(type, ref, name, url) {
  var icon  = type === 'ticket' ? '🎫' : '📄';
  var label = type === 'ticket' ? '*[TICKET]*' : '*[TASK]*';
  var refLink = ref && url ? '<' + url + '|' + ref + '>' : (ref || '');
  return { text: icon + ' ' + label + ' - ' + refLink + (name ? ', _' + name + '_' : '') };
}

function onChatPostLink(e) {
  console.log('onChatPostLink called');
  var spaceId   = getActionParameter_(e, 'space_id');
  var url       = getActionParameter_(e, 'url');
  var ref       = getActionParameter_(e, 'ref');
  var name      = getActionParameter_(e, 'name');
  var type      = getActionParameter_(e, 'type') || 'task';
  var spaceName = spaceId && spaceId !== 'default' ? ('spaces/' + spaceId) : '';

  if (!spaceName || !url) {
    return dialogUpdate_({ header: { title: 'Post Link' }, sections: [{ widgets: [
      { textParagraph: { text: 'Cannot post: missing space or URL.' } },
      { buttonList: { buttons: [{ text: 'Close', onClick: { action: { function: 'onChatCloseDialog' } } }] } }
    ]}]});
  }
  try {
    postChatMessage_(spaceName, buildShareCardPayload_(type, ref, name, url));
  } catch (err) {
    return dialogUpdate_({ header: { title: 'Post Link' }, sections: [{ widgets: [
      { textParagraph: { text: 'Error: ' + (err.message || String(err)) } },
      { buttonList: { buttons: [{ text: 'Close', onClick: { action: { function: 'onChatCloseDialog' } } }] } }
    ]}]});
  }
  return dialogClose_();
}

// ─── SPACE RECENT MEMORY ─────────────────────────────────────────────────────

function getSpaceRecentTasks_(spaceId) {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty('space_' + spaceId + '_recent_tasks_json');
    var arr = raw ? JSON.parse(raw) : [];
    return Array.isArray(arr) ? arr : [];
  } catch (_) { return []; }
}

function getSpaceRecentTickets_(spaceId) {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty('space_' + spaceId + '_recent_tickets_json');
    var arr = raw ? JSON.parse(raw) : [];
    return Array.isArray(arr) ? arr : [];
  } catch (_) { return []; }
}

function rememberSpaceRecentTasks_(spaceId, tasks) {
  if (!spaceId || !tasks || !tasks.length) return;
  var byId = {};
  getSpaceRecentTasks_(spaceId).forEach(function(t) { byId[String(t.id)] = t; });
  tasks.forEach(function(t) {
    if (!t || !t.id) return;
    byId[String(t.id)] = { id: t.id, task_number: t.task_number || '', name: t.name || '',
      project_name: t.project_name || '', stage_name: t.stage_name || '',
      url: t.url || '', write_date: t.write_date || '' };
  });
  var merged = Object.keys(byId).map(function(k) { return byId[k]; });
  merged.sort(function(a, b) { return a.write_date > b.write_date ? -1 : a.write_date < b.write_date ? 1 : 0; });
  PropertiesService.getScriptProperties().setProperty('space_' + spaceId + '_recent_tasks_json', JSON.stringify(merged.slice(0, 20)));
}

function rememberSpaceRecentTickets_(spaceId, tickets) {
  if (!spaceId || !tickets || !tickets.length) return;
  var byId = {};
  getSpaceRecentTickets_(spaceId).forEach(function(t) { byId[String(t.id)] = t; });
  tickets.forEach(function(t) {
    if (!t || !t.id) return;
    byId[String(t.id)] = { id: t.id, ticket_ref: t.ticket_ref || '', name: t.name || '',
      team_name: t.team_name || '', stage_name: t.stage_name || '',
      url: t.url || '', write_date: t.write_date || '' };
  });
  var merged = Object.keys(byId).map(function(k) { return byId[k]; });
  merged.sort(function(a, b) { return a.write_date > b.write_date ? -1 : a.write_date < b.write_date ? 1 : 0; });
  PropertiesService.getScriptProperties().setProperty('space_' + spaceId + '_recent_tickets_json', JSON.stringify(merged.slice(0, 20)));
}

// ─── HOME CARD (sidebar, uses CardService) ───────────────────────────────────

function buildChatHomeCard_(e) {
  var spaceId    = getSpaceId_(e);
  var spaceName  = (e && e.chat && e.chat.space && e.chat.space.displayName) || '';
  var threadName = (e && e.chat && e.chat.message && e.chat.message.thread && e.chat.message.thread.name) || '';

  var fetchResults = odooFetchAll_([
    { path: '/gmail_addon/task/search',   params: { limit: 5, offset: 0 } },
    { path: '/gmail_addon/ticket/search', params: { limit: 5, offset: 0 } }
  ]);
  var recentTasks   = (fetchResults[0] && fetchResults[0].tasks)   || [];
  var recentTickets = (fetchResults[1] && fetchResults[1].tickets) || [];

  var recentItems = [];
  recentTasks.forEach(function(t)   { recentItems.push({ type: 'task',   data: t, write_date: t.write_date || '' }); });
  recentTickets.forEach(function(t) { recentItems.push({ type: 'ticket', data: t, write_date: t.write_date || '' }); });
  recentItems.sort(function(a, b) { return a.write_date > b.write_date ? -1 : a.write_date < b.write_date ? 1 : 0; });

  var cardBuilder = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Odoo Tasks & Tickets').setSubtitle(spaceName || ''));

  var recentSection = CardService.newCardSection().setHeader('Recent tasks & tickets');
  if (!recentItems.length) {
    recentSection.addWidget(CardService.newTextParagraph().setText('No recent records found.'));
  } else {
    recentItems.forEach(function(item) {
      var d      = item.data;
      var isTask = item.type === 'task';
      var ref    = isTask ? (d.task_number || ('#' + d.id)) : (d.ticket_ref || ('#' + d.id));
      var scope  = isTask ? (d.project_name || '') : (d.team_name || '');
      var assignee = d.user_name ? ' · ' + d.user_name : '';
      recentSection.addWidget(CardService.newDecoratedText()
        .setTopLabel(ref + ' · ' + scope)
        .setText(d.name)
        .setBottomLabel((d.stage_name || '') + assignee)
        .setStartIcon(CardService.newIconImage().setIcon(
          isTask ? CardService.Icon.DESCRIPTION : CardService.Icon.CONFIRMATION_NUMBER_ICON
        ))
        .setOnClickAction(CardService.newAction().setFunctionName('onChatOpenLink').setParameters({ url: d.url }))
        .setButton(CardService.newImageButton()
          .setIconUrl('https://www.gstatic.com/images/icons/material/system/1x/insert_link_black_18dp.png')
          .setAltText('Post link to Chat')
          .setOnClickAction(CardService.newAction().setFunctionName('onChatPostLinkFromSidebar')
            .setParameters({ reference: ref, name: d.name, type: item.type, url: d.url, space_id: spaceId })
          )
        )
      );
    });
  }

  var navSection = CardService.newCardSection();
  navSection.addWidget(CardService.newTextButton()
    .setText('Settings')
    .setOnClickAction(CardService.newAction().setFunctionName('onChatNavigateToSettings'))
  );

  return cardBuilder.addSection(recentSection).addSection(navSection).build();
}

function onChatOpenLink(e) {
  var url = (e && e.parameters && e.parameters.url) || '';
  return CardService.newActionResponseBuilder()
    .setOpenLink(CardService.newOpenLink().setUrl(url))
    .build();
}

function onChatPostLinkFromSidebar(e) {
  var p       = e.parameters || {};
  var spaceId = p.space_id   || '';
  var ref     = p.reference  || '';
  var name    = p.name       || '';
  var type    = p.type       || 'task';
  var url     = p.url        || '';
  var spaceName = spaceId && spaceId !== 'default' ? ('spaces/' + spaceId) : '';
  try {
    postChatMessage_(spaceName, buildShareCardPayload_(type, ref, name, url));
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Link posted.'))
      .build();
  } catch (err) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Failed: ' + (err.message || String(err))))
      .build();
  }
}

function onChatNavigateToSettings(e) {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(buildLoginCard_()))
    .build();
}
