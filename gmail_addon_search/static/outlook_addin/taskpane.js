const RECORD_LABELS = {
  task: "Task",
  ticket: "Ticket",
  lead: "Lead",
};

const SEARCH_ENDPOINTS = {
  task: "/gmail_addon/task/search",
  ticket: "/gmail_addon/ticket/search",
  lead: "/gmail_addon/lead/search",
};

const CREATE_ENDPOINTS = {
  task: "/gmail_addon/task/create",
  ticket: "/gmail_addon/ticket/create",
  lead: "/gmail_addon/lead/create",
};

const MODEL_BY_TYPE = {
  task: "project.task",
  ticket: "helpdesk.ticket",
  lead: "crm.lead",
};

const DEFAULT_CONFIG = {
  odooUrl: "",
  token: "",
};

const DEFAULT_MAILBOX = {
  mode: "read",
  senderEmail: "",
  senderName: "",
  subject: "",
  bodyHtml: "",
  ccAddresses: "",
  itemId: "",
  conversationId: "",
  internetMessageId: "",
};

const state = {
  config: { ...DEFAULT_CONFIG },
  connected: false,
  loading: true,
  busy: "",
  error: "",
  success: "",
  view: "home",
  initialized: false,
  mailbox: { ...DEFAULT_MAILBOX },
  context: null,
  linkedRecords: [],
  search: {
    recordType: "task",
    query: "",
    filters: {},
    results: [],
    total: 0,
  },
  create: {
    recordType: "task",
    form: {},
    result: null,
  },
  schemas: {},
  lookups: {
    projects: [],
    taskStages: {},
    ticketTeams: [],
    ticketStages: {},
    crmTeams: [],
    crmStages: {},
    users: [],
  },
  partnerPicker: {
    searchTerm: "",
    results: [],
    createName: "",
    createEmail: "",
    createCompany: "",
  },
};

const appEl = document.getElementById("app");

document.addEventListener("click", onClick);
document.addEventListener("submit", onSubmit);
document.addEventListener("change", onChange);

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Outlook) {
    state.error = "This add-in only runs inside Outlook.";
    state.loading = false;
    render();
    return;
  }

  state.config = loadConfig();
  render();

  try {
    state.mailbox = await getMailboxContext();
    hydrateDefaults();
    if (isConfigured()) {
      await connectAndRefresh();
    }
  } catch (error) {
    state.error = error.message || "Failed to read Outlook message context.";
  } finally {
    state.loading = false;
    state.initialized = true;
    render();
  }
});

function loadConfig() {
  const roaming = getRoamingSettings();
  const roamingUrl = roaming ? roaming.get("odooUrl") || "" : "";
  const roamingToken = roaming ? roaming.get("odooToken") || "" : "";
  return {
    odooUrl: (roamingUrl || window.localStorage.getItem("outlook.odooUrl") || "").replace(/\/$/, ""),
    token: roamingToken || window.localStorage.getItem("outlook.odooToken") || "",
  };
}

async function saveConfig(config) {
  const roaming = getRoamingSettings();
  state.config = {
    odooUrl: (config.odooUrl || "").replace(/\/$/, ""),
    token: config.token || "",
  };

  if (roaming) {
    roaming.set("odooUrl", state.config.odooUrl);
    roaming.set("odooToken", state.config.token);
    await saveRoamingSettings(roaming);
  }

  window.localStorage.setItem("outlook.odooUrl", state.config.odooUrl);
  window.localStorage.setItem("outlook.odooToken", state.config.token);
}

async function clearConfig() {
  const roaming = getRoamingSettings();
  if (roaming) {
    roaming.remove("odooUrl");
    roaming.remove("odooToken");
    await saveRoamingSettings(roaming);
  }
  window.localStorage.removeItem("outlook.odooUrl");
  window.localStorage.removeItem("outlook.odooToken");
  state.config = { ...DEFAULT_CONFIG };
  state.connected = false;
}

function getRoamingSettings() {
  try {
    return Office.context.roamingSettings || null;
  } catch (_) {
    return null;
  }
}

function saveRoamingSettings(roaming) {
  return new Promise((resolve, reject) => {
    roaming.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
      } else {
        reject(new Error(result.error && result.error.message ? result.error.message : "Failed to save Outlook settings."));
      }
    });
  });
}

function isConfigured() {
  return Boolean(state.config.odooUrl && state.config.token);
}

async function connectAndRefresh() {
  setBusy("Connecting to Odoo…");
  try {
    await odooPost("/gmail_addon/ping", {});
    state.connected = true;
    state.error = "";
    await Promise.all([
      loadHomeData(),
      ensureLookups("task"),
      ensureLookups("ticket"),
      ensureLookups("lead"),
      ensureSchema("task"),
      ensureSchema("ticket"),
      ensureSchema("lead"),
    ]);
    hydrateDefaults();
  } catch (error) {
    state.connected = false;
    state.error = error.message || "Failed to connect to Odoo.";
  } finally {
    clearBusy();
    render();
  }
}

async function loadHomeData() {
  if (!isConfigured()) {
    state.context = null;
    state.linkedRecords = [];
    return;
  }

  const senderEmail = state.mailbox.senderEmail || "";
  const [contextResult, linkedResult] = await Promise.all([
    odooPost("/gmail_addon/suggest_context", { sender_email: senderEmail, filter_mine: false }),
    odooPost("/gmail_addon/email/linked_records", {
      rfc_message_id: state.mailbox.internetMessageId || "",
      outlook_item_id: state.mailbox.itemId || "",
      outlook_conversation_id: state.mailbox.conversationId || "",
    }),
  ]);

  state.context = contextResult || null;
  state.linkedRecords = (linkedResult && linkedResult.records) || [];
}

async function ensureLookups(recordType) {
  await ensureUsers();
  if (recordType === "task") {
    await ensureProjects();
    const projectId = state.search.recordType === "task" ? state.search.filters.project_id : state.create.form.project_id;
    await ensureTaskStages(projectId || "");
  } else if (recordType === "ticket") {
    await ensureTicketTeams();
    const teamId = state.search.recordType === "ticket" ? state.search.filters.team_id : state.create.form.team_id;
    await ensureTicketStages(teamId || "");
  } else if (recordType === "lead") {
    await ensureCrmTeams();
    const teamId = state.search.recordType === "lead" ? state.search.filters.team_id : state.create.form.team_id;
    await ensureCrmStages(teamId || "");
  }
}

async function ensureUsers() {
  if (state.lookups.users.length) {
    return;
  }
  const response = await odooPost("/gmail_addon/user/dropdown", {});
  state.lookups.users = response.users || [];
}

async function ensureProjects() {
  if (state.lookups.projects.length) {
    return;
  }
  const response = await odooPost("/gmail_addon/project/dropdown", {});
  state.lookups.projects = response.projects || [];
}

async function ensureTicketTeams() {
  if (state.lookups.ticketTeams.length) {
    return;
  }
  const response = await odooPost("/gmail_addon/team/dropdown", {});
  state.lookups.ticketTeams = response.teams || [];
}

async function ensureCrmTeams() {
  if (state.lookups.crmTeams.length) {
    return;
  }
  const response = await odooPost("/gmail_addon/crm/team/dropdown", {});
  state.lookups.crmTeams = response.teams || [];
}

async function ensureTaskStages(projectId) {
  const key = projectId || "__all__";
  if (state.lookups.taskStages[key]) {
    return;
  }
  const response = await odooPost("/gmail_addon/stage/dropdown", {
    project_id: projectId || null,
    record_type: "task",
  });
  state.lookups.taskStages[key] = response.stages || [];
}

async function ensureTicketStages(teamId) {
  const key = teamId || "__all__";
  if (state.lookups.ticketStages[key]) {
    return;
  }
  const response = await odooPost("/gmail_addon/stage/dropdown", {
    team_id: teamId || null,
    record_type: "ticket",
  });
  state.lookups.ticketStages[key] = response.stages || [];
}

async function ensureCrmStages(teamId) {
  const key = teamId || "__all__";
  if (state.lookups.crmStages[key]) {
    return;
  }
  const response = await odooPost("/gmail_addon/crm/stage/dropdown", {
    team_id: teamId || null,
  });
  state.lookups.crmStages[key] = response.stages || [];
}

async function ensureSchema(recordType) {
  if (state.schemas[recordType]) {
    return;
  }
  state.schemas[recordType] = await odooPost("/gmail_addon/form/schema", {
    record_type: recordType,
  });
}

function hydrateDefaults() {
  const cleanSubject = normalizeSubject(state.mailbox.subject);
  const bodyPreview = trimText(htmlToText(state.mailbox.bodyHtml), 800);
  const context = state.context || {};

  if (!state.search.filters || Object.keys(state.search.filters).length === 0) {
    state.search.filters = defaultSearchFilters(state.search.recordType, context);
  }

  if (!state.create.form || Object.keys(state.create.form).length === 0) {
    state.create.form = defaultCreateForm(state.create.recordType, context, cleanSubject, bodyPreview);
  }

  if (!state.partnerPicker.createName) {
    state.partnerPicker.createName = state.mailbox.senderName || "";
    state.partnerPicker.createEmail = state.mailbox.senderEmail || "";
  }
}

function defaultSearchFilters(recordType, context) {
  if (recordType === "task") {
    return {
      project_id: stringValue(context.suggested_project_id),
      stage_id: "",
      user_id: "",
    };
  }
  if (recordType === "ticket") {
    return {
      team_id: stringValue(context.suggested_team_id),
      stage_id: "",
      user_id: "",
    };
  }
  return {
    lead_type: "all",
    team_id: stringValue(context.suggested_crm_team_id),
    stage_id: "",
    user_id: "",
  };
}

function defaultCreateForm(recordType, context, subject, description) {
  const base = {
    name: subject || "",
    description: description || "",
    partner_id: stringValue(context.partner_id),
    partner_label: context.partner_name || "",
  };

  if (recordType === "task") {
    return {
      ...base,
      project_id: stringValue(context.suggested_project_id),
      user_id: "",
      extraValues: {},
    };
  }

  if (recordType === "ticket") {
    return {
      ...base,
      team_id: stringValue(context.suggested_team_id),
      priority: "1",
      extraValues: {},
    };
  }

  return {
    ...base,
    lead_type: "lead",
    team_id: stringValue(context.suggested_crm_team_id),
    contact_name: state.mailbox.senderName || "",
    partner_name: context.company_partner_name || "",
    email_from: state.mailbox.senderEmail || "",
    extraValues: {},
  };
}

function setBusy(message) {
  state.busy = message;
  render();
}

function clearBusy() {
  state.busy = "";
}

async function odooPost(path, params) {
  if (!isConfigured()) {
    throw new Error("Odoo URL and API key are required.");
  }

  const response = await fetch(`${state.config.odooUrl}${path}`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${state.config.token}`,
    },
    body: JSON.stringify({
      jsonrpc: "2.0",
      method: "call",
      id: Date.now(),
      params: params || {},
    }),
  });

  if (!response.ok) {
    throw new Error(`HTTP ${response.status} from ${path}`);
  }

  const payload = await response.json();
  if (payload.error) {
    const message = payload.error.data && payload.error.data.message
      ? payload.error.data.message
      : payload.error.message || "Unexpected Odoo error.";
    throw new Error(message);
  }
  return payload.result;
}

async function getMailboxContext() {
  const item = Office.context.mailbox.item;
  if (!item) {
    return { ...DEFAULT_MAILBOX };
  }

  const [subject, bodyHtml, cc, internetMessageId] = await Promise.all([
    readItemSubject(item),
    readItemBody(item),
    readRecipients(item, "cc"),
    readInternetMessageId(item),
  ]);

  const sender = readSender(item);
  return {
    mode: isComposeItem(item) ? "compose" : "read",
    senderEmail: sender.email || "",
    senderName: sender.name || "",
    subject: subject || "",
    bodyHtml: bodyHtml || "",
    ccAddresses: cc.join(", "),
    itemId: item.itemId || "",
    conversationId: item.conversationId || "",
    internetMessageId: internetMessageId || "",
  };
}

function isComposeItem(item) {
  return item && typeof item.subject !== "string";
}

function readSender(item) {
  if (item.from && item.from.emailAddress) {
    return {
      email: item.from.emailAddress,
      name: item.from.displayName || item.from.emailAddress,
    };
  }
  return {
    email: Office.context.mailbox.userProfile.emailAddress || "",
    name: Office.context.mailbox.userProfile.displayName || "",
  };
}

async function readItemSubject(item) {
  if (!item) {
    return "";
  }
  if (typeof item.subject === "string") {
    return item.subject;
  }
  if (item.subject && typeof item.subject.getAsync === "function") {
    return officeAsync(item.subject, "getAsync");
  }
  return "";
}

async function readItemBody(item) {
  if (!item || !item.body || typeof item.body.getAsync !== "function") {
    return "";
  }
  try {
    return await officeAsync(item.body, "getAsync", Office.CoercionType.Html);
  } catch (_) {
    return "";
  }
}

async function readRecipients(item, kind) {
  if (!item || !item[kind]) {
    return [];
  }
  const target = item[kind];
  if (Array.isArray(target)) {
    return target.map((entry) => entry.emailAddress || "").filter(Boolean);
  }
  if (typeof target.getAsync === "function") {
    const recipients = await officeAsync(target, "getAsync");
    return (recipients || []).map((entry) => entry.emailAddress || "").filter(Boolean);
  }
  return [];
}

async function readInternetMessageId(item) {
  if (!item) {
    return "";
  }
  if (item.internetMessageId) {
    return item.internetMessageId;
  }
  if (typeof item.getAllInternetHeadersAsync === "function") {
    try {
      const headers = await officeAsync(item, "getAllInternetHeadersAsync");
      const match = String(headers || "").match(/^message-id:\s*(.+)$/im);
      return match ? match[1].trim() : "";
    } catch (_) {
      return "";
    }
  }
  return "";
}

function officeAsync(target, methodName, ...args) {
  return new Promise((resolve, reject) => {
    target[methodName](...args, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error(result.error && result.error.message ? result.error.message : `${methodName} failed.`));
      }
    });
  });
}

async function onSubmit(event) {
  const form = event.target;
  if (!(form instanceof HTMLFormElement)) {
    return;
  }

  if (form.id === "settings-form") {
    event.preventDefault();
    await saveSettings(form);
    return;
  }

  if (form.id === "search-form") {
    event.preventDefault();
    await performSearch(form);
    return;
  }

  if (form.id === "create-form") {
    event.preventDefault();
    await createRecord(form);
    return;
  }
}

async function onClick(event) {
  const trigger = event.target.closest("[data-action]");
  if (!trigger) {
    return;
  }

  syncDraftsFromDom();
  const action = trigger.dataset.action;

  if (action === "navigate") {
    event.preventDefault();
    await navigate(trigger.dataset.view);
    return;
  }

  if (action === "refresh-home") {
    event.preventDefault();
    await refreshHome();
    return;
  }

  if (action === "open-record") {
    event.preventDefault();
    openExternal(trigger.dataset.url);
    return;
  }

  if (action === "log-record") {
    event.preventDefault();
    await logEmailToRecord(trigger.dataset.model, trigger.dataset.id);
    return;
  }

  if (action === "insert-body") {
    event.preventDefault();
    await insertReference(trigger.dataset.reference, "body");
    return;
  }

  if (action === "insert-subject") {
    event.preventDefault();
    await insertReference(trigger.dataset.reference, "subject");
    return;
  }

  if (action === "select-search-type") {
    event.preventDefault();
    await switchSearchType(trigger.dataset.recordType);
    return;
  }

  if (action === "select-create-type") {
    event.preventDefault();
    await switchCreateType(trigger.dataset.recordType);
    return;
  }

  if (action === "search-partner") {
    event.preventDefault();
    await searchPartners();
    return;
  }

  if (action === "select-partner") {
    event.preventDefault();
    selectPartner(trigger.dataset.partnerId, trigger.dataset.partnerName);
    return;
  }

  if (action === "clear-partner") {
    event.preventDefault();
    clearPartnerSelection();
    return;
  }

  if (action === "create-partner") {
    event.preventDefault();
    await createPartner();
    return;
  }

  if (action === "clear-settings") {
    event.preventDefault();
    await clearSettingsAction();
  }
}

async function onChange(event) {
  const target = event.target;
  if (!(target instanceof HTMLElement)) {
    return;
  }

  syncDraftsFromDom();
  if (target.id === "search-project") {
    state.search.filters.project_id = target.value;
    state.search.filters.stage_id = "";
    await ensureTaskStages(target.value);
    render();
    return;
  }

  if (target.id === "search-team") {
    state.search.filters.team_id = target.value;
    state.search.filters.stage_id = "";
    await ensureTicketStages(target.value);
    render();
    return;
  }

  if (target.id === "search-crm-team") {
    state.search.filters.team_id = target.value;
    state.search.filters.stage_id = "";
    await ensureCrmStages(target.value);
    render();
    return;
  }

  if (target.id === "create-project") {
    state.create.form.project_id = target.value;
    await ensureTaskStages(target.value);
    render();
    return;
  }

  if (target.id === "create-team") {
    state.create.form.team_id = target.value;
    await ensureTicketStages(target.value);
    render();
    return;
  }

  if (target.id === "create-crm-team") {
    state.create.form.team_id = target.value;
    await ensureCrmStages(target.value);
    render();
    return;
  }
}

async function navigate(view) {
  state.view = view;
  state.error = "";
  state.success = "";
  if (view === "home" && isConfigured()) {
    await refreshHome();
  }
  if (view === "search" && isConfigured()) {
    await ensureLookups(state.search.recordType);
  }
  if (view === "create" && isConfigured()) {
    await Promise.all([ensureLookups(state.create.recordType), ensureSchema(state.create.recordType)]);
  }
  render();
}

async function refreshHome() {
  if (!isConfigured()) {
    state.view = "settings";
    render();
    return;
  }
  setBusy("Refreshing mailbox context…");
  try {
    state.mailbox = await getMailboxContext();
    await loadHomeData();
    hydrateDefaults();
    state.success = "Mailbox context refreshed.";
    state.error = "";
  } catch (error) {
    state.error = error.message || "Failed to refresh Outlook context.";
  } finally {
    clearBusy();
    render();
  }
}

async function saveSettings(form) {
  const data = new FormData(form);
  const nextConfig = {
    odooUrl: String(data.get("odooUrl") || "").trim(),
    token: String(data.get("token") || "").trim(),
  };

  setBusy("Saving settings…");
  try {
    await saveConfig(nextConfig);
    await connectAndRefresh();
    state.success = "Connection saved.";
    state.view = "home";
  } catch (error) {
    state.error = error.message || "Failed to save settings.";
  } finally {
    clearBusy();
    render();
  }
}

async function clearSettingsAction() {
  setBusy("Clearing settings…");
  try {
    await clearConfig();
    state.context = null;
    state.linkedRecords = [];
    state.success = "Connection cleared.";
    state.view = "settings";
  } catch (error) {
    state.error = error.message || "Failed to clear settings.";
  } finally {
    clearBusy();
    render();
  }
}

async function switchSearchType(recordType) {
  state.search.recordType = recordType;
  state.search.query = "";
  state.search.filters = defaultSearchFilters(recordType, state.context || {});
  state.search.results = [];
  state.search.total = 0;
  await ensureLookups(recordType);
  render();
}

async function switchCreateType(recordType) {
  const subject = normalizeSubject(state.mailbox.subject);
  const bodyPreview = trimText(htmlToText(state.mailbox.bodyHtml), 800);
  state.create.recordType = recordType;
  state.create.form = defaultCreateForm(recordType, state.context || {}, subject, bodyPreview);
  state.create.result = null;
  state.partnerPicker.results = [];
  state.partnerPicker.searchTerm = "";
  state.partnerPicker.createName = state.mailbox.senderName || "";
  state.partnerPicker.createEmail = state.mailbox.senderEmail || "";
  state.partnerPicker.createCompany = "";
  await Promise.all([ensureLookups(recordType), ensureSchema(recordType)]);
  render();
}

async function performSearch(form) {
  setBusy(`Searching ${RECORD_LABELS[state.search.recordType].toLowerCase()}s…`);
  state.error = "";
  state.success = "";
  try {
    const payload = collectSearchPayload(form);
    const result = await odooPost(SEARCH_ENDPOINTS[state.search.recordType], payload);
    if (state.search.recordType === "task") {
      state.search.results = result.tasks || [];
    } else if (state.search.recordType === "ticket") {
      state.search.results = result.tickets || [];
    } else {
      state.search.results = result.leads || [];
    }
    state.search.total = result.total || 0;
  } catch (error) {
    state.error = error.message || "Search failed.";
  } finally {
    clearBusy();
    render();
  }
}

function syncDraftsFromDom() {
  const searchForm = document.getElementById("search-form");
  if (searchForm instanceof HTMLFormElement) {
    syncSearchDraft(searchForm);
  }

  const createForm = document.getElementById("create-form");
  if (createForm instanceof HTMLFormElement) {
    syncCreateDraft(createForm);
  }
}

function syncSearchDraft(form) {
  const data = new FormData(form);
  state.search.query = String(data.get("search_term") || "").trim();

  if (state.search.recordType === "task") {
    state.search.filters = {
      project_id: String(data.get("project_id") || ""),
      stage_id: String(data.get("stage_id") || ""),
      user_id: String(data.get("user_id") || ""),
    };
    return;
  }

  if (state.search.recordType === "ticket") {
    state.search.filters = {
      team_id: String(data.get("team_id") || ""),
      stage_id: String(data.get("stage_id") || ""),
      user_id: String(data.get("user_id") || ""),
    };
    return;
  }

  state.search.filters = {
    lead_type: String(data.get("lead_type") || "all"),
    team_id: String(data.get("team_id") || ""),
    stage_id: String(data.get("stage_id") || ""),
    user_id: String(data.get("user_id") || ""),
  };
}

function syncCreateDraft(form) {
  const data = new FormData(form);
  const schema = state.schemas[state.create.recordType] || { extra_fields: [] };
  const nextForm = {
    name: String(data.get("name") || ""),
    description: String(data.get("description") || ""),
    partner_id: String(data.get("partner_id") || ""),
    partner_label: state.create.form.partner_label || "",
    extraValues: {},
  };

  if (state.create.recordType === "task") {
    nextForm.project_id = String(data.get("project_id") || "");
    nextForm.user_id = String(data.get("user_id") || "");
  } else if (state.create.recordType === "ticket") {
    nextForm.team_id = String(data.get("team_id") || "");
    nextForm.priority = String(data.get("priority") || "1");
  } else {
    nextForm.lead_type = String(data.get("lead_type") || "lead");
    nextForm.team_id = String(data.get("team_id") || "");
    nextForm.contact_name = String(data.get("contact_name") || "");
    nextForm.partner_name = String(data.get("partner_name") || "");
    nextForm.email_from = String(data.get("email_from") || "");
  }

  schema.extra_fields.forEach((field) => {
    if (field.type === "boolean") {
      nextForm.extraValues[field.name] = Boolean(form.querySelector(`[name="extra__${field.name}"]`)?.checked);
      return;
    }
    if (field.type === "many2many") {
      const select = form.querySelector(`[name="extra__${field.name}"]`);
      nextForm.extraValues[field.name] = Array.from(select?.selectedOptions || []).map((option) => option.value);
      return;
    }
    nextForm.extraValues[field.name] = data.get(`extra__${field.name}`) || "";
  });

  state.create.form = nextForm;
}

function collectSearchPayload(form) {
  const data = new FormData(form);
  state.search.query = String(data.get("search_term") || "").trim();

  if (state.search.recordType === "task") {
    state.search.filters = {
      project_id: String(data.get("project_id") || ""),
      stage_id: String(data.get("stage_id") || ""),
      user_id: String(data.get("user_id") || ""),
    };
    return {
      search_term: state.search.query,
      project_id: emptyToNull(state.search.filters.project_id),
      stage_id: emptyToNull(state.search.filters.stage_id),
      user_id: emptyToNull(state.search.filters.user_id),
      limit: 20,
      offset: 0,
    };
  }

  if (state.search.recordType === "ticket") {
    state.search.filters = {
      team_id: String(data.get("team_id") || ""),
      stage_id: String(data.get("stage_id") || ""),
      user_id: String(data.get("user_id") || ""),
    };
    return {
      search_term: state.search.query,
      team_id: emptyToNull(state.search.filters.team_id),
      stage_id: emptyToNull(state.search.filters.stage_id),
      user_id: emptyToNull(state.search.filters.user_id),
      limit: 20,
      offset: 0,
    };
  }

  state.search.filters = {
    lead_type: String(data.get("lead_type") || "all"),
    team_id: String(data.get("team_id") || ""),
    stage_id: String(data.get("stage_id") || ""),
    user_id: String(data.get("user_id") || ""),
  };
  return {
    search_term: state.search.query,
    lead_type: state.search.filters.lead_type,
    team_id: emptyToNull(state.search.filters.team_id),
    stage_id: emptyToNull(state.search.filters.stage_id),
    user_id: emptyToNull(state.search.filters.user_id),
    limit: 20,
    offset: 0,
  };
}

async function createRecord(form) {
  setBusy(`Creating ${RECORD_LABELS[state.create.recordType].toLowerCase()}…`);
  state.error = "";
  state.success = "";
  try {
    const payload = collectCreatePayload(form);
    const result = await odooPost(CREATE_ENDPOINTS[state.create.recordType], payload);
    state.create.result = result;
    const url = result.task_url || result.ticket_url || result.lead_url || "";
    const ref = result.task_number || result.ticket_ref || result.lead_ref || "";
    state.success = `${RECORD_LABELS[state.create.recordType]} created${ref ? ` as ${ref}` : ""}.`;
    if (url) {
      openExternal(url);
    }
    await loadHomeData();
  } catch (error) {
    state.error = error.message || "Create failed.";
  } finally {
    clearBusy();
    render();
  }
}

function collectCreatePayload(form) {
  const data = new FormData(form);
  const schema = state.schemas[state.create.recordType] || { extra_fields: [] };
  const extraValues = {};

  schema.extra_fields.forEach((field) => {
    if (field.type === "boolean") {
      extraValues[field.name] = form.querySelector(`[name="extra__${field.name}"]`)?.checked || false;
      return;
    }
    if (field.type === "many2many") {
      const select = form.querySelector(`[name="extra__${field.name}"]`);
      extraValues[field.name] = Array.from(select?.selectedOptions || []).map((option) => option.value);
      return;
    }
    extraValues[field.name] = data.get(`extra__${field.name}`) || "";
  });

  const payload = {
    name: String(data.get("name") || "").trim(),
    description: String(data.get("description") || "").trim(),
    extra_values: extraValues,
    cc_addresses: state.mailbox.ccAddresses || "",
    email_body: state.mailbox.bodyHtml || "",
    email_subject: state.mailbox.subject || "",
    author_email: state.mailbox.senderEmail || "",
    rfc_message_id: state.mailbox.internetMessageId || "",
    outlook_item_id: state.mailbox.itemId || "",
    outlook_conversation_id: state.mailbox.conversationId || "",
  };

  if (state.create.recordType === "task") {
    payload.project_id = String(data.get("project_id") || "");
    payload.user_id = emptyToNull(String(data.get("user_id") || ""));
    payload.partner_id = emptyToNull(String(data.get("partner_id") || ""));
    return payload;
  }

  if (state.create.recordType === "ticket") {
    payload.team_id = String(data.get("team_id") || "");
    payload.priority = String(data.get("priority") || "1");
    payload.partner_id = emptyToNull(String(data.get("partner_id") || ""));
    return payload;
  }

  payload.lead_type = String(data.get("lead_type") || "lead");
  payload.team_id = emptyToNull(String(data.get("team_id") || ""));
  payload.partner_id = emptyToNull(String(data.get("partner_id") || ""));
  payload.contact_name = String(data.get("contact_name") || "").trim();
  payload.partner_name = String(data.get("partner_name") || "").trim();
  payload.email_from = String(data.get("email_from") || "").trim();
  return payload;
}

async function logEmailToRecord(model, recordId) {
  setBusy("Logging email to Odoo…");
  state.error = "";
  state.success = "";
  try {
    await odooPost("/gmail_addon/log_email", {
      res_model: model,
      res_id: Number(recordId),
      email_body: state.mailbox.bodyHtml || "",
      email_subject: state.mailbox.subject || "",
      author_email: state.mailbox.senderEmail || "",
      rfc_message_id: state.mailbox.internetMessageId || "",
      outlook_item_id: state.mailbox.itemId || "",
      outlook_conversation_id: state.mailbox.conversationId || "",
    });
    state.success = "Email logged to Odoo.";
    await loadHomeData();
  } catch (error) {
    state.error = error.message || "Failed to log email.";
  } finally {
    clearBusy();
    render();
  }
}

async function insertReference(reference, target) {
  if (!reference) {
    state.error = "This record does not expose a reference value.";
    render();
    return;
  }

  const item = Office.context.mailbox.item;
  if (!item || !isComposeItem(item)) {
    state.error = "Reference insertion is only available while composing in Outlook.";
    render();
    return;
  }

  setBusy(target === "body" ? "Inserting reference into message…" : "Updating subject…");
  try {
    if (target === "body") {
      await officeAsync(item.body, "prependAsync", `<p><strong>Reference:</strong> ${escapeHtml(reference)}</p>`, {
        coercionType: Office.CoercionType.Html,
      });
      state.success = "Reference inserted into the draft body.";
    } else {
      const currentSubject = await readItemSubject(item);
      const nextSubject = currentSubject.startsWith(`[${reference}]`) ? currentSubject : `[${reference}] ${currentSubject}`.trim();
      await officeAsync(item.subject, "setAsync", nextSubject);
      state.success = "Reference added to the draft subject.";
    }
  } catch (error) {
    state.error = error.message || "Failed to insert the reference.";
  } finally {
    clearBusy();
    render();
  }
}

async function searchPartners() {
  const input = document.getElementById("partner-search-term");
  const term = input instanceof HTMLInputElement ? input.value.trim() : "";
  state.partnerPicker.searchTerm = term;
  state.partnerPicker.results = [];

  if (!term) {
    render();
    return;
  }

  setBusy("Searching contacts…");
  try {
    const result = await odooPost("/gmail_addon/partner/autocomplete", {
      search_term: term,
      limit: 10,
    });
    state.partnerPicker.results = result.partners || [];
    if (!state.partnerPicker.results.length) {
      state.success = "No matching contact found. You can create one below.";
    }
  } catch (error) {
    state.error = error.message || "Partner search failed.";
  } finally {
    clearBusy();
    render();
  }
}

function selectPartner(partnerId, partnerName) {
  state.create.form.partner_id = String(partnerId || "");
  state.create.form.partner_label = partnerName || "";
  state.partnerPicker.results = [];
  state.partnerPicker.searchTerm = "";
  render();
}

function clearPartnerSelection() {
  state.create.form.partner_id = "";
  state.create.form.partner_label = "";
  render();
}

async function createPartner() {
  const nameInput = document.getElementById("partner-create-name");
  const emailInput = document.getElementById("partner-create-email");
  const companyInput = document.getElementById("partner-create-company");
  const name = nameInput instanceof HTMLInputElement ? nameInput.value.trim() : "";
  const email = emailInput instanceof HTMLInputElement ? emailInput.value.trim() : "";
  const companyName = companyInput instanceof HTMLInputElement ? companyInput.value.trim() : "";

  if (!name || !email) {
    state.error = "Partner name and email are required.";
    render();
    return;
  }

  setBusy("Creating contact in Odoo…");
  try {
    const result = await odooPost("/gmail_addon/partner/create", {
      name,
      email,
      company_name: companyName,
    });
    state.create.form.partner_id = stringValue(result.partner_id);
    state.create.form.partner_label = result.partner_name || name;
    state.partnerPicker.results = [];
    state.success = result.already_exists ? "Existing contact selected." : "Partner created in Odoo.";
  } catch (error) {
    state.error = error.message || "Partner creation failed.";
  } finally {
    clearBusy();
    render();
  }
}

function openExternal(url) {
  if (!url) {
    return;
  }
  try {
    if (Office.context.ui && typeof Office.context.ui.openBrowserWindow === "function") {
      Office.context.ui.openBrowserWindow(url);
      return;
    }
  } catch (_) {
    // Fallback below.
  }
  window.open(url, "_blank", "noopener,noreferrer");
}

function render() {
  if (!appEl) {
    return;
  }

  appEl.innerHTML = `
    <div class="layout">
      ${renderHero()}
      ${state.busy ? `<div class="banner banner--info">${escapeHtml(state.busy)}</div>` : ""}
      ${state.error ? `<div class="banner banner--error">${escapeHtml(state.error)}</div>` : ""}
      ${state.success ? `<div class="banner banner--success">${escapeHtml(state.success)}</div>` : ""}
      ${renderMain()}
    </div>
  `;
}

function renderHero() {
  const configured = isConfigured();
  const sender = state.mailbox.senderEmail || "No sender detected";
  return `
    <section class="hero">
      <div class="hero__top">
        <div>
          <p class="hero__eyebrow">Outlook to Odoo</p>
          <h1>Mail workspace for tasks, tickets, and leads</h1>
          <p class="hero__meta">${escapeHtml(sender)}${state.mailbox.subject ? ` · ${escapeHtml(trimText(state.mailbox.subject, 72))}` : ""}</p>
        </div>
        <div class="hero__status">
          <span class="hero__status-dot ${configured && state.connected ? "is-ok" : ""}"></span>
          <span>${configured && state.connected ? "Connected" : "Not connected"}</span>
        </div>
      </div>
      <div class="nav">
        ${renderNavButton("home", "Home")}
        ${renderNavButton("search", "Search")}
        ${renderNavButton("create", "Create")}
        ${renderNavButton("settings", "Settings")}
      </div>
    </section>
  `;
}

function renderNavButton(view, label) {
  return `<button class="nav__button ${state.view === view ? "is-active" : ""}" type="button" data-action="navigate" data-view="${view}">${label}</button>`;
}

function renderMain() {
  if (state.loading && !state.initialized) {
    return `
      <section class="panel">
        <div class="panel__body loading">
          <div class="loading__bar"></div>
          <p>Reading Outlook context and Odoo settings…</p>
        </div>
      </section>
    `;
  }

  if (!isConfigured() && state.view !== "settings") {
    return renderSettings("Connect Odoo first to enable search, create, and log actions.");
  }

  if (state.view === "settings") {
    return renderSettings();
  }
  if (state.view === "search") {
    return renderSearch();
  }
  if (state.view === "create") {
    return renderCreate();
  }
  return renderHome();
}

function renderHome() {
  const context = state.context || {};
  const recentTasks = context.recent_tasks || [];
  const recentTickets = context.recent_tickets || [];
  const recentLeads = context.recent_leads || [];
  const companyTasks = context.company_tasks || [];
  const companyTickets = context.company_tickets || [];
  const companyLeads = context.company_leads || [];
  return `
    <section class="panel">
      <div class="panel__header">
        <h2 class="panel__title">Mailbox context</h2>
        <p class="panel__subtitle">Linked records are recovered from Outlook item or conversation identifiers, then merged with sender-based suggestions from Odoo.</p>
      </div>
      <div class="panel__body stack">
        <div class="summary-grid">
          <div class="summary-card">
            <div class="summary-card__label">Linked</div>
            <div class="summary-card__value">${state.linkedRecords.length}</div>
          </div>
          <div class="summary-card">
            <div class="summary-card__label">Recent tasks</div>
            <div class="summary-card__value">${recentTasks.length}</div>
          </div>
          <div class="summary-card">
            <div class="summary-card__label">Recent leads</div>
            <div class="summary-card__value">${recentLeads.length}</div>
          </div>
        </div>
        <div class="actions">
          <button class="button" type="button" data-action="refresh-home">Refresh context</button>
          <button class="button button--secondary" type="button" data-action="navigate" data-view="search">Open search</button>
          <button class="button button--secondary" type="button" data-action="navigate" data-view="create">Create record</button>
        </div>
        ${context.partner_name ? `
          <div class="banner banner--info">
            Contact match: <strong>${escapeHtml(context.partner_name)}</strong>${context.company_partner_name ? ` at ${escapeHtml(context.company_partner_name)}` : ""}.
          </div>
        ` : `
          <div class="empty-state">No matching contact was inferred from the current sender yet. Search or create will still work.</div>
        `}
        ${renderRecordSection("Linked records", "Records already associated with this Outlook item or conversation.", state.linkedRecords)}
        ${renderRecordSection("Recent tasks", "Most recent project work related to this sender.", recentTasks)}
        ${renderRecordSection("Recent tickets", "Most recent helpdesk tickets for this sender.", recentTickets)}
        ${renderRecordSection("Recent leads", "Most recent leads or opportunities for this sender.", recentLeads)}
        ${context.company_partner_name ? renderRecordSection(`Company tasks · ${context.company_partner_name}`, "Recent work related to the sender's company.", companyTasks) : ""}
        ${context.company_partner_name ? renderRecordSection(`Company tickets · ${context.company_partner_name}`, "Recent support history for the sender's company.", companyTickets) : ""}
        ${context.company_partner_name ? renderRecordSection(`Company leads · ${context.company_partner_name}`, "Recent CRM activity for the sender's company.", companyLeads) : ""}
      </div>
    </section>
  `;
}

function renderSearch() {
  const recordType = state.search.recordType;
  const filters = state.search.filters || defaultSearchFilters(recordType, state.context || {});
  const stageOptions = getStageOptions(recordType, filters);

  return `
    <section class="panel">
      <div class="panel__header">
        <h2 class="panel__title">Search Odoo</h2>
        <p class="panel__subtitle">Use the same task, ticket, and lead endpoints as Gmail, including configurable reference-field search.</p>
      </div>
      <div class="panel__body stack">
        <div class="inline-actions">
          ${renderTypeButtons("select-search-type", recordType)}
        </div>
        <form id="search-form" class="stack">
          <div class="field">
            <label for="search-term">Search term</label>
            <input id="search-term" name="search_term" value="${escapeHtml(state.search.query || "")}" placeholder="Name, reference, email, or contact" />
          </div>
          ${renderSearchFilters(recordType, filters, stageOptions)}
          <div class="actions">
            <button class="button" type="submit">Search ${escapeHtml(RECORD_LABELS[recordType])}s</button>
          </div>
        </form>
        ${state.search.total ? `<p class="section-subtitle">${state.search.total} result(s)</p>` : ""}
        ${renderRecordSection("Results", "Search results open in Odoo and can be logged or inserted into a draft.", state.search.results, true)}
      </div>
    </section>
  `;
}

function renderCreate() {
  const recordType = state.create.recordType;
  const form = state.create.form || defaultCreateForm(recordType, state.context || {}, normalizeSubject(state.mailbox.subject), trimText(htmlToText(state.mailbox.bodyHtml), 800));
  const schema = state.schemas[recordType] || { extra_fields: [] };
  const stageOptions = getStageOptions(recordType, form);

  return `
    <section class="panel">
      <div class="panel__header">
        <h2 class="panel__title">Create in Odoo</h2>
        <p class="panel__subtitle">Standard fields stay explicit; extra fields are loaded dynamically from Odoo settings.</p>
      </div>
      <div class="panel__body stack">
        <div class="inline-actions">
          ${renderTypeButtons("select-create-type", recordType)}
        </div>
        <form id="create-form" class="stack">
          ${renderCreateFields(recordType, form, stageOptions)}
          ${renderPartnerPicker(form)}
          ${schema.extra_fields.length ? `
            <div class="panel">
              <div class="panel__header">
                <h3 class="panel__title">Extra fields</h3>
                <p class="panel__subtitle">Configured in Odoo settings for ${escapeHtml(RECORD_LABELS[recordType].toLowerCase())} creation.</p>
              </div>
              <div class="panel__body stack">
                ${schema.extra_fields.map((field) => renderExtraField(field, form.extraValues || {})).join("")}
              </div>
            </div>
          ` : ""}
          <div class="actions">
            <button class="button" type="submit">Create ${escapeHtml(RECORD_LABELS[recordType])}</button>
          </div>
        </form>
        ${state.create.result ? `
          <div class="banner banner--success">
            ${escapeHtml(RECORD_LABELS[recordType])} created successfully. Use Home to see the linked record immediately.
          </div>
        ` : ""}
      </div>
    </section>
  `;
}

function renderSettings(message) {
  return `
    <section class="panel">
      <div class="panel__header">
        <h2 class="panel__title">Connection settings</h2>
        <p class="panel__subtitle">The add-in uses the same Odoo JSON-RPC endpoints as Gmail and authenticates with an Odoo API key.</p>
      </div>
      <div class="panel__body stack">
        ${message ? `<div class="banner banner--info">${escapeHtml(message)}</div>` : ""}
        <form id="settings-form" class="stack">
          <div class="field">
            <label for="odoo-url">Odoo base URL</label>
            <input id="odoo-url" name="odooUrl" value="${escapeHtml(state.config.odooUrl || "")}" placeholder="https://odoo.example.com" />
          </div>
          <div class="field">
            <label for="odoo-token">API key</label>
            <input id="odoo-token" name="token" type="password" value="${escapeHtml(state.config.token || "")}" placeholder="mail_plugin API key" />
          </div>
          <div class="actions">
            <button class="button" type="submit">Save and connect</button>
            ${isConfigured() ? `<button class="button button--danger" type="button" data-action="clear-settings">Clear</button>` : ""}
          </div>
        </form>
      </div>
    </section>
  `;
}

function renderSearchFilters(recordType, filters, stageOptions) {
  if (recordType === "task") {
    return `
      <div class="grid grid--2">
        ${renderSelectField("project_id", "search-project", "Project", state.lookups.projects, filters.project_id)}
        ${renderSelectField("stage_id", "search-task-stage", "Stage", stageOptions, filters.stage_id)}
        ${renderSelectField("user_id", "search-task-user", "Assignee", state.lookups.users, filters.user_id)}
      </div>
    `;
  }
  if (recordType === "ticket") {
    return `
      <div class="grid grid--2">
        ${renderSelectField("team_id", "search-team", "Team", state.lookups.ticketTeams, filters.team_id)}
        ${renderSelectField("stage_id", "search-ticket-stage", "Stage", stageOptions, filters.stage_id)}
        ${renderSelectField("user_id", "search-ticket-user", "Assignee", state.lookups.users, filters.user_id)}
      </div>
    `;
  }
  return `
    <div class="grid grid--2">
      <div class="field">
        <label for="search-lead-type">Type</label>
        <select id="search-lead-type" name="lead_type">
          <option value="all" ${filters.lead_type === "all" ? "selected" : ""}>Both</option>
          <option value="lead" ${filters.lead_type === "lead" ? "selected" : ""}>Lead</option>
          <option value="opportunity" ${filters.lead_type === "opportunity" ? "selected" : ""}>Opportunity</option>
        </select>
      </div>
      ${renderSelectField("team_id", "search-crm-team", "Sales team", state.lookups.crmTeams, filters.team_id)}
      ${renderSelectField("stage_id", "search-lead-stage", "Stage", stageOptions, filters.stage_id)}
      ${renderSelectField("user_id", "search-lead-user", "Assignee", state.lookups.users, filters.user_id)}
    </div>
  `;
}

function renderCreateFields(recordType, form, stageOptions) {
  const shared = `
    <div class="field">
      <label for="create-name">Name</label>
      <input id="create-name" name="name" value="${escapeHtml(form.name || "")}" required />
    </div>
  `;

  const description = `
    <div class="field">
      <label for="create-description">Description</label>
      <textarea id="create-description" name="description">${escapeHtml(form.description || "")}</textarea>
    </div>
  `;

  if (recordType === "task") {
    return `
      ${shared}
      <div class="grid grid--2">
        ${renderSelectField("project_id", "create-project", "Project", state.lookups.projects, form.project_id, true)}
        ${renderSelectField("user_id", "create-task-user", "Assignee", state.lookups.users, form.user_id)}
      </div>
      ${description}
    `;
  }

  if (recordType === "ticket") {
    return `
      ${shared}
      <div class="grid grid--2">
        ${renderSelectField("team_id", "create-team", "Team", state.lookups.ticketTeams, form.team_id, true)}
        <div class="field">
          <label for="create-priority">Priority</label>
          <select id="create-priority" name="priority">
            <option value="0" ${form.priority === "0" ? "selected" : ""}>Low</option>
            <option value="1" ${form.priority === "1" ? "selected" : ""}>Normal</option>
            <option value="2" ${form.priority === "2" ? "selected" : ""}>High</option>
            <option value="3" ${form.priority === "3" ? "selected" : ""}>Urgent</option>
          </select>
        </div>
      </div>
      ${description}
    `;
  }

  return `
    ${shared}
    <div class="grid grid--2">
      <div class="field">
        <label for="create-lead-type">Type</label>
        <select id="create-lead-type" name="lead_type">
          <option value="lead" ${form.lead_type === "lead" ? "selected" : ""}>Lead</option>
          <option value="opportunity" ${form.lead_type === "opportunity" ? "selected" : ""}>Opportunity</option>
        </select>
      </div>
      ${renderSelectField("team_id", "create-crm-team", "Sales team", state.lookups.crmTeams, form.team_id)}
    </div>
    <div class="grid grid--2">
      <div class="field">
        <label for="create-contact-name">Contact name</label>
        <input id="create-contact-name" name="contact_name" value="${escapeHtml(form.contact_name || "")}" />
      </div>
      <div class="field">
        <label for="create-partner-name">Company</label>
        <input id="create-partner-name" name="partner_name" value="${escapeHtml(form.partner_name || "")}" />
      </div>
      <div class="field">
        <label for="create-email-from">Email</label>
        <input id="create-email-from" name="email_from" value="${escapeHtml(form.email_from || "")}" />
      </div>
    </div>
    ${description}
  `;
}

function renderPartnerPicker(form) {
  return `
    <div class="panel">
      <div class="panel__header">
        <h3 class="panel__title">Contact</h3>
        <p class="panel__subtitle">Keep the inferred partner, search another one, or create a new contact in Odoo.</p>
      </div>
      <div class="panel__body stack">
        <input type="hidden" name="partner_id" value="${escapeHtml(form.partner_id || "")}" />
        ${form.partner_id ? `
          <div class="banner banner--info">
            Selected partner: <strong>${escapeHtml(form.partner_label || "Odoo contact")}</strong>
          </div>
        ` : `
          <div class="empty-state">No partner selected. Tasks and tickets can still be created without one.</div>
        `}
        <div class="grid grid--2">
          <div class="field">
            <label for="partner-search-term">Find existing partner</label>
            <input id="partner-search-term" value="${escapeHtml(state.partnerPicker.searchTerm || "")}" placeholder="Name or email" />
          </div>
          <div class="actions" style="align-items:end;">
            <button class="button button--secondary" type="button" data-action="search-partner">Search</button>
            ${form.partner_id ? `<button class="button button--ghost" type="button" data-action="clear-partner">Clear</button>` : ""}
          </div>
        </div>
        ${state.partnerPicker.results.length ? `
          <div class="record-list">
            ${state.partnerPicker.results.map((partner) => `
              <div class="partner-result">
                <div>
                  <strong>${escapeHtml(partner.name)}</strong>
                  <div class="partner-result__text">${escapeHtml(partner.email || "No email")}</div>
                </div>
                <button class="button button--secondary button--small" type="button" data-action="select-partner" data-partner-id="${partner.id}" data-partner-name="${escapeHtmlAttribute(partner.name)}">Use</button>
              </div>
            `).join("")}
          </div>
        ` : ""}
        <div class="grid grid--2">
          <div class="field">
            <label for="partner-create-name">Create partner name</label>
            <input id="partner-create-name" value="${escapeHtml(state.partnerPicker.createName || "")}" />
          </div>
          <div class="field">
            <label for="partner-create-email">Create partner email</label>
            <input id="partner-create-email" value="${escapeHtml(state.partnerPicker.createEmail || "")}" />
          </div>
          <div class="field">
            <label for="partner-create-company">Company</label>
            <input id="partner-create-company" value="${escapeHtml(state.partnerPicker.createCompany || "")}" />
          </div>
        </div>
        <div class="actions">
          <button class="button button--secondary" type="button" data-action="create-partner">Create partner</button>
        </div>
      </div>
    </div>
  `;
}

function renderExtraField(field, extraValues) {
  const value = extraValues[field.name] ?? "";
  if (field.type === "boolean") {
    return `
      <label class="checkbox">
        <input type="checkbox" name="extra__${field.name}" ${value ? "checked" : ""} />
        <span>${escapeHtml(field.label)}</span>
      </label>
    `;
  }
  if (field.type === "selection" || field.type === "many2one") {
    return renderSelectField(`extra__${field.name}`, `extra__${field.name}`, field.label, field.options || [], value);
  }
  if (field.type === "many2many") {
    const selectedValues = Array.isArray(value) ? value.map(String) : [];
    return `
      <div class="field">
        <label for="extra__${field.name}">${escapeHtml(field.label)}</label>
        <select id="extra__${field.name}" name="extra__${field.name}" multiple size="4">
          ${(field.options || []).map((option) => `
            <option value="${escapeHtml(option.value)}" ${selectedValues.includes(String(option.value)) ? "selected" : ""}>${escapeHtml(option.label)}</option>
          `).join("")}
        </select>
      </div>
    `;
  }
  if (field.type === "text" || field.type === "html") {
    return `
      <div class="field">
        <label for="extra__${field.name}">${escapeHtml(field.label)}</label>
        <textarea id="extra__${field.name}" name="extra__${field.name}">${escapeHtml(String(value || ""))}</textarea>
      </div>
    `;
  }
  const inputType = field.type === "date" ? "date" : field.type === "datetime" ? "datetime-local" : field.type === "integer" || field.type === "float" ? "number" : "text";
  return `
    <div class="field">
      <label for="extra__${field.name}">${escapeHtml(field.label)}</label>
      <input id="extra__${field.name}" name="extra__${field.name}" type="${inputType}" value="${escapeHtml(String(value || ""))}" />
    </div>
  `;
}

function renderSelectField(name, id, label, options, selectedValue, required = false) {
  const normalizedOptions = (options || []).map((option) => ({
    value: String(option.id || option.value || ""),
    label: option.name || option.label || option.value || "",
  }));
  return `
    <div class="field">
      <label for="${id}">${escapeHtml(label)}</label>
      <select id="${id}" name="${name}" ${required ? "required" : ""}>
        <option value="">${required ? "Select…" : "Any"}</option>
        ${normalizedOptions.map((option) => `
          <option value="${escapeHtml(option.value)}" ${String(selectedValue || "") === String(option.value) ? "selected" : ""}>${escapeHtml(option.label)}</option>
        `).join("")}
      </select>
    </div>
  `;
}

function renderTypeButtons(action, activeType) {
  return ["task", "ticket", "lead"].map((recordType) => `
    <button class="button ${recordType === activeType ? "" : "button--secondary"} button--small" type="button" data-action="${action}" data-record-type="${recordType}">
      ${escapeHtml(RECORD_LABELS[recordType])}
    </button>
  `).join("");
}

function renderRecordSection(title, subtitle, records, searchable) {
  return `
    <div class="stack">
      <div>
        <h3 class="section-title">${escapeHtml(title)}</h3>
        <p class="section-subtitle">${escapeHtml(subtitle)}</p>
      </div>
      ${records && records.length ? `
        <div class="record-list">
          ${records.map((record) => renderRecordCard(record, searchable)).join("")}
        </div>
      ` : `
        <div class="empty-state">No records to show.</div>
      `}
    </div>
  `;
}

function renderRecordCard(record) {
  const reference = record.reference || record.task_number || record.ticket_ref || record.lead_ref || "";
  const model = record.type === "task" ? MODEL_BY_TYPE.task : record.type === "ticket" ? MODEL_BY_TYPE.ticket : MODEL_BY_TYPE.lead;
  const meta = [
    reference ? `Reference: ${reference}` : "",
    record.project_name || record.team_name || record.type_label || "",
    record.stage_name || record.stage || "",
    record.user_name || "",
    record.partner_name || "",
    record.email_from || "",
  ].filter(Boolean).join(" · ");
  const badgeLabel = record.type === "lead" && record.type_label ? record.type_label : RECORD_LABELS[record.type] || "Record";
  return `
    <article class="record-card">
      <div class="record-card__top">
        <div>
          <h4 class="record-card__title">${escapeHtml(record.name || "Unnamed record")}</h4>
          <p class="record-card__meta">${escapeHtml(meta || "No additional metadata")}</p>
        </div>
        <span class="badge">${escapeHtml(badgeLabel)}</span>
      </div>
      <div class="record-actions" style="margin-top:12px;">
        ${record.url ? `<button class="button button--secondary button--small" type="button" data-action="open-record" data-url="${escapeHtmlAttribute(record.url)}">Open</button>` : ""}
        <button class="button button--secondary button--small" type="button" data-action="log-record" data-model="${model}" data-id="${record.id}">Log email</button>
        <button class="button button--ghost button--small" type="button" data-action="insert-body" data-reference="${escapeHtmlAttribute(reference)}">Insert body</button>
        <button class="button button--ghost button--small" type="button" data-action="insert-subject" data-reference="${escapeHtmlAttribute(reference)}">Insert subject</button>
      </div>
    </article>
  `;
}

function getStageOptions(recordType, values) {
  if (recordType === "task") {
    return state.lookups.taskStages[(values.project_id || "") || "__all__"] || [];
  }
  if (recordType === "ticket") {
    return state.lookups.ticketStages[(values.team_id || "") || "__all__"] || [];
  }
  return state.lookups.crmStages[(values.team_id || "") || "__all__"] || [];
}

function normalizeSubject(subject) {
  return String(subject || "").replace(/^\s*((re|fw|fwd)\s*:\s*)+/i, "").trim();
}

function htmlToText(html) {
  const doc = new DOMParser().parseFromString(html || "", "text/html");
  return doc.body.textContent || "";
}

function trimText(value, max) {
  if (!value) {
    return "";
  }
  return value.length > max ? `${value.slice(0, max - 1)}…` : value;
}

function stringValue(value) {
  return value ? String(value) : "";
}

function emptyToNull(value) {
  return value || null;
}

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function escapeHtmlAttribute(value) {
  return escapeHtml(value).replace(/`/g, "&#96;");
}
