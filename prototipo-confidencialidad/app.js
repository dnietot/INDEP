let currentUser = {
  name: "Diego Nieto",
  email: "diego.nieto@bakertilly.co",
  role: "user"
};

const defaultTenantId = "100c493f-7265-4dd6-9a05-63e1a210e604";
const defaultClientId = "9d3e6808-f124-4324-875c-7e6da0b0a3bf";
const guidPattern = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

function normalizeTenant(value) {
  const tenant = (value || "").trim();
  return tenant.toLowerCase() === "bakertilly.co" || !tenant ? defaultTenantId : tenant;
}

function normalizeClientId(value) {
  const clientId = (value || "").trim();
  return guidPattern.test(clientId) ? clientId : defaultClientId;
}

const office365Auth = {
  tenant: normalizeTenant(window.CONFIDENCIALIDAD_CONFIG?.tenant),
  clientId: normalizeClientId(window.CONFIDENCIALIDAD_CONFIG?.clientId),
  scopes: ["User.Read"]
};

const allowedEmailDomain = window.CONFIDENCIALIDAD_CONFIG?.allowedEmailDomain || "@bakertilly.co";
const graphMeEndpoint = "https://graph.microsoft.com/v1.0/me?$select=displayName,mail,userPrincipalName,otherMails";
const requestSenderEmail = (window.CONFIDENCIALIDAD_CONFIG?.requestSenderEmail || "accesos@bakertilly.co").toLowerCase();
const clientsCsvUrl = window.CONFIDENCIALIDAD_CONFIG?.clientsCsvUrl || "clientes.csv";
const assignmentsApiUrl = window.CONFIDENCIALIDAD_CONFIG?.assignmentsApiUrl || "/api/assignments";
const accessRecordsApiUrl = window.CONFIDENCIALIDAD_CONFIG?.accessRecordsApiUrl || "/api/access-records";
const approvalApiUrl = `${accessRecordsApiUrl.replace(/\/$/, "")}/approve`;
const showAllClientsWhenUnassigned = window.CONFIDENCIALIDAD_CONFIG?.showAllClientsWhenUnassigned === true;
const temporaryLogin = {
  enabled: window.CONFIDENCIALIDAD_CONFIG?.temporaryLoginEnabled !== false,
  name: window.CONFIDENCIALIDAD_CONFIG?.temporaryLoginName || "Diego Nieto",
  email: (window.CONFIDENCIALIDAD_CONFIG?.temporaryLoginEmail || "diego.nieto@bakertilly.co").toLowerCase(),
  passwordHash: window.CONFIDENCIALIDAD_CONFIG?.temporaryPasswordHash || "8ff2593d80ac7ff8a06a33e35c9ee1ee9d72fb8fd9e9d7c9b57b36d139563543"
};
const temporaryAdmin = {
  enabled: window.CONFIDENCIALIDAD_CONFIG?.temporaryAdminEnabled !== false,
  login: (window.CONFIDENCIALIDAD_CONFIG?.temporaryAdminLogin || "admin").toLowerCase(),
  name: window.CONFIDENCIALIDAD_CONFIG?.temporaryAdminName || "Admin",
  email: (window.CONFIDENCIALIDAD_CONFIG?.temporaryAdminEmail || "admin@bakertilly.co").toLowerCase(),
  passwordHash: window.CONFIDENCIALIDAD_CONFIG?.temporaryAdminPasswordHash || "8d90ed647b948fa80c3c9bbf5316c78f151723f52fb9d6101f818af8afff69ec"
};
const clientsStorageKey = "confidencialidadClients";
const clientAssignmentsStorageKey = "confidencialidadClientAssignments";
const accessRecordsStorageKey = "confidencialidadAccessRecords";

const accessTeam = [
  "accesos@bakertilly.co",
  "seguridad.informacion@bakertilly.co"
];

const defaultClients = [
  {
    id: "CLI-001",
    name: "Andes Retail S.A.S.",
    nit: "900.100.200-1",
    huddleName: "Andes Retail S.A.S.",
    focusName: "Andes Retail",
    serviceLine: "Sin linea",
    manager: "Sin responsable",
    partnerName: "",
    partnerEmail: "",
    partnerEmails: [],
    assignedTo: [],
    confidentialityStatus: "Pendiente"
  },
  {
    id: "CLI-002",
    name: "Constructora Norte Ltda.",
    nit: "830.222.118-7",
    huddleName: "Constructora Norte Ltda.",
    focusName: "Constructora Norte",
    serviceLine: "Sin linea",
    manager: "Sin responsable",
    partnerName: "",
    partnerEmail: "",
    partnerEmails: [],
    assignedTo: [],
    confidentialityStatus: "Vigente"
  },
  {
    id: "CLI-003",
    name: "Servicios Logisticos Delta",
    nit: "901.431.771-4",
    huddleName: "Servicios Logisticos Delta",
    focusName: "Logisticos Delta",
    serviceLine: "Sin linea",
    manager: "Sin responsable",
    partnerName: "",
    partnerEmail: "",
    partnerEmails: [],
    assignedTo: [],
    confidentialityStatus: "Pendiente"
  },
  {
    id: "CLI-004",
    name: "Fondo Aurora",
    nit: "800.765.330-8",
    huddleName: "Fondo Aurora",
    focusName: "Fondo Aurora",
    serviceLine: "Sin linea",
    manager: "Sin responsable",
    partnerName: "",
    partnerEmail: "",
    partnerEmails: [],
    assignedTo: [],
    confidentialityStatus: "Pendiente"
  }
];

let clients = cloneDefaultClients();
let selectedClient = null;
let accessRecords = loadLocalAccessRecords();
let history = buildHistoryFromAccessRecords();
let msalClient = null;
let remoteAssignmentsEnabled = false;
let remoteAccessRecordsEnabled = false;
let accessRecordsRefreshTimer = null;

const authScreen = document.querySelector("#authScreen");
const appShell = document.querySelector("#appShell");
const passwordLoginForm = document.querySelector("#passwordLoginForm");
const passwordEmail = document.querySelector("#passwordEmail");
const passwordInput = document.querySelector("#passwordInput");
const loginButton = document.querySelector("#loginButton");
const logoutButton = document.querySelector("#logoutButton");
const sessionName = document.querySelector("#sessionName");
const sessionMail = document.querySelector("#sessionMail");
const searchInput = document.querySelector("#searchInput");
const clientList = document.querySelector("#clientList");
const surveyForm = document.querySelector("#surveyForm");
const submitSurveyButton = document.querySelector("#submitSurveyButton");
const emptyState = document.querySelector("#emptyState");
const selectedClientName = document.querySelector("#selectedClientName");
const selectedClientMeta = document.querySelector("#selectedClientMeta");
const clearButton = document.querySelector("#clearButton");
const mailPreview = document.querySelector("#mailPreview");
const closePreview = document.querySelector("#closePreview");
const mailTo = document.querySelector("#mailTo");
const mailSubject = document.querySelector("#mailSubject");
const mailBody = document.querySelector("#mailBody");
const historyRows = document.querySelector("#historyRows");
const toast = document.querySelector("#toast");
const metricClientes = document.querySelector("#metricClientes");
const metricPendientes = document.querySelector("#metricPendientes");
const metricEnviadas = document.querySelector("#metricEnviadas");
const adminPanel = document.querySelector("#adminPanel");
const smtpConfigInput = document.querySelector("#smtpConfigInput");
const clientAdminForm = document.querySelector("#clientAdminForm");
const assignmentClientSelect = document.querySelector("#assignmentClientSelect");
const assignmentClientMeta = document.querySelector("#assignmentClientMeta");
const assignmentEmailsInput = document.querySelector("#assignmentEmailsInput");
const exportClientsCsvButton = document.querySelector("#exportClientsCsvButton");
const adminClientsRows = document.querySelector("#adminClientsRows");
const adminAccessRows = document.querySelector("#adminAccessRows");
const partnerApprovalPanel = document.querySelector("#partnerApprovalPanel");
const partnerApprovalRows = document.querySelector("#partnerApprovalRows");

function cloneDefaultClients() {
  return defaultClients.map((client) => ({
    ...client,
    assignedTo: [...client.assignedTo],
    partnerEmails: [...(client.partnerEmails || [])]
  }));
}

function parseCsv(text) {
  const rows = [];
  let row = [];
  let value = "";
  let inQuotes = false;

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    const next = text[index + 1];

    if (char === '"' && inQuotes && next === '"') {
      value += '"';
      index += 1;
    } else if (char === '"') {
      inQuotes = !inQuotes;
    } else if (char === "," && !inQuotes) {
      row.push(value);
      value = "";
    } else if ((char === "\n" || char === "\r") && !inQuotes) {
      if (char === "\r" && next === "\n") {
        index += 1;
      }
      row.push(value);
      if (row.some((cell) => cell.trim())) {
        rows.push(row);
      }
      row = [];
      value = "";
    } else {
      value += char;
    }
  }

  row.push(value);
  if (row.some((cell) => cell.trim())) {
    rows.push(row);
  }

  return rows;
}

function normalizeCsvHeader(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function getCsvValue(row, ...keys) {
  const normalized = keys.map(normalizeCsvHeader);
  const key = Object.keys(row).find((candidate) => normalized.includes(candidate));
  return key ? row[key] : "";
}

function buildClientIdFromNit(nit, index) {
  const cleanNit = String(nit || "").replace(/[^0-9a-z]/gi, "");
  return cleanNit ? `NIT-${cleanNit}` : `CSV-${index + 1}`;
}

function clientsFromCsv(text) {
  const rows = parseCsv(text);
  const headers = rows.shift()?.map(normalizeCsvHeader) || [];

  return rows
    .map((cells, index) => {
      const row = headers.reduce((accumulator, header, cellIndex) => {
        accumulator[header] = String(cells[cellIndex] || "").trim();
        return accumulator;
      }, {});
      const name = getCsvValue(row, "nombre", "cliente", "nombre cliente");
      const nit = getCsvValue(row, "nit", "NIT");
      const huddleName = getCsvValue(row, "nombre en huddle", "huddle", "nombre huddle") || name;
      const focusName = getCsvValue(row, "nombre en focus", "focus", "nombre focus") || name;
      const assignedTo = parseAssignedTo(getCsvValue(row, "correos asignados", "asignados", "assigned to", "assignedTo"));
      const partnerName = getCsvValue(row, "socios asignados", "socio asignado", "socio", "partner", "partner name");
      const partnerEmails = parseAssignedTo(getCsvValue(row, "correo socios", "correo socio", "email socio", "partner email"));
      const partnerEmail = partnerEmails[0] || "";

      return {
        id: buildClientIdFromNit(nit, index),
        name,
        nit,
        huddleName,
        focusName,
        serviceLine: "Sin linea",
        manager: "Sin responsable",
        partnerName,
        partnerEmail,
        partnerEmails,
        assignedTo,
        confidentialityStatus: "Pendiente"
      };
    })
    .filter((client) => client.name && client.nit);
}

function loadLegacyClients() {
  try {
    const saved = localStorage.getItem(clientsStorageKey);
    return saved ? JSON.parse(saved) : [];
  } catch (error) {
    return [];
  }
}

function loadLocalClientAssignments() {
  try {
    const saved = localStorage.getItem(clientAssignmentsStorageKey);
    return saved ? JSON.parse(saved) : {};
  } catch (error) {
    return {};
  }
}

async function loadRemoteClientAssignments() {
  try {
    const response = await fetch(assignmentsApiUrl, { cache: "no-store" });
    if (!response.ok) {
      throw new Error(`Assignments returned ${response.status}`);
    }

    const payload = await response.json();
    remoteAssignmentsEnabled = true;
    return payload.assignments || {};
  } catch (error) {
    remoteAssignmentsEnabled = false;
    return null;
  }
}

async function loadClientAssignments() {
  const remoteAssignments = await loadRemoteClientAssignments();
  return remoteAssignments || loadLocalClientAssignments();
}

function findLegacyClient(client, legacyClients) {
  return legacyClients.find((legacy) => {
    return legacy.id === client.id || legacy.nit === client.nit || normalizeEmail(legacy.name) === normalizeEmail(client.name);
  });
}

async function applySavedClientState(baseClients) {
  const assignments = await loadClientAssignments();
  const legacyClients = loadLegacyClients();

  return baseClients.map((client) => {
    const legacy = findLegacyClient(client, legacyClients);
    const savedAssignments = assignments[client.id] || assignments[client.nit];
    const partnerEmails = uniqueValues(legacy?.partnerEmails || client.partnerEmails || [legacy?.partnerEmail, client.partnerEmail]);

    return {
      ...client,
      assignedTo: uniqueValues(savedAssignments || client.assignedTo || []),
      partnerName: legacy?.partnerName || client.partnerName || "",
      partnerEmails,
      partnerEmail: partnerEmails[0] || "",
      confidentialityStatus: legacy?.confidentialityStatus || client.confidentialityStatus || "Pendiente"
    };
  });
}

async function loadClients() {
  try {
    const response = await fetch(clientsCsvUrl, { cache: "no-store" });
    if (!response.ok) {
      throw new Error(`CSV returned ${response.status}`);
    }

    const csvClients = clientsFromCsv(await response.text());
    return await applySavedClientState(csvClients.length > 0 ? csvClients : cloneDefaultClients());
  } catch (error) {
    return await applySavedClientState(cloneDefaultClients());
  }
}

function buildClientAssignments() {
  return clients.reduce((accumulator, client) => {
    accumulator[client.id] = uniqueValues(client.assignedTo || []);
    accumulator[client.nit] = uniqueValues(client.assignedTo || []);
    return accumulator;
  }, {});
}

async function saveRemoteClientAssignments(assignments) {
  if (!remoteAssignmentsEnabled) {
    return false;
  }

  try {
    const response = await fetch(assignmentsApiUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({ assignments })
    });

    if (!response.ok) {
      throw new Error(`Assignments save returned ${response.status}`);
    }

    return true;
  } catch (error) {
    remoteAssignmentsEnabled = false;
    return false;
  }
}

async function saveClients() {
  const assignments = buildClientAssignments();

  localStorage.setItem(clientAssignmentsStorageKey, JSON.stringify(assignments));
  localStorage.setItem(clientsStorageKey, JSON.stringify(clients));
  return await saveRemoteClientAssignments(assignments);
}

function loadLocalAccessRecords() {
  try {
    const saved = localStorage.getItem(accessRecordsStorageKey);
    return saved ? JSON.parse(saved) : [];
  } catch (error) {
    return [];
  }
}

async function loadRemoteAccessRecords() {
  try {
    const response = await fetch(accessRecordsApiUrl, { cache: "no-store" });
    if (!response.ok) {
      throw new Error(`Access records returned ${response.status}`);
    }

    const payload = await response.json();
    remoteAccessRecordsEnabled = true;
    return Array.isArray(payload.records) ? payload.records : [];
  } catch (error) {
    remoteAccessRecordsEnabled = false;
    return null;
  }
}

async function loadAccessRecords() {
  const remoteRecords = await loadRemoteAccessRecords();
  return remoteRecords || loadLocalAccessRecords();
}

async function saveRemoteAccessRecords(records) {
  if (!remoteAccessRecordsEnabled) {
    return false;
  }

  try {
    const response = await fetch(accessRecordsApiUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({ records })
    });

    if (!response.ok) {
      throw new Error(`Access records save returned ${response.status}`);
    }

    return true;
  } catch (error) {
    remoteAccessRecordsEnabled = false;
    return false;
  }
}

async function appendRemoteAccessRecord(record) {
  if (!remoteAccessRecordsEnabled) {
    return false;
  }

  try {
    const response = await fetch(accessRecordsApiUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({ record })
    });

    if (!response.ok) {
      throw new Error(`Access record append returned ${response.status}`);
    }

    const payload = await response.json();
    const savedRecord = payload.record || record;

    accessRecords = Array.isArray(payload.records)
      ? payload.records
      : [savedRecord, ...accessRecords.filter((existingRecord) => existingRecord.requestId !== savedRecord.requestId)];
    localStorage.setItem(accessRecordsStorageKey, JSON.stringify(accessRecords));
    return true;
  } catch (error) {
    remoteAccessRecordsEnabled = false;
    return false;
  }
}

async function saveAccessRecords() {
  localStorage.setItem(accessRecordsStorageKey, JSON.stringify(accessRecords));
  return await saveRemoteAccessRecords(accessRecords);
}

async function addAccessRecord(record) {
  const savedRemoteRecord = await appendRemoteAccessRecord(record);
  if (savedRemoteRecord) {
    return true;
  }

  accessRecords.unshift(record);
  localStorage.setItem(accessRecordsStorageKey, JSON.stringify(accessRecords));
  return false;
}

async function refreshAccessRecordsFromRemote() {
  if (!canReviewAccessRecords()) return;

  const remoteRecords = await loadRemoteAccessRecords();
  if (!remoteRecords) return;

  accessRecords = remoteRecords;
  localStorage.setItem(accessRecordsStorageKey, JSON.stringify(accessRecords));
  renderHistory();
  renderMetrics();
  renderAdminPanel();
  renderPartnerApprovalPanel();
}

function startAccessRecordsRefresh() {
  window.clearInterval(accessRecordsRefreshTimer);
  if (!canReviewAccessRecords()) return;

  refreshAccessRecordsFromRemote();
  accessRecordsRefreshTimer = window.setInterval(refreshAccessRecordsFromRemote, 30000);
}

function stopAccessRecordsRefresh() {
  window.clearInterval(accessRecordsRefreshTimer);
  accessRecordsRefreshTimer = null;
}

function buildHistoryFromAccessRecords() {
  return accessRecords.map((record) => ({
    date: new Date(record.submittedAt).toLocaleString("es-CO", { dateStyle: "short", timeStyle: "short" }),
    clientName: record.clientName,
    user: record.requesterEmail,
    accesses: record.accesses,
    status: getApprovalStatusLabel(record),
    statusClass: getApprovalStatusClass(record)
  }));
}

function getTemporaryUsers() {
  const users = [];

  if (temporaryLogin.enabled) {
    users.push({
      login: temporaryLogin.email,
      name: temporaryLogin.name,
      email: temporaryLogin.email,
      role: "user",
      passwordHash: temporaryLogin.passwordHash
    });
  }

  if (temporaryAdmin.enabled) {
    users.push({
      login: temporaryAdmin.login,
      name: temporaryAdmin.name,
      email: temporaryAdmin.email,
      role: "admin",
      passwordHash: temporaryAdmin.passwordHash
    });
  }

  return users;
}

function isAdmin() {
  return currentUser.role === "admin";
}

function getApprovalStatus(record) {
  return record?.approvalStatus || "pending_partner";
}

function getApprovalStatusLabel(record) {
  const status = getApprovalStatus(record);

  if (status === "sent_to_access_team") return "Enviado a accesos";
  if (status === "approved_pending_email") return "Aprobado, correo pendiente";
  if (status === "approved") return "Aprobado por socio";
  return "Pendiente socio";
}

function getApprovalStatusClass(record) {
  const status = getApprovalStatus(record);
  return status === "sent_to_access_team" || status === "approved" ? "done" : "pending";
}

function requireAdminAction() {
  if (isAdmin()) {
    return true;
  }

  showToast("Solo el perfil admin puede modificar clientes, asignaciones o configuracion.");
  renderAdminPanel();
  return false;
}

function normalizeEmail(value) {
  return String(value || "").trim().toLowerCase();
}

function parseAssignedTo(value) {
  return String(value || "")
    .split(/[,\n;]/)
    .map(normalizeEmail)
    .filter(Boolean);
}

function parseRequestedUsers(value) {
  return uniqueValues(String(value || "")
    .split(/[,\n;]/)
    .map(normalizeEmail)
    .filter(Boolean));
}

function isBakerEmail(value) {
  const email = normalizeEmail(value);
  return email.includes("@") && email.endsWith(allowedEmailDomain);
}

function getRecordRequestedUsers(record) {
  if (Array.isArray(record.requestedUsers)) {
    return uniqueValues(record.requestedUsers);
  }

  if (Array.isArray(record.requestedUserEmails)) {
    return uniqueValues(record.requestedUserEmails);
  }

  if (typeof record.requestedUserEmails === "string") {
    const users = parseRequestedUsers(record.requestedUserEmails);
    return users.length > 0 ? users : uniqueValues([record.requesterEmail]);
  }

  return uniqueValues([record.requesterEmail]);
}

function formatUsersForTextarea(users) {
  return uniqueValues(users).join(", ");
}

function mergeEmails(currentEmails, newEmails) {
  return uniqueValues([...(currentEmails || []), ...(newEmails || [])]);
}

function uniqueValues(values) {
  const list = Array.isArray(values) ? values : [values];
  return [...new Set(list.map(normalizeEmail).filter(Boolean))];
}

function currentUserEmails() {
  return uniqueValues([currentUser.email, ...(currentUser.aliases || [])]);
}

function matchesAssignedToken(token, userEmail) {
  const assigned = normalizeEmail(token);
  const email = normalizeEmail(userEmail);
  const localPart = email.split("@")[0];

  if (!assigned || !email) return false;
  if (assigned === email || assigned === localPart) return true;
  if (assigned.startsWith("*@") && email.endsWith(assigned.slice(1))) return true;
  if (!assigned.includes("@") && email.endsWith(`@${assigned}`)) return true;

  return false;
}

function isClientAssignedToCurrentUser(client) {
  const userEmails = currentUserEmails();
  return (client.assignedTo || []).some((assigned) => {
    return userEmails.some((email) => matchesAssignedToken(assigned, email));
  });
}

function getClientPartnerEmails(client) {
  return uniqueValues(client.partnerEmails || [client.partnerEmail]);
}

function isPartnerForCurrentUser(client) {
  const userEmails = currentUserEmails();
  const partnerEmails = getClientPartnerEmails(client);
  return partnerEmails.some((partnerEmail) => userEmails.includes(partnerEmail));
}

function partnerClients() {
  return clients.filter(isPartnerForCurrentUser);
}

function canReviewAccessRecords() {
  return isAdmin() || partnerClients().length > 0;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function assignedClients() {
  if (isAdmin()) {
    return clients;
  }

  const assigned = clients.filter((client) => isClientAssignedToCurrentUser(client) || isPartnerForCurrentUser(client));
  return assigned.length > 0 || !showAllClientsWhenUnassigned ? assigned : clients;
}

function renderMetrics() {
  const mine = assignedClients();
  metricClientes.textContent = mine.length;
  metricPendientes.textContent = mine.filter((client) => client.confidentialityStatus === "Pendiente").length;
  metricEnviadas.textContent = isAdmin()
    ? accessRecords.length
    : partnerClients().length > 0
      ? pendingPartnerApprovals().length
      : history.length;
}

function renderClients() {
  const query = searchInput.value.trim().toLowerCase();
  const mine = assignedClients().filter((client) => {
    return `${client.name} ${client.nit} ${client.huddleName} ${client.focusName}`.toLowerCase().includes(query);
  });

  clientList.innerHTML = "";

  if (mine.length === 0) {
    const empty = document.createElement("p");
    empty.className = "muted";
    empty.textContent = query
      ? "No hay clientes para el filtro actual."
      : "No tienes clientes asignados. Un admin debe asignar tu correo a un cliente.";
    clientList.append(empty);
    return;
  }

  mine.forEach((client) => {
    const item = document.createElement("article");
    item.className = `client-item${selectedClient?.id === client.id ? " is-selected" : ""}`;
    const statusClass = client.confidentialityStatus === "Pendiente" ? "pending" : "done";
    const pendingApprovalsCount = isPartnerForCurrentUser(client) ? pendingApprovalsForClient(client.id).length : 0;

    item.innerHTML = `
      <div class="client-top">
        <div>
          <p class="client-name">${escapeHtml(client.name)}</p>
          <p class="client-meta">NIT ${escapeHtml(client.nit)}</p>
          <p class="client-line">Huddle: ${escapeHtml(client.huddleName)}</p>
          <p class="client-line">Focus: ${escapeHtml(client.focusName)}</p>
          ${pendingApprovalsCount > 0 ? `<p class="client-line">Solicitudes pendientes por aprobar: ${pendingApprovalsCount}</p>` : ""}
        </div>
        <span class="badge ${statusClass}">${client.confidentialityStatus}</span>
      </div>
      <div class="client-action">
        <button class="secondary" type="button" data-client-id="${client.id}">Confidencialidad</button>
      </div>
    `;

    clientList.append(item);
  });
}

function selectClient(clientId) {
  selectedClient = clients.find((client) => client.id === clientId);
  if (!selectedClient) return;

  selectedClientName.textContent = selectedClient.name;
  const partnerEmails = getClientPartnerEmails(selectedClient).join(", ") || "sin correo";
  selectedClientMeta.textContent = `NIT ${selectedClient.nit} - Huddle: ${selectedClient.huddleName} - Focus: ${selectedClient.focusName} - Socio: ${selectedClient.partnerName || "Sin socio asignado"} (${partnerEmails})`;
  emptyState.classList.add("is-hidden");
  surveyForm.classList.remove("is-hidden");
  mailPreview.classList.add("is-hidden");
  surveyForm.reset();
  setDefaultDate();
  setDefaultRequestedUsers();
  updateSubmitState();
  renderClients();
}

function setDefaultDate() {
  const date = new Date();
  date.setMonth(date.getMonth() + 3);
  surveyForm.elements.vigencia.value = date.toISOString().slice(0, 10);
}

function setDefaultRequestedUsers() {
}

function getSelectedAccesses(formData = new FormData(surveyForm)) {
  return formData.getAll("tipoAcceso");
}

function updateSubmitState() {
  const formData = new FormData(surveyForm);
  const hasAccess = getSelectedAccesses(formData).length > 0;
  const hasDate = Boolean(formData.get("vigencia"));
  const hasWork = Boolean((formData.get("trabajo") || "").trim());
  const hasNoConflict = formData.get("sinConflicto") === "on";
  const hasConfirmation = formData.get("aceptacion") === "on";

  submitSurveyButton.disabled = !(hasAccess && hasDate && hasWork && hasNoConflict && hasConfirmation);
}

function showToast(message) {
  toast.textContent = message;
  toast.classList.add("is-visible");
  window.clearTimeout(showToast.timer);
  showToast.timer = window.setTimeout(() => toast.classList.remove("is-visible"), 2800);
}

function renderHistory() {
  historyRows.innerHTML = "";
  history = buildHistoryFromAccessRecords()
    .filter((item) => isAdmin() || item.user === currentUser.email);

  if (history.length === 0) {
    historyRows.innerHTML = `
      <tr>
        <td colspan="5" class="muted">Sin envios registrados en esta sesion.</td>
      </tr>
    `;
    return;
  }

  history.forEach((item) => {
    const row = document.createElement("tr");
    row.innerHTML = `
      <td>${escapeHtml(item.date)}</td>
      <td>${escapeHtml(item.clientName)}</td>
      <td>${escapeHtml(item.user)}</td>
      <td>${escapeHtml(item.accesses)}</td>
      <td><span class="badge ${escapeHtml(item.statusClass)}">${escapeHtml(item.status)}</span></td>
    `;
    historyRows.append(row);
  });
}

function buildAccessSummaryRows() {
  return accessRecords.map((record) => ({
    clientName: record.clientName,
    requesterEmail: record.requesterEmail,
    partnerName: record.partnerName || "Sin socio asignado",
    partnerEmail: getRecordPartnerEmails(record).join(", "),
    accesses: record.accesses,
    expiresAt: record.expiresAt,
    workToDevelop: record.workToDevelop,
    status: getApprovalStatusLabel(record),
    statusClass: getApprovalStatusClass(record),
    mailError: record.mailError || ""
  }));
}

function getRecordPartnerEmails(record) {
  return uniqueValues(record.partnerEmails || [record.partnerEmail]);
}

function isPendingPartnerApproval(record) {
  return getApprovalStatus(record) === "pending_partner";
}

function isRecordPendingForCurrentPartner(record) {
  const userEmails = currentUserEmails();
  const partnerEmails = getRecordPartnerEmails(record);
  return isPendingPartnerApproval(record) && partnerEmails.some((email) => userEmails.includes(email));
}

function pendingPartnerApprovals() {
  return accessRecords.filter(isRecordPendingForCurrentPartner);
}

function pendingApprovalsForClient(clientId) {
  return pendingPartnerApprovals().filter((record) => record.clientId === clientId);
}

function renderPartnerApprovalPanel() {
  if (!partnerApprovalPanel || !partnerApprovalRows) return;

  const approvals = pendingPartnerApprovals();
  partnerApprovalPanel.classList.toggle("is-hidden", approvals.length === 0 && partnerClients().length === 0);
  partnerApprovalRows.innerHTML = "";

  if (approvals.length === 0) {
    partnerApprovalRows.innerHTML = `
      <tr>
        <td colspan="7" class="muted">No tienes solicitudes pendientes de aprobacion.</td>
      </tr>
    `;
    return;
  }

  approvals.forEach((record) => {
    const row = document.createElement("tr");
    row.innerHTML = `
      <td>${escapeHtml(record.clientName)}</td>
      <td>${escapeHtml(record.requesterEmail)}</td>
      <td>${escapeHtml(record.accesses)}</td>
      <td>${escapeHtml(record.expiresAt)}</td>
      <td>${escapeHtml(record.workToDevelop)}</td>
      <td><span class="badge pending">${escapeHtml(getApprovalStatusLabel(record))}</span></td>
      <td>
        <button class="primary small-button" type="button" data-approve-request="${escapeHtml(record.requestId)}">Aprobar</button>
      </td>
    `;
    partnerApprovalRows.append(row);
  });
}

async function approvePartnerRequest(requestId) {
  const record = accessRecords.find((item) => item.requestId === requestId);
  if (!record || !isRecordPendingForCurrentPartner(record)) {
    showToast("No tienes permisos para aprobar esta solicitud.");
    return;
  }

  try {
    const response = await fetch(approvalApiUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        requestId,
        approverEmail: currentUser.email
      })
    });
    const payload = await response.json();

    if (!response.ok || !payload.ok) {
      throw new Error(payload.error || `Approval returned ${response.status}`);
    }

    accessRecords = Array.isArray(payload.records) ? payload.records : accessRecords.map((item) => {
      return item.requestId === requestId ? payload.record : item;
    });
    localStorage.setItem(accessRecordsStorageKey, JSON.stringify(accessRecords));
    renderHistory();
    renderMetrics();
    renderAdminPanel();
    renderPartnerApprovalPanel();
    showToast(payload.record?.approvalStatus === "sent_to_access_team"
      ? "Solicitud aprobada y enviada al equipo de accesos."
      : "Solicitud aprobada. El correo final queda pendiente de configuracion SMTP.");
  } catch (error) {
    showToast("No se pudo aprobar la solicitud. Revisa la configuracion o intenta de nuevo.");
  }
}

function updateAssignmentFields() {
  const client = clients.find((item) => item.id === assignmentClientSelect.value);
  if (!client) {
    assignmentClientMeta.textContent = "";
    assignmentEmailsInput.value = "";
    return;
  }

  const assigned = (client.assignedTo || []).join(", ") || "Sin correos asignados";
  const partnerEmails = getClientPartnerEmails(client).join(", ") || "sin correo";
  const partner = client.partnerName
    ? `${client.partnerName} (${partnerEmails})`
    : "Sin socio asignado";
  assignmentClientMeta.textContent = `NIT ${client.nit} - Huddle: ${client.huddleName} - Focus: ${client.focusName} - Socio: ${partner} - Asignados: ${assigned}`;
  assignmentEmailsInput.value = "";
}

function renderAssignmentSelector() {
  const selectedId = assignmentClientSelect.value;
  assignmentClientSelect.innerHTML = "";

  clients.forEach((client) => {
    const option = document.createElement("option");
    option.value = client.id;
    option.textContent = `${client.name} - NIT ${client.nit}`;
    assignmentClientSelect.append(option);
  });

  if (clients.some((client) => client.id === selectedId)) {
    assignmentClientSelect.value = selectedId;
  } else if (clients[0]) {
    assignmentClientSelect.value = clients[0].id;
  }

  updateAssignmentFields();
}

function renderAdminPanel() {
  adminPanel.classList.toggle("is-hidden", !isAdmin());
  if (!isAdmin()) return;

  smtpConfigInput.value = `Remitente previsto: ${requestSenderEmail}`;
  renderAssignmentSelector();
  adminClientsRows.innerHTML = "";
  adminAccessRows.innerHTML = "";

  clients.forEach((client) => {
    const row = document.createElement("tr");
    row.innerHTML = `
      <td>${escapeHtml(client.name)}</td>
      <td>${escapeHtml(client.nit)}</td>
      <td>${escapeHtml(client.huddleName)}</td>
      <td>${escapeHtml(client.focusName)}</td>
      <td>${escapeHtml(client.partnerName || "Sin socio")}</td>
      <td>${escapeHtml(getClientPartnerEmails(client).join(", "))}</td>
      <td>
        <input class="inline-input" type="text" value="${escapeHtml((client.assignedTo || []).join(", "))}" data-assignments-for="${escapeHtml(client.id)}">
      </td>
      <td>
        <div class="row-actions">
          <button class="secondary small-button" type="button" data-save-assignments="${escapeHtml(client.id)}">Guardar</button>
        </div>
      </td>
    `;
    adminClientsRows.append(row);
  });

  const accessRows = buildAccessSummaryRows();

  if (accessRows.length === 0) {
    adminAccessRows.innerHTML = `
      <tr>
        <td colspan="8" class="muted">Aun no hay accesos solicitados.</td>
      </tr>
    `;
    return;
  }

  accessRows.forEach((record) => {
    const row = document.createElement("tr");
    row.innerHTML = `
      <td>${escapeHtml(record.clientName)}</td>
      <td>${escapeHtml(record.requesterEmail)}</td>
      <td>${escapeHtml(record.accesses)}</td>
      <td>${escapeHtml(record.expiresAt)}</td>
      <td>${escapeHtml(record.workToDevelop)}</td>
      <td>${escapeHtml(record.partnerName)}${record.partnerEmail ? `<br><span class="muted">${escapeHtml(record.partnerEmail)}</span>` : ""}</td>
      <td><span class="badge ${escapeHtml(record.statusClass)}">${escapeHtml(record.status)}</span></td>
      <td>${escapeHtml(record.mailError)}</td>
    `;
    adminAccessRows.append(row);
  });
}

function buildAccessRequestPayload(formData) {
  const accesses = getSelectedAccesses(formData).join(", ");
  const requestedUsers = uniqueValues([currentUser.email]);

  return {
    requestId: `${Date.now()}-${Math.random().toString(36).slice(2)}`,
    submittedAt: new Date().toISOString(),
    clientId: selectedClient.id,
    clientName: selectedClient.name,
    nit: selectedClient.nit,
    huddleName: selectedClient.huddleName,
    focusName: selectedClient.focusName,
    serviceLine: selectedClient.serviceLine,
    manager: selectedClient.manager,
    partnerName: selectedClient.partnerName,
    partnerEmail: selectedClient.partnerEmail,
    partnerEmails: getClientPartnerEmails(selectedClient),
    requesterName: currentUser.name,
    requesterEmail: currentUser.email,
    senderEmail: requestSenderEmail,
    approvalStatus: "pending_partner",
    requestedUsers,
    requestedUserEmails: requestedUsers.join(", "),
    accesses,
    expiresAt: formData.get("vigencia"),
    workToDevelop: formData.get("trabajo"),
    noConflictOfInterest: true,
    authorizedUseConfirmation: true,
    recipients: accessTeam
  };
}

function buildEmailPreview(payload) {
  const subject = `[Confidencialidad] Aprobacion requerida - ${payload.clientName} - ${payload.requesterEmail}`;
  const body = [
    `Remitente sugerido: ${payload.senderEmail}`,
    `Cliente: ${payload.clientName}`,
    `NIT: ${payload.nit}`,
    `Nombre en Huddle: ${payload.huddleName}`,
    `Nombre en Focus: ${payload.focusName}`,
    `Solicitante: ${payload.requesterName} (${payload.requesterEmail})`,
    `Socio aprobador: ${payload.partnerName || "Sin socio asignado"} (${(payload.partnerEmails || [payload.partnerEmail]).filter(Boolean).join(", ") || "sin correo"})`,
    `Accesos solicitados: ${payload.accesses}`,
    `Vigencia maxima: ${payload.expiresAt}`,
    `Trabajo a desarrollar: ${payload.workToDevelop}`,
    "Sin conflicto de interes: Si",
    "Confirmacion de uso autorizado: Si"
  ].join(" | ");

  mailTo.textContent = (payload.partnerEmails || [payload.partnerEmail]).filter(Boolean).join("; ") || "Sin correo de socio configurado";
  mailSubject.textContent = subject;
  mailBody.textContent = body;
  mailPreview.classList.remove("is-hidden");
}

async function submitSurvey(event) {
  event.preventDefault();
  if (!selectedClient) return;

  const formData = new FormData(surveyForm);
  const accesses = getSelectedAccesses(formData);

  if (accesses.length === 0 || submitSurveyButton.disabled) {
    showToast("Selecciona los accesos y confirma las declaraciones requeridas.");
    return;
  }

  const payload = buildAccessRequestPayload(formData);
  const now = new Date();
  const savedRemoteRecord = await addAccessRecord(payload);

  history.unshift({
    date: now.toLocaleString("es-CO", { dateStyle: "short", timeStyle: "short" }),
    clientName: selectedClient.name,
    user: currentUser.email,
    accesses: accesses.join(", "),
    status: "Pendiente socio",
    statusClass: "pending"
  });

  selectedClient.confidentialityStatus = "Vigente";
  buildEmailPreview(payload);
  renderHistory();
  renderClients();
  renderMetrics();
  renderAdminPanel();
  renderPartnerApprovalPanel();
  updateSubmitState();

  showToast(savedRemoteRecord
    ? "Solicitud registrada. Queda pendiente de aprobacion del socio."
    : "Solicitud registrada en este navegador. La API del panel admin no respondio.");
}

function getRedirectUri() {
  return `${window.location.origin}${window.location.pathname}`;
}

async function sha256Hex(value) {
  const bytes = new TextEncoder().encode(value);
  const hashBuffer = await window.crypto.subtle.digest("SHA-256", bytes);

  return Array.from(new Uint8Array(hashBuffer))
    .map((byte) => byte.toString(16).padStart(2, "0"))
    .join("");
}

function rememberTemporarySession(user) {
  sessionStorage.setItem("temporaryPasswordSession", "true");
  sessionStorage.setItem("temporaryPasswordLogin", user.login);
}

function forgetTemporarySession() {
  sessionStorage.removeItem("temporaryPasswordSession");
  sessionStorage.removeItem("temporaryPasswordLogin");
}

function applyTemporaryUser(user = getTemporaryUsers()[0]) {
  if (!user) return;

  currentUser = {
    name: user.name,
    email: user.email,
    role: user.role,
    aliases: uniqueValues([user.email, user.login])
  };
}

async function loginWithPassword(event) {
  event.preventDefault();

  const login = passwordEmail.value.trim().toLowerCase();
  const user = getTemporaryUsers().find((candidate) => {
    return candidate.login === login || candidate.email === login;
  });

  if (!user) {
    passwordInput.value = "";
    showToast("Usuario o contrasena incorrectos.");
    return;
  }

  const passwordHash = await sha256Hex(passwordInput.value);

  if (passwordHash !== user.passwordHash) {
    passwordInput.value = "";
    showToast("Usuario o contrasena incorrectos.");
    return;
  }

  applyTemporaryUser(user);
  rememberTemporarySession(user);
  passwordInput.value = "";
  startSession();
}

function hasValidClientIdFormat() {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(office365Auth.clientId);
}

function getClientIdConfigMessage() {
  if (office365Auth.clientId === "REEMPLAZAR_CLIENT_ID_ENTRA") {
    return "Configura el Application (client) ID real de Entra ID antes de iniciar sesion.";
  }

  if (!hasValidClientIdFormat()) {
    return "El Application (client) ID debe ser un GUID real, no texto descriptivo.";
  }

  return "";
}

function isOffice365Configured() {
  return getClientIdConfigMessage() === "";
}

function getMsalClient(options = {}) {
  const quiet = options.quiet === true;
  const configMessage = getClientIdConfigMessage();

  if (configMessage) {
    if (!quiet) {
      showToast(configMessage);
    }
    return null;
  }

  if (typeof msal === "undefined") {
    if (!quiet) {
      showToast("No se pudo cargar MSAL. Revisa la conexion a internet e intenta de nuevo.");
    }
    return null;
  }

  if (!msalClient) {
    msalClient = new msal.PublicClientApplication({
      auth: {
        clientId: office365Auth.clientId,
        authority: `https://login.microsoftonline.com/${office365Auth.tenant}`,
        redirectUri: getRedirectUri(),
        navigateToLoginRequestUrl: false
      },
      cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false
      }
    });
  }

  return msalClient;
}

async function redirectToOffice365() {
  const client = getMsalClient({ quiet: true });
  if (!client) return;

  await client.loginRedirect({
    scopes: office365Auth.scopes,
    prompt: "select_account",
    loginHint: currentUser.email
  });
}

function startSession() {
  sessionName.textContent = currentUser.name;
  sessionMail.textContent = isAdmin() ? `${currentUser.email} - admin` : currentUser.email;
  authScreen.classList.add("is-hidden");
  appShell.classList.remove("is-hidden");
  renderClients();
  renderHistory();
  renderAdminPanel();
  renderPartnerApprovalPanel();
  renderMetrics();
  startAccessRecordsRefresh();
}

function getUserEmailFromGraph(profile, account) {
  return (profile.mail || profile.userPrincipalName || account.username || "").toLowerCase();
}

function applyAuthenticatedUser(profile, account) {
  const email = getUserEmailFromGraph(profile, account);

  if (!email.endsWith(allowedEmailDomain)) {
    showToast(`El correo autenticado debe terminar en ${allowedEmailDomain}.`);
    return false;
  }

  currentUser = {
    name: profile.displayName || account.name || email,
    email,
    role: "user",
    aliases: uniqueValues([
      profile.mail,
      profile.userPrincipalName,
      account.username,
      ...(profile.otherMails || [])
    ])
  };

  return true;
}

async function loadProfileAndStartSession(account) {
  const client = getMsalClient();
  if (!client || !account) return;

  client.setActiveAccount(account);

  try {
    const tokenResponse = await client.acquireTokenSilent({
      account,
      scopes: office365Auth.scopes
    });

    const response = await fetch(graphMeEndpoint, {
      headers: {
        Authorization: `Bearer ${tokenResponse.accessToken}`
      }
    });

    if (!response.ok) {
      throw new Error(`Graph returned ${response.status}`);
    }

    const profile = await response.json();
    if (applyAuthenticatedUser(profile, account)) {
      window.history.replaceState({}, document.title, window.location.pathname);
      startSession();
    }
  } catch (error) {
    if (error instanceof msal.InteractionRequiredAuthError) {
      await client.acquireTokenRedirect({
        account,
        scopes: office365Auth.scopes
      });
      return;
    }

    showToast("No se pudo leer el perfil desde Microsoft Graph.");
  }
}

async function finishOffice365Redirect() {
  const params = new URLSearchParams(window.location.search);

  if (params.get("auth") === "local") {
    window.history.replaceState({}, document.title, window.location.pathname);
    startSession();
    return;
  }

  if (sessionStorage.getItem("temporaryPasswordSession") === "true") {
    const savedLogin = sessionStorage.getItem("temporaryPasswordLogin");
    const savedUser = getTemporaryUsers().find((user) => user.login === savedLogin || user.email === savedLogin);
    applyTemporaryUser(savedUser);
    startSession();
    return;
  }

  const client = getMsalClient();
  if (!client) return;

  try {
    const redirectResponse = await client.handleRedirectPromise();
    const account = redirectResponse?.account || client.getAllAccounts()[0];

    if (account) {
      await loadProfileAndStartSession(account);
    }
  } catch (error) {
    showToast("No se pudo completar el inicio de sesion con Office 365.");
  }
}

async function logout() {
  const client = isOffice365Configured() ? getMsalClient() : null;
  const activeAccount = client?.getActiveAccount();

  selectedClient = null;
  history = [];
  surveyForm.reset();
  forgetTemporarySession();
  stopAccessRecordsRefresh();

  if (client && activeAccount) {
    await client.logoutRedirect({
      account: activeAccount,
      postLogoutRedirectUri: getRedirectUri()
    });
    return;
  }

  authScreen.classList.remove("is-hidden");
  appShell.classList.add("is-hidden");
  adminPanel.classList.add("is-hidden");
  partnerApprovalPanel?.classList.add("is-hidden");
}

async function handleClientAdminSubmit(event) {
  event.preventDefault();
  if (!requireAdminAction()) return;

  const formData = new FormData(clientAdminForm);
  const client = clients.find((item) => item.id === String(formData.get("clientId")));
  const assignedTo = parseAssignedTo(formData.get("assignedTo"));

  if (!client) {
    showToast("Selecciona un cliente.");
    return;
  }

  if (assignedTo.length === 0) {
    showToast("Escribe al menos un correo para agregar.");
    return;
  }

  if (!assignedTo.every(isBakerEmail)) {
    showToast(`Agrega correos validos que terminen en ${allowedEmailDomain}.`);
    return;
  }

  const beforeCount = uniqueValues(client.assignedTo || []).length;
  client.assignedTo = mergeEmails(client.assignedTo, assignedTo);
  const addedCount = client.assignedTo.length - beforeCount;

  const savedRemote = await saveClients();
  assignmentEmailsInput.value = "";
  renderClients();
  renderAdminPanel();
  renderPartnerApprovalPanel();
  renderMetrics();
  showToast(addedCount > 0
    ? savedRemote ? "Correos agregados y guardados globalmente." : "Correos agregados en este navegador. La API global no respondio."
    : "Esos correos ya estaban asignados.");
}

async function removeClient(clientId) {
  if (!requireAdminAction()) return;

  clients = clients.filter((client) => client.id !== clientId);
  accessRecords = accessRecords.filter((record) => record.clientId !== clientId);
  await saveClients();
  await saveAccessRecords();

  if (selectedClient?.id === clientId) {
    selectedClient = null;
    surveyForm.classList.add("is-hidden");
    emptyState.classList.remove("is-hidden");
  }

  renderClients();
  renderHistory();
  renderAdminPanel();
  renderPartnerApprovalPanel();
  renderMetrics();
  showToast("Cliente eliminado.");
}

async function saveClientAssignments(clientId, value) {
  if (!requireAdminAction()) return;

  const client = clients.find((item) => item.id === clientId);
  if (!client) return;

  const assignedTo = parseAssignedTo(value);
  if (assignedTo.length > 0 && !assignedTo.every(isBakerEmail)) {
    showToast(`Agrega correos validos que terminen en ${allowedEmailDomain}.`);
    return;
  }

  const beforeCount = uniqueValues(client.assignedTo || []).length;
  client.assignedTo = mergeEmails(client.assignedTo, assignedTo);
  const addedCount = client.assignedTo.length - beforeCount;

  const savedRemote = await saveClients();
  renderClients();
  renderAdminPanel();
  renderPartnerApprovalPanel();
  renderMetrics();
  showToast(assignedTo.length === 0
    ? "No se agregaron correos."
    : addedCount > 0
      ? savedRemote ? "Asignaciones agregadas y guardadas globalmente." : "Asignaciones agregadas en este navegador. La API global no respondio."
      : "Esos correos ya estaban asignados.");
}

function escapeCsvValue(value) {
  const text = String(value ?? "");
  return /[",\n\r]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
}

function buildClientsCsv() {
  const headers = ["nombre", "NIT", "nombre en huddle", "nombre en focus", "socios asignados", "correo socios", "correos asignados"];
  const rows = clients.map((client) => [
    client.name,
    client.nit,
    client.huddleName,
    client.focusName,
    client.partnerName || "",
    getClientPartnerEmails(client).join("; "),
    (client.assignedTo || []).join("; ")
  ]);

  return [headers, ...rows]
    .map((row) => row.map(escapeCsvValue).join(","))
    .join("\n");
}

function downloadClientsCsv() {
  if (!requireAdminAction()) return;

  const blob = new Blob([buildClientsCsv()], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = url;
  link.download = "clientes.csv";
  document.body.append(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
  showToast("CSV de clientes generado.");
}

passwordEmail.value = temporaryLogin.email;
passwordLoginForm.addEventListener("submit", loginWithPassword);
loginButton.addEventListener("click", redirectToOffice365);
logoutButton.addEventListener("click", logout);
clientAdminForm.addEventListener("submit", handleClientAdminSubmit);
assignmentClientSelect.addEventListener("change", updateAssignmentFields);
exportClientsCsvButton.addEventListener("click", downloadClientsCsv);
searchInput.addEventListener("input", renderClients);
surveyForm.addEventListener("change", updateSubmitState);
surveyForm.addEventListener("input", updateSubmitState);
surveyForm.addEventListener("submit", submitSurvey);
clearButton.addEventListener("click", () => {
  surveyForm.reset();
  setDefaultDate();
  setDefaultRequestedUsers();
  mailPreview.classList.add("is-hidden");
  updateSubmitState();
});
closePreview.addEventListener("click", () => mailPreview.classList.add("is-hidden"));

clientList.addEventListener("click", (event) => {
  const button = event.target.closest("[data-client-id]");
  if (!button) return;
  selectClient(button.dataset.clientId);
});

partnerApprovalRows?.addEventListener("click", (event) => {
  const approveButton = event.target.closest("[data-approve-request]");
  if (!approveButton) return;
  approvePartnerRequest(approveButton.dataset.approveRequest);
});

adminClientsRows.addEventListener("click", (event) => {
  const saveButton = event.target.closest("[data-save-assignments]");
  if (saveButton) {
    const row = saveButton.closest("tr");
    const input = row.querySelector("[data-assignments-for]");
    saveClientAssignments(saveButton.dataset.saveAssignments, input.value);
    return;
  }

  const removeButton = event.target.closest("[data-remove-client]");
  if (!removeButton) return;
  removeClient(removeButton.dataset.removeClient);
});

async function initializeApp() {
  clients = await loadClients();
  accessRecords = await loadAccessRecords();
  history = buildHistoryFromAccessRecords();
  await finishOffice365Redirect();
}

initializeApp();
