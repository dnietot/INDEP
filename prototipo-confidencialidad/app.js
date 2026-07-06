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
const emailWebhookUrl = window.CONFIDENCIALIDAD_CONFIG?.emailWebhookUrl || "";
const showAllClientsWhenUnassigned = window.CONFIDENCIALIDAD_CONFIG?.showAllClientsWhenUnassigned !== false;
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
const accessRecordsStorageKey = "confidencialidadAccessRecords";
const webhookStorageKey = "confidencialidadEmailWebhookUrl";

const accessTeam = [
  "accesos@bakertilly.co",
  "seguridad.informacion@bakertilly.co"
];

const defaultClients = [
  {
    id: "CLI-001",
    name: "Andes Retail S.A.S.",
    nit: "900.100.200-1",
    serviceLine: "Auditoria",
    manager: "Laura Mendoza",
    assignedTo: ["diego.nieto@bakertilly.co"],
    confidentialityStatus: "Pendiente"
  },
  {
    id: "CLI-002",
    name: "Constructora Norte Ltda.",
    nit: "830.222.118-7",
    serviceLine: "Impuestos",
    manager: "Felipe Rojas",
    assignedTo: ["diego.nieto@bakertilly.co"],
    confidentialityStatus: "Vigente"
  },
  {
    id: "CLI-003",
    name: "Servicios Logisticos Delta",
    nit: "901.431.771-4",
    serviceLine: "BPO",
    manager: "Mariana Cruz",
    assignedTo: ["diego.nieto@bakertilly.co"],
    confidentialityStatus: "Pendiente"
  },
  {
    id: "CLI-004",
    name: "Fondo Aurora",
    nit: "800.765.330-8",
    serviceLine: "Consultoria",
    manager: "Camilo Vargas",
    assignedTo: ["otro.usuario@bakertilly.co"],
    confidentialityStatus: "Pendiente"
  }
];

let clients = loadClients();
let selectedClient = null;
let accessRecords = loadAccessRecords();
let history = buildHistoryFromAccessRecords();
let msalClient = null;

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
const emailConfigForm = document.querySelector("#emailConfigForm");
const webhookUrlInput = document.querySelector("#webhookUrlInput");
const clearWebhookButton = document.querySelector("#clearWebhookButton");
const clientAdminForm = document.querySelector("#clientAdminForm");
const adminClientsRows = document.querySelector("#adminClientsRows");
const adminAccessRows = document.querySelector("#adminAccessRows");

function cloneDefaultClients() {
  return defaultClients.map((client) => ({
    ...client,
    assignedTo: [...client.assignedTo]
  }));
}

function loadClients() {
  try {
    const saved = localStorage.getItem(clientsStorageKey);
    return saved ? JSON.parse(saved) : cloneDefaultClients();
  } catch (error) {
    return cloneDefaultClients();
  }
}

function saveClients() {
  localStorage.setItem(clientsStorageKey, JSON.stringify(clients));
}

function loadAccessRecords() {
  try {
    const saved = localStorage.getItem(accessRecordsStorageKey);
    return saved ? JSON.parse(saved) : [];
  } catch (error) {
    return [];
  }
}

function saveAccessRecords() {
  localStorage.setItem(accessRecordsStorageKey, JSON.stringify(accessRecords));
}

function buildHistoryFromAccessRecords() {
  return accessRecords.map((record) => ({
    date: new Date(record.submittedAt).toLocaleString("es-CO", { dateStyle: "short", timeStyle: "short" }),
    clientName: record.clientName,
    user: record.requesterEmail,
    requestedUsers: getRecordRequestedUsers(record).join(", "),
    accesses: record.accesses
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

function getEffectiveEmailWebhookUrl() {
  return localStorage.getItem(webhookStorageKey) || emailWebhookUrl;
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

  const assigned = clients.filter(isClientAssignedToCurrentUser);
  return assigned.length > 0 || !showAllClientsWhenUnassigned ? assigned : clients;
}

function renderMetrics() {
  const mine = assignedClients();
  metricClientes.textContent = mine.length;
  metricPendientes.textContent = mine.filter((client) => client.confidentialityStatus === "Pendiente").length;
  metricEnviadas.textContent = isAdmin() ? accessRecords.length : history.length;
}

function renderClients() {
  const query = searchInput.value.trim().toLowerCase();
  const mine = assignedClients().filter((client) => {
    return `${client.name} ${client.nit} ${client.serviceLine}`.toLowerCase().includes(query);
  });

  clientList.innerHTML = "";

  if (mine.length === 0) {
    const empty = document.createElement("p");
    empty.className = "muted";
    empty.textContent = "No hay clientes para el filtro actual.";
    clientList.append(empty);
    return;
  }

  mine.forEach((client) => {
    const item = document.createElement("article");
    item.className = `client-item${selectedClient?.id === client.id ? " is-selected" : ""}`;
    const statusClass = client.confidentialityStatus === "Pendiente" ? "pending" : "done";

    item.innerHTML = `
      <div class="client-top">
        <div>
          <p class="client-name">${client.name}</p>
          <p class="client-meta">NIT ${client.nit}</p>
          <p class="client-line">${client.serviceLine} - ${client.manager}</p>
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
  selectedClientMeta.textContent = `${selectedClient.id} - NIT ${selectedClient.nit} - ${selectedClient.serviceLine}`;
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
  surveyForm.elements.usuariosAcceso.value = formatUsersForTextarea([currentUser.email]);
}

function getSelectedAccesses(formData = new FormData(surveyForm)) {
  return formData.getAll("tipoAcceso");
}

function updateSubmitState() {
  const formData = new FormData(surveyForm);
  const hasAccess = getSelectedAccesses(formData).length > 0;
  const requestedUsers = parseRequestedUsers(formData.get("usuariosAcceso"));
  const hasValidRequestedUsers = requestedUsers.length > 0 && requestedUsers.every(isBakerEmail);
  const hasDate = Boolean(formData.get("vigencia"));
  const hasWork = Boolean((formData.get("trabajo") || "").trim());
  const hasNoConflict = formData.get("sinConflicto") === "on";
  const hasConfirmation = formData.get("aceptacion") === "on";

  submitSurveyButton.disabled = !(hasAccess && hasValidRequestedUsers && hasDate && hasWork && hasNoConflict && hasConfirmation);
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
        <td colspan="6" class="muted">Sin envios registrados en esta sesion.</td>
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
      <td>${escapeHtml(item.requestedUsers)}</td>
      <td>${escapeHtml(item.accesses)}</td>
      <td><span class="badge done">Correo preparado</span></td>
    `;
    historyRows.append(row);
  });
}

function buildAccessSummaryRows() {
  return accessRecords.flatMap((record) => {
    return getRecordRequestedUsers(record).map((requestedUser) => ({
      clientName: record.clientName,
      requestedUser,
      accesses: record.accesses,
      expiresAt: record.expiresAt,
      workToDevelop: record.workToDevelop,
      requesterEmail: record.requesterEmail
    }));
  });
}

function renderAdminPanel() {
  adminPanel.classList.toggle("is-hidden", !isAdmin());
  if (!isAdmin()) return;

  webhookUrlInput.value = getEffectiveEmailWebhookUrl();
  adminClientsRows.innerHTML = "";
  adminAccessRows.innerHTML = "";

  clients.forEach((client) => {
    const row = document.createElement("tr");
    row.innerHTML = `
      <td>${escapeHtml(client.name)}</td>
      <td>${escapeHtml(client.nit)}</td>
      <td>${escapeHtml(client.serviceLine)}</td>
      <td>
        <input class="inline-input" type="text" value="${escapeHtml((client.assignedTo || []).join(", "))}" data-assignments-for="${escapeHtml(client.id)}">
      </td>
      <td>
        <div class="row-actions">
          <button class="secondary small-button" type="button" data-save-assignments="${escapeHtml(client.id)}">Guardar</button>
          <button class="secondary small-button" type="button" data-remove-client="${escapeHtml(client.id)}">Quitar</button>
        </div>
      </td>
    `;
    adminClientsRows.append(row);
  });

  const accessRows = buildAccessSummaryRows();

  if (accessRows.length === 0) {
    adminAccessRows.innerHTML = `
      <tr>
        <td colspan="6" class="muted">Aun no hay accesos solicitados.</td>
      </tr>
    `;
    return;
  }

  accessRows.forEach((record) => {
    const row = document.createElement("tr");
    row.innerHTML = `
      <td>${escapeHtml(record.clientName)}</td>
      <td>${escapeHtml(record.requestedUser)}</td>
      <td>${escapeHtml(record.accesses)}</td>
      <td>${escapeHtml(record.expiresAt)}</td>
      <td>${escapeHtml(record.workToDevelop)}</td>
      <td>${escapeHtml(record.requesterEmail)}</td>
    `;
    adminAccessRows.append(row);
  });
}

function buildAccessRequestPayload(formData) {
  const accesses = getSelectedAccesses(formData).join(", ");
  const requestedUsers = parseRequestedUsers(formData.get("usuariosAcceso"));

  return {
    requestId: `${Date.now()}-${Math.random().toString(36).slice(2)}`,
    submittedAt: new Date().toISOString(),
    clientId: selectedClient.id,
    clientName: selectedClient.name,
    nit: selectedClient.nit,
    serviceLine: selectedClient.serviceLine,
    manager: selectedClient.manager,
    requesterName: currentUser.name,
    requesterEmail: currentUser.email,
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
  const subject = `[Confidencialidad] Solicitud de acceso - ${payload.clientName} - ${payload.requesterEmail}`;
  const body = [
    `Cliente: ${payload.clientName}`,
    `NIT: ${payload.nit}`,
    `Solicitante: ${payload.requesterName} (${payload.requesterEmail})`,
    `Usuarios que requieren acceso: ${payload.requestedUserEmails}`,
    `Accesos solicitados: ${payload.accesses}`,
    `Vigencia maxima: ${payload.expiresAt}`,
    `Trabajo a desarrollar: ${payload.workToDevelop}`,
    "Sin conflicto de interes: Si",
    "Confirmacion de uso autorizado: Si"
  ].join(" | ");

  mailTo.textContent = accessTeam.join("; ");
  mailSubject.textContent = subject;
  mailBody.textContent = body;
  mailPreview.classList.remove("is-hidden");
}

async function sendEmailWebhook(payload) {
  const webhookUrl = getEffectiveEmailWebhookUrl();

  if (!webhookUrl) {
    return false;
  }

  await fetch(webhookUrl, {
    method: "POST",
    mode: "no-cors",
    headers: {
      "Content-Type": "text/plain;charset=utf-8"
    },
    body: JSON.stringify(payload)
  });

  return true;
}

async function submitSurvey(event) {
  event.preventDefault();
  if (!selectedClient) return;

  const formData = new FormData(surveyForm);
  const accesses = getSelectedAccesses(formData);
  const requestedUsers = parseRequestedUsers(formData.get("usuariosAcceso"));

  if (requestedUsers.length === 0 || !requestedUsers.every(isBakerEmail)) {
    showToast(`Agrega correos validos que terminen en ${allowedEmailDomain}.`);
    return;
  }

  if (accesses.length === 0 || submitSurveyButton.disabled) {
    showToast("Selecciona los accesos, usuarios y confirma las declaraciones requeridas.");
    return;
  }

  const payload = buildAccessRequestPayload(formData);
  const now = new Date();
  accessRecords.unshift(payload);
  saveAccessRecords();

  history.unshift({
    date: now.toLocaleString("es-CO", { dateStyle: "short", timeStyle: "short" }),
    clientName: selectedClient.name,
    user: currentUser.email,
    accesses: accesses.join(", ")
  });

  selectedClient.confidentialityStatus = "Vigente";
  buildEmailPreview(payload);
  renderHistory();
  renderClients();
  renderMetrics();
  renderAdminPanel();
  updateSubmitState();

  try {
    const sentToWebhook = await sendEmailWebhook(payload);
    showToast(sentToWebhook
      ? "Solicitud registrada y enviada al flujo de correo."
      : "Solicitud registrada. Falta configurar EMAIL_WEBHOOK_URL para enviar el correo.");
  } catch (error) {
    showToast("Solicitud registrada, pero no se pudo llamar el flujo de correo.");
  }
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
  renderMetrics();
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
}

function handleEmailConfigSubmit(event) {
  event.preventDefault();
  const url = webhookUrlInput.value.trim();

  if (url) {
    localStorage.setItem(webhookStorageKey, url);
    showToast("URL del flujo guardada en este navegador.");
  } else {
    localStorage.removeItem(webhookStorageKey);
    showToast("URL del flujo eliminada.");
  }

  renderAdminPanel();
}

function clearWebhookConfig() {
  webhookUrlInput.value = "";
  localStorage.removeItem(webhookStorageKey);
  showToast("URL del flujo eliminada.");
  renderAdminPanel();
}

function handleClientAdminSubmit(event) {
  event.preventDefault();
  const formData = new FormData(clientAdminForm);
  const assignedTo = parseAssignedTo(formData.get("assignedTo"));

  const client = {
    id: `CLI-${Date.now()}`,
    name: String(formData.get("name") || "").trim(),
    nit: String(formData.get("nit") || "").trim(),
    serviceLine: String(formData.get("serviceLine") || "").trim(),
    manager: String(formData.get("manager") || "").trim(),
    assignedTo,
    confidentialityStatus: "Pendiente"
  };

  if (!client.name || !client.nit || !client.serviceLine || !client.manager || assignedTo.length === 0) {
    showToast("Completa todos los campos del cliente.");
    return;
  }

  clients.push(client);
  saveClients();
  clientAdminForm.reset();
  renderClients();
  renderAdminPanel();
  renderMetrics();
  showToast("Cliente agregado.");
}

function removeClient(clientId) {
  clients = clients.filter((client) => client.id !== clientId);
  accessRecords = accessRecords.filter((record) => record.clientId !== clientId);
  saveClients();
  saveAccessRecords();

  if (selectedClient?.id === clientId) {
    selectedClient = null;
    surveyForm.classList.add("is-hidden");
    emptyState.classList.remove("is-hidden");
  }

  renderClients();
  renderHistory();
  renderAdminPanel();
  renderMetrics();
  showToast("Cliente eliminado.");
}

function saveClientAssignments(clientId, value) {
  const client = clients.find((item) => item.id === clientId);
  if (!client) return;

  const assignedTo = parseAssignedTo(value);
  if (assignedTo.length === 0) {
    showToast("Agrega al menos un correo asignado.");
    return;
  }

  client.assignedTo = assignedTo;
  saveClients();
  renderClients();
  renderAdminPanel();
  renderMetrics();
  showToast("Asignaciones guardadas.");
}

passwordEmail.value = temporaryLogin.email;
passwordLoginForm.addEventListener("submit", loginWithPassword);
loginButton.addEventListener("click", redirectToOffice365);
logoutButton.addEventListener("click", logout);
emailConfigForm.addEventListener("submit", handleEmailConfigSubmit);
clearWebhookButton.addEventListener("click", clearWebhookConfig);
clientAdminForm.addEventListener("submit", handleClientAdminSubmit);
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

finishOffice365Redirect();
