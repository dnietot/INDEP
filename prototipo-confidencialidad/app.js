let currentUser = {
  name: "Diego Nieto",
  email: "diego.nieto@bakertilly.co",
  role: "Senior"
};

const office365Auth = {
  tenant: window.CONFIDENCIALIDAD_CONFIG?.tenant || "bakertilly.co",
  clientId: window.CONFIDENCIALIDAD_CONFIG?.clientId || "REEMPLAZAR_CLIENT_ID_ENTRA",
  scopes: ["User.Read"]
};

const allowedEmailDomain = window.CONFIDENCIALIDAD_CONFIG?.allowedEmailDomain || "@bakertilly.co";
const graphMeEndpoint = "https://graph.microsoft.com/v1.0/me";
const temporaryLogin = {
  enabled: window.CONFIDENCIALIDAD_CONFIG?.temporaryLoginEnabled !== false,
  name: window.CONFIDENCIALIDAD_CONFIG?.temporaryLoginName || "Diego Nieto",
  email: (window.CONFIDENCIALIDAD_CONFIG?.temporaryLoginEmail || "diego.nieto@bakertilly.co").toLowerCase(),
  passwordHash: window.CONFIDENCIALIDAD_CONFIG?.temporaryPasswordHash || "8ff2593d80ac7ff8a06a33e35c9ee1ee9d72fb8fd9e9d7c9b57b36d139563543"
};

const accessTeam = [
  "accesos@bakertilly.co",
  "seguridad.informacion@bakertilly.co"
];

const clients = [
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

let selectedClient = null;
let history = [];
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

function assignedClients() {
  return clients.filter((client) => client.assignedTo.includes(currentUser.email));
}

function renderMetrics() {
  const mine = assignedClients();
  metricClientes.textContent = mine.length;
  metricPendientes.textContent = mine.filter((client) => client.confidentialityStatus === "Pendiente").length;
  metricEnviadas.textContent = history.length;
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
  updateSubmitState();
  renderClients();
}

function setDefaultDate() {
  const date = new Date();
  date.setMonth(date.getMonth() + 3);
  surveyForm.elements.vigencia.value = date.toISOString().slice(0, 10);
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
      <td>${item.date}</td>
      <td>${item.clientName}</td>
      <td>${item.user}</td>
      <td>${item.accesses}</td>
      <td><span class="badge done">Correo preparado</span></td>
    `;
    historyRows.append(row);
  });
}

function buildEmailPreview(formData) {
  const accesses = getSelectedAccesses(formData).join(", ");
  const subject = `[Confidencialidad] Solicitud de acceso - ${selectedClient.name} - ${currentUser.email}`;
  const body = [
    `Cliente: ${selectedClient.name}`,
    `NIT: ${selectedClient.nit}`,
    `Solicitante: ${currentUser.name} (${currentUser.email})`,
    `Accesos solicitados: ${accesses}`,
    `Vigencia maxima: ${formData.get("vigencia")}`,
    `Trabajo a desarrollar: ${formData.get("trabajo")}`,
    "Sin conflicto de interes: Si",
    "Confirmacion de uso autorizado: Si"
  ].join(" | ");

  mailTo.textContent = accessTeam.join("; ");
  mailSubject.textContent = subject;
  mailBody.textContent = body;
  mailPreview.classList.remove("is-hidden");
}

function submitSurvey(event) {
  event.preventDefault();
  if (!selectedClient) return;

  const formData = new FormData(surveyForm);
  const accesses = getSelectedAccesses(formData);

  if (accesses.length === 0 || submitSurveyButton.disabled) {
    showToast("Selecciona los accesos y confirma las declaraciones requeridas.");
    return;
  }

  const now = new Date();
  history.unshift({
    date: now.toLocaleString("es-CO", { dateStyle: "short", timeStyle: "short" }),
    clientName: selectedClient.name,
    user: currentUser.email,
    accesses: accesses.join(", ")
  });

  selectedClient.confidentialityStatus = "Vigente";
  buildEmailPreview(formData);
  renderHistory();
  renderClients();
  renderMetrics();
  updateSubmitState();
  showToast("Encuesta registrada y correo preparado para accesos.");
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

function rememberTemporarySession() {
  sessionStorage.setItem("temporaryPasswordSession", "true");
}

function forgetTemporarySession() {
  sessionStorage.removeItem("temporaryPasswordSession");
}

function applyTemporaryUser() {
  currentUser = {
    name: temporaryLogin.name,
    email: temporaryLogin.email,
    role: currentUser.role
  };
}

async function loginWithPassword(event) {
  event.preventDefault();

  if (!temporaryLogin.enabled) {
    showToast("El ingreso temporal con contrasena no esta habilitado.");
    return;
  }

  const email = passwordEmail.value.trim().toLowerCase();
  const passwordHash = await sha256Hex(passwordInput.value);

  if (email !== temporaryLogin.email || passwordHash !== temporaryLogin.passwordHash) {
    passwordInput.value = "";
    showToast("Correo o contrasena incorrectos.");
    return;
  }

  applyTemporaryUser();
  rememberTemporarySession();
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
  sessionMail.textContent = currentUser.email;
  authScreen.classList.add("is-hidden");
  appShell.classList.remove("is-hidden");
  renderClients();
  renderHistory();
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
    role: currentUser.role
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
    applyTemporaryUser();
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
}

passwordEmail.value = temporaryLogin.email;
passwordLoginForm.addEventListener("submit", loginWithPassword);
loginButton.addEventListener("click", redirectToOffice365);
logoutButton.addEventListener("click", logout);
searchInput.addEventListener("input", renderClients);
surveyForm.addEventListener("change", updateSubmitState);
surveyForm.addEventListener("input", updateSubmitState);
surveyForm.addEventListener("submit", submitSurvey);
clearButton.addEventListener("click", () => {
  surveyForm.reset();
  setDefaultDate();
  mailPreview.classList.add("is-hidden");
  updateSubmitState();
});
closePreview.addEventListener("click", () => mailPreview.classList.add("is-hidden"));

clientList.addEventListener("click", (event) => {
  const button = event.target.closest("[data-client-id]");
  if (!button) return;
  selectClient(button.dataset.clientId);
});

finishOffice365Redirect();
