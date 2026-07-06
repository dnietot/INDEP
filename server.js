const fs = require("fs");
const http = require("http");
const crypto = require("crypto");
const os = require("os");
const path = require("path");
const tls = require("tls");
const { URL } = require("url");

const port = Number(process.env.PORT || 8766);
const publicDir = path.join(__dirname, "prototipo-confidencialidad");
const assignmentsPath = process.env.ASSIGNMENTS_FILE || path.join(os.tmpdir(), "confidencialidad-assignments.json");
const accessRecordsPath = process.env.ACCESS_RECORDS_FILE || path.join(os.tmpdir(), "confidencialidad-access-records.json");
const appBaseUrl = process.env.APP_BASE_URL || `http://localhost:${port}`;
const accessTeamEmails = (process.env.ACCESS_TEAM_EMAILS || "dnieto@bakertilly.co")
  .split(/[;,]/)
  .map(normalizeEmail)
  .filter(Boolean);

const contentTypes = {
  ".css": "text/css; charset=utf-8",
  ".csv": "text/csv; charset=utf-8",
  ".html": "text/html; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".txt": "text/plain; charset=utf-8"
};

function sendJson(response, statusCode, payload) {
  response.writeHead(statusCode, {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store"
  });
  response.end(JSON.stringify(payload));
}

function readBody(request) {
  return new Promise((resolve, reject) => {
    let body = "";

    request.on("data", (chunk) => {
      body += chunk;
      if (body.length > 1_000_000) {
        reject(new Error("Body too large"));
        request.destroy();
      }
    });
    request.on("end", () => resolve(body));
    request.on("error", reject);
  });
}

function normalizeEmail(value) {
  return String(value || "").trim().toLowerCase();
}

function normalizeEmailList(value) {
  const list = Array.isArray(value)
    ? value
    : String(value || "").split(/[,\n;]/);

  return [...new Set(list.map(normalizeEmail).filter(Boolean))];
}

function normalizeAssignments(assignments) {
  if (!assignments || typeof assignments !== "object" || Array.isArray(assignments)) {
    return {};
  }

  return Object.entries(assignments).reduce((accumulator, [key, value]) => {
    accumulator[String(key)] = normalizeEmailList(value);
    return accumulator;
  }, {});
}

function readAssignments() {
  try {
    return normalizeAssignments(JSON.parse(fs.readFileSync(assignmentsPath, "utf8")));
  } catch (error) {
    return {};
  }
}

function writeAssignments(assignments) {
  fs.mkdirSync(path.dirname(assignmentsPath), { recursive: true });
  fs.writeFileSync(assignmentsPath, JSON.stringify(normalizeAssignments(assignments), null, 2), "utf8");
}

function generateApprovalToken() {
  return crypto.randomBytes(24).toString("hex");
}

function getApprovalStatus(record) {
  return String(record.approvalStatus || record.status || "pending_partner");
}

function normalizeAccessRecord(record) {
  if (!record || typeof record !== "object" || Array.isArray(record)) {
    return null;
  }

  const requesterEmail = normalizeEmail(record.requesterEmail);
  const requestedUsers = Array.isArray(record.requestedUsers)
    ? [...new Set(record.requestedUsers.map(normalizeEmail).filter(Boolean))]
    : [];
  const requestedUserEmails = String(record.requestedUserEmails || requestedUsers.join(", ") || requesterEmail);
  const partnerEmail = normalizeEmail(record.partnerEmail || record.socioEmail);
  const partnerEmails = normalizeEmailList([
    partnerEmail,
    ...normalizeEmailList(record.partnerEmails || record.socioEmails)
  ]);

  return {
    requestId: String(record.requestId || `${Date.now()}-${Math.random().toString(36).slice(2)}`),
    submittedAt: String(record.submittedAt || new Date().toISOString()),
    clientId: String(record.clientId || ""),
    clientName: String(record.clientName || ""),
    nit: String(record.nit || ""),
    huddleName: String(record.huddleName || ""),
    focusName: String(record.focusName || ""),
    serviceLine: String(record.serviceLine || ""),
    manager: String(record.manager || ""),
    requesterName: String(record.requesterName || ""),
    requesterEmail,
    senderEmail: normalizeEmail(record.senderEmail),
    partnerName: String(record.partnerName || record.socio || ""),
    partnerEmail: partnerEmails[0] || partnerEmail,
    partnerEmails,
    approvalStatus: getApprovalStatus(record),
    approvalToken: String(record.approvalToken || generateApprovalToken()),
    approvalRequestedAt: String(record.approvalRequestedAt || ""),
    approvedAt: String(record.approvedAt || ""),
    approvedBy: normalizeEmail(record.approvedBy),
    accessEmailSentAt: String(record.accessEmailSentAt || ""),
    mailError: String(record.mailError || ""),
    requestedUsers: requestedUsers.length > 0 ? requestedUsers : [requesterEmail].filter(Boolean),
    requestedUserEmails,
    accesses: String(record.accesses || ""),
    expiresAt: String(record.expiresAt || ""),
    workToDevelop: String(record.workToDevelop || ""),
    noConflictOfInterest: Boolean(record.noConflictOfInterest),
    authorizedUseConfirmation: Boolean(record.authorizedUseConfirmation),
    recipients: Array.isArray(record.recipients)
      ? record.recipients.map(normalizeEmail).filter(Boolean)
      : []
  };
}

function normalizeAccessRecords(records) {
  if (!Array.isArray(records)) {
    return [];
  }

  return records
    .map(normalizeAccessRecord)
    .filter(Boolean)
    .slice(0, 1000);
}

function readAccessRecords() {
  try {
    return normalizeAccessRecords(JSON.parse(fs.readFileSync(accessRecordsPath, "utf8")));
  } catch (error) {
    return [];
  }
}

function writeAccessRecords(records) {
  fs.mkdirSync(path.dirname(accessRecordsPath), { recursive: true });
  fs.writeFileSync(accessRecordsPath, JSON.stringify(normalizeAccessRecords(records), null, 2), "utf8");
}

function appendAccessRecord(record) {
  const normalizedRecord = normalizeAccessRecord(record);
  if (!normalizedRecord) {
    return null;
  }

  const records = readAccessRecords();
  const mergedRecords = [
    normalizedRecord,
    ...records.filter((existingRecord) => existingRecord.requestId !== normalizedRecord.requestId)
  ].slice(0, 1000);

  writeAccessRecords(mergedRecords);
  return { record: normalizedRecord, records: mergedRecords };
}

function updateAccessRecord(requestId, updater) {
  const records = readAccessRecords();
  const index = records.findIndex((record) => record.requestId === requestId);

  if (index < 0) {
    return null;
  }

  const updatedRecord = normalizeAccessRecord(updater(records[index]));
  records[index] = updatedRecord;
  writeAccessRecords(records);
  return { record: updatedRecord, records };
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function smtpConfigured() {
  return Boolean(process.env.SMTP_HOST && process.env.SMTP_USER && process.env.SMTP_PASS);
}

function smtpResponseReader(socket) {
  let buffer = "";

  return () => new Promise((resolve, reject) => {
    const onData = (chunk) => {
      buffer += chunk;
      const lines = buffer.split(/\r?\n/).filter(Boolean);
      const lastLine = lines[lines.length - 1] || "";

      if (/^\d{3} /.test(lastLine)) {
        socket.off("data", onData);
        socket.off("error", onError);
        const response = buffer;
        buffer = "";
        resolve(response);
      }
    };

    const onError = (error) => {
      socket.off("data", onData);
      socket.off("error", onError);
      reject(error);
    };

    socket.on("data", onData);
    socket.on("error", onError);
  });
}

function assertSmtpOk(response, expectedCodes) {
  const code = Number(String(response).slice(0, 3));
  if (!expectedCodes.includes(code)) {
    throw new Error(`SMTP error ${String(response).trim()}`);
  }
}

function dotStuff(value) {
  return String(value).replace(/^\./gm, "..");
}

function formatAddress(email, name = "") {
  const cleanEmail = normalizeEmail(email);
  return name ? `"${String(name).replaceAll('"', "'")}" <${cleanEmail}>` : cleanEmail;
}

async function sendSmtpMail({ to, subject, text, html }) {
  const recipients = (Array.isArray(to) ? to : [to]).map(normalizeEmail).filter(Boolean);

  if (recipients.length === 0) {
    return { ok: false, skipped: true, error: "No recipients" };
  }

  if (!smtpConfigured()) {
    return { ok: false, skipped: true, error: "SMTP not configured" };
  }

  const host = process.env.SMTP_HOST;
  const portNumber = Number(process.env.SMTP_PORT || 465);
  const user = process.env.SMTP_USER;
  const password = process.env.SMTP_PASS;
  const from = normalizeEmail(process.env.SMTP_FROM || user);
  const timeoutMs = Number(process.env.SMTP_TIMEOUT_MS || 12000);

  return await new Promise((resolve, reject) => {
    const socket = tls.connect({
      host,
      port: portNumber,
      servername: host,
      timeout: timeoutMs
    });
    socket.setEncoding("utf8");

    const readResponse = smtpResponseReader(socket);

    const sendCommand = async (command, expectedCodes = [250]) => {
      socket.write(`${command}\r\n`);
      assertSmtpOk(await readResponse(), expectedCodes);
    };

    socket.on("timeout", () => {
      socket.destroy(new Error("SMTP timeout"));
    });

    socket.once("error", reject);
    socket.once("secureConnect", async () => {
      try {
        assertSmtpOk(await readResponse(), [220]);
        await sendCommand("EHLO indep.onrender.com", [250]);
        await sendCommand(`AUTH PLAIN ${Buffer.from(`\0${user}\0${password}`).toString("base64")}`, [235]);
        await sendCommand(`MAIL FROM:<${from}>`, [250]);

        for (const recipient of recipients) {
          await sendCommand(`RCPT TO:<${recipient}>`, [250, 251]);
        }

        await sendCommand("DATA", [354]);

        const htmlPart = html || `<pre>${escapeHtml(text || "")}</pre>`;
        const textPart = text || String(htmlPart).replace(/<[^>]+>/g, " ");
        const message = [
          `From: ${formatAddress(from, "Confidencialidad")}`,
          `To: ${recipients.join(", ")}`,
          `Subject: ${subject}`,
          "MIME-Version: 1.0",
          "Content-Type: multipart/alternative; boundary=confidencialidad-boundary",
          "",
          "--confidencialidad-boundary",
          "Content-Type: text/plain; charset=utf-8",
          "",
          textPart,
          "--confidencialidad-boundary",
          "Content-Type: text/html; charset=utf-8",
          "",
          htmlPart,
          "--confidencialidad-boundary--",
          ""
        ].join("\r\n");

        socket.write(`${dotStuff(message)}\r\n.\r\n`);
        assertSmtpOk(await readResponse(), [250]);
        await sendCommand("QUIT", [221]);
        socket.end();
        resolve({ ok: true });
      } catch (error) {
        socket.destroy();
        reject(error);
      }
    });
  }).catch((error) => ({ ok: false, error: error.message }));
}

function requestSummaryHtml(record) {
  return `
    <table>
      <tr><td><strong>Cliente</strong></td><td>${escapeHtml(record.clientName)}</td></tr>
      <tr><td><strong>NIT</strong></td><td>${escapeHtml(record.nit)}</td></tr>
      <tr><td><strong>Nombre en Huddle</strong></td><td>${escapeHtml(record.huddleName)}</td></tr>
      <tr><td><strong>Nombre en Focus</strong></td><td>${escapeHtml(record.focusName)}</td></tr>
      <tr><td><strong>Solicitante</strong></td><td>${escapeHtml(record.requesterName)} (${escapeHtml(record.requesterEmail)})</td></tr>
      <tr><td><strong>Accesos solicitados</strong></td><td>${escapeHtml(record.accesses)}</td></tr>
      <tr><td><strong>Vigencia maxima</strong></td><td>${escapeHtml(record.expiresAt)}</td></tr>
      <tr><td><strong>Trabajo a desarrollar</strong></td><td>${escapeHtml(record.workToDevelop)}</td></tr>
      <tr><td><strong>Socio</strong></td><td>${escapeHtml(record.partnerName)} (${escapeHtml((record.partnerEmails || [record.partnerEmail]).filter(Boolean).join(", "))})</td></tr>
    </table>
  `;
}

async function sendPartnerApprovalRequest(record) {
  const partnerRecipients = record.partnerEmails?.length ? record.partnerEmails : [record.partnerEmail].filter(Boolean);

  if (partnerRecipients.length === 0) {
    return updateAccessRecord(record.requestId, (current) => ({
      ...current,
      approvalStatus: "pending_partner",
      mailError: "Cliente sin correo de socio"
    }));
  }

  const result = await sendSmtpMail({
    to: partnerRecipients,
    subject: `[Confidencialidad] Aprobacion requerida - ${record.clientName}`,
    text: `Se registro una solicitud de acceso para ${record.clientName}. Ingresa a ${appBaseUrl} con tu correo para aprobarla.`,
    html: `
      <p>Se registro una solicitud de acceso que requiere aprobacion del socio.</p>
      ${requestSummaryHtml(record)}
      <p>Ingresa a <a href="${escapeHtml(appBaseUrl)}">${escapeHtml(appBaseUrl)}</a> con tu correo para aprobarla desde tu perfil.</p>
    `
  });

  return updateAccessRecord(record.requestId, (current) => ({
    ...current,
    approvalStatus: "pending_partner",
    approvalRequestedAt: result.ok ? new Date().toISOString() : current.approvalRequestedAt,
    mailError: result.ok ? "" : result.error
  }));
}

async function sendAccessRequestToTeam(record) {
  return await sendSmtpMail({
    to: accessTeamEmails,
    subject: `[Confidencialidad] Solicitud aprobada - ${record.clientName} - ${record.requesterEmail}`,
    text: `Solicitud aprobada para ${record.clientName} por ${record.approvedBy || record.partnerEmail}.`,
    html: `
      <p>El socio aprobo la siguiente solicitud de acceso.</p>
      ${requestSummaryHtml(record)}
      <p><strong>Aprobado por:</strong> ${escapeHtml(record.approvedBy || record.partnerEmail)}</p>
      <p><strong>Fecha de aprobacion:</strong> ${escapeHtml(record.approvedAt)}</p>
    `
  });
}

function canApproveRecord(record, approverEmail) {
  const approver = normalizeEmail(approverEmail);
  const partnerEmails = record.partnerEmails?.length ? record.partnerEmails : [record.partnerEmail].filter(Boolean);
  return Boolean(approver && partnerEmails.includes(approver));
}

async function approveAccessRecord(record, approvedBy) {
  if (record.accessEmailSentAt) {
    return { ok: true, alreadyApproved: true, record, records: readAccessRecords(), emailResult: { ok: true } };
  }

  const approvedAt = new Date().toISOString();
  const approved = updateAccessRecord(record.requestId, (current) => ({
    ...current,
    approvalStatus: "approved",
    approvedAt,
    approvedBy: normalizeEmail(approvedBy)
  }));

  if (!approved) {
    return { ok: false, error: "Record not found" };
  }

  const emailResult = await sendAccessRequestToTeam(approved.record);
  const finalStatus = emailResult.ok ? "sent_to_access_team" : "approved_pending_email";
  const finalResult = updateAccessRecord(record.requestId, (current) => ({
    ...current,
    approvalStatus: finalStatus,
    accessEmailSentAt: emailResult.ok ? new Date().toISOString() : current.accessEmailSentAt,
    mailError: emailResult.ok ? "" : emailResult.error
  }));

  return {
    ok: true,
    record: finalResult.record,
    records: finalResult.records,
    emailResult
  };
}

async function handleAssignments(request, response) {
  if (request.method === "GET") {
    sendJson(response, 200, { assignments: readAssignments() });
    return;
  }

  if (request.method === "POST") {
    try {
      const body = JSON.parse(await readBody(request) || "{}");
      const assignments = normalizeAssignments(body.assignments);
      writeAssignments(assignments);
      sendJson(response, 200, { ok: true, assignments });
    } catch (error) {
      sendJson(response, 400, { ok: false, error: "Invalid assignments payload" });
    }
    return;
  }

  sendJson(response, 405, { ok: false, error: "Method not allowed" });
}

async function handleAccessRecords(request, response) {
  if (request.method === "GET") {
    sendJson(response, 200, { records: readAccessRecords() });
    return;
  }

  if (request.method === "POST") {
    try {
      const body = JSON.parse(await readBody(request) || "{}");

      if (body.record) {
        const result = appendAccessRecord(body.record);
        if (!result) {
          throw new Error("Invalid access record");
        }

        const approvalResult = await sendPartnerApprovalRequest(result.record);
        sendJson(response, 200, { ok: true, ...(approvalResult || result) });
        return;
      }

      if (!Array.isArray(body.records)) {
        throw new Error("Invalid access records list");
      }

      const records = normalizeAccessRecords(body.records);
      writeAccessRecords(records);
      sendJson(response, 200, { ok: true, records });
    } catch (error) {
      sendJson(response, 400, { ok: false, error: "Invalid access records payload" });
    }
    return;
  }

  sendJson(response, 405, { ok: false, error: "Method not allowed" });
}

function sendHtml(response, statusCode, body) {
  response.writeHead(statusCode, {
    "Content-Type": "text/html; charset=utf-8",
    "Cache-Control": "no-store"
  });
  response.end(`<!doctype html>
<html lang="es">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Aprobacion de solicitud</title>
    <style>
      body { font-family: Arial, sans-serif; margin: 0; color: #222; background: #f6f7f8; }
      main { max-width: 720px; margin: 8vh auto; background: #fff; padding: 32px; border-radius: 8px; border: 1px solid #dfe3e6; }
      h1 { margin-top: 0; }
      p { line-height: 1.55; }
      .muted { color: #667078; }
    </style>
  </head>
  <body><main>${body}</main></body>
</html>`);
}

async function handleAccessApproval(request, response, url) {
  if (request.method === "POST") {
    try {
      const body = JSON.parse(await readBody(request) || "{}");
      const requestId = String(body.requestId || "");
      const approverEmail = normalizeEmail(body.approverEmail);
      const record = readAccessRecords().find((item) => item.requestId === requestId);

      if (!record) {
        sendJson(response, 404, { ok: false, error: "Solicitud no encontrada" });
        return;
      }

      if (!canApproveRecord(record, approverEmail)) {
        sendJson(response, 403, { ok: false, error: "El correo autenticado no corresponde al socio asignado" });
        return;
      }

      const result = await approveAccessRecord(record, approverEmail);
      sendJson(response, result.ok ? 200 : 400, result);
    } catch (error) {
      sendJson(response, 400, { ok: false, error: "Invalid approval payload" });
    }
    return;
  }

  if (request.method !== "GET") {
    sendJson(response, 405, { ok: false, error: "Method not allowed" });
    return;
  }

  sendHtml(response, 200, `
    <h1>Aprobacion desde la app</h1>
    <p>Para aprobar solicitudes, ingresa a la pagina con tu correo de socio y usa la bandeja de aprobaciones pendientes.</p>
    <p><a href="${escapeHtml(appBaseUrl)}">${escapeHtml(appBaseUrl)}</a></p>
  `);
}

function safeStaticPath(pathname) {
  const decodedPath = decodeURIComponent(pathname === "/" ? "/index.html" : pathname);
  const requestedPath = path.normalize(path.join(publicDir, decodedPath));

  if (requestedPath !== publicDir && !requestedPath.startsWith(publicDir + path.sep)) {
    return null;
  }

  return requestedPath;
}

function serveStatic(request, response, pathname) {
  const filePath = safeStaticPath(pathname);

  if (!filePath || !fs.existsSync(filePath) || fs.statSync(filePath).isDirectory()) {
    response.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
    response.end("Not found");
    return;
  }

  const extension = path.extname(filePath).toLowerCase();
  response.writeHead(200, {
    "Content-Type": contentTypes[extension] || "application/octet-stream",
    "Cache-Control": extension === ".html" ? "no-store" : "public, max-age=60"
  });
  fs.createReadStream(filePath).pipe(response);
}

const server = http.createServer(async (request, response) => {
  const url = new URL(request.url, `http://${request.headers.host || "localhost"}`);

  if (url.pathname === "/api/assignments") {
    await handleAssignments(request, response);
    return;
  }

  if (url.pathname === "/api/access-records") {
    await handleAccessRecords(request, response);
    return;
  }

  if (url.pathname === "/api/access-records/approve") {
    await handleAccessApproval(request, response, url);
    return;
  }

  serveStatic(request, response, url.pathname);
});

server.listen(port, () => {
  console.log(`Confidencialidad app listening on ${port}`);
});
