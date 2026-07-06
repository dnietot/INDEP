const fs = require("fs");
const http = require("http");
const os = require("os");
const path = require("path");
const { URL } = require("url");

const port = Number(process.env.PORT || 8766);
const publicDir = path.join(__dirname, "prototipo-confidencialidad");
const assignmentsPath = process.env.ASSIGNMENTS_FILE || path.join(os.tmpdir(), "confidencialidad-assignments.json");
const accessRecordsPath = process.env.ACCESS_RECORDS_FILE || path.join(os.tmpdir(), "confidencialidad-access-records.json");

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

function normalizeAssignments(assignments) {
  if (!assignments || typeof assignments !== "object" || Array.isArray(assignments)) {
    return {};
  }

  return Object.entries(assignments).reduce((accumulator, [key, value]) => {
    const emails = Array.isArray(value)
      ? value.map(normalizeEmail).filter(Boolean)
      : [];

    accumulator[String(key)] = [...new Set(emails)];
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

function normalizeAccessRecord(record) {
  if (!record || typeof record !== "object" || Array.isArray(record)) {
    return null;
  }

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
    requesterEmail: normalizeEmail(record.requesterEmail),
    senderEmail: normalizeEmail(record.senderEmail),
    requestedUsers: Array.isArray(record.requestedUsers)
      ? [...new Set(record.requestedUsers.map(normalizeEmail).filter(Boolean))]
      : [],
    requestedUserEmails: String(record.requestedUserEmails || ""),
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

        sendJson(response, 200, { ok: true, ...result });
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

  serveStatic(request, response, url.pathname);
});

server.listen(port, () => {
  console.log(`Confidencialidad app listening on ${port}`);
});
