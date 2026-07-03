const fs = require("fs");
const path = require("path");

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

const config = {
  tenant: normalizeTenant(process.env.ENTRA_TENANT),
  clientId: normalizeClientId(process.env.ENTRA_CLIENT_ID),
  allowedEmailDomain: process.env.ALLOWED_EMAIL_DOMAIN || "@bakertilly.co",
  temporaryLoginEnabled: process.env.TEMP_LOGIN_ENABLED !== "false",
  temporaryLoginName: process.env.TEMP_LOGIN_NAME || "Diego Nieto",
  temporaryLoginEmail: process.env.TEMP_LOGIN_EMAIL || "diego.nieto@bakertilly.co",
  temporaryPasswordHash: process.env.TEMP_LOGIN_PASSWORD_HASH || "8ff2593d80ac7ff8a06a33e35c9ee1ee9d72fb8fd9e9d7c9b57b36d139563543",
  temporaryAdminEnabled: process.env.TEMP_ADMIN_ENABLED !== "false",
  temporaryAdminLogin: process.env.TEMP_ADMIN_LOGIN || "admin",
  temporaryAdminName: process.env.TEMP_ADMIN_NAME || "Admin",
  temporaryAdminEmail: process.env.TEMP_ADMIN_EMAIL || "admin@bakertilly.co",
  temporaryAdminPasswordHash: process.env.TEMP_ADMIN_PASSWORD_HASH || "8d90ed647b948fa80c3c9bbf5316c78f151723f52fb9d6101f818af8afff69ec",
  emailWebhookUrl: process.env.EMAIL_WEBHOOK_URL || ""
};

const output = `window.CONFIDENCIALIDAD_CONFIG = ${JSON.stringify(config, null, 2)};\n`;
const outputPath = path.join(__dirname, "prototipo-confidencialidad", "env.js");

fs.writeFileSync(outputPath, output, "utf8");
console.log(`Wrote ${outputPath}`);
