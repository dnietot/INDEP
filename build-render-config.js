const fs = require("fs");
const path = require("path");

const config = {
  tenant: process.env.ENTRA_TENANT || "bakertilly.co",
  clientId: process.env.ENTRA_CLIENT_ID || "REEMPLAZAR_CLIENT_ID_ENTRA",
  allowedEmailDomain: process.env.ALLOWED_EMAIL_DOMAIN || "@bakertilly.co",
  temporaryLoginEnabled: process.env.TEMP_LOGIN_ENABLED !== "false",
  temporaryLoginName: process.env.TEMP_LOGIN_NAME || "Diego Nieto",
  temporaryLoginEmail: process.env.TEMP_LOGIN_EMAIL || "diego.nieto@bakertilly.co",
  temporaryPasswordHash: process.env.TEMP_LOGIN_PASSWORD_HASH || "8ff2593d80ac7ff8a06a33e35c9ee1ee9d72fb8fd9e9d7c9b57b36d139563543",
  emailWebhookUrl: process.env.EMAIL_WEBHOOK_URL || ""
};

const output = `window.CONFIDENCIALIDAD_CONFIG = ${JSON.stringify(config, null, 2)};\n`;
const outputPath = path.join(__dirname, "prototipo-confidencialidad", "env.js");

fs.writeFileSync(outputPath, output, "utf8");
console.log(`Wrote ${outputPath}`);
