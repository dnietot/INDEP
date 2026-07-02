const fs = require("fs");
const path = require("path");

const config = {
  tenant: process.env.ENTRA_TENANT || "bakertilly.co",
  clientId: process.env.ENTRA_CLIENT_ID || "REEMPLAZAR_CLIENT_ID_ENTRA",
  allowedEmailDomain: process.env.ALLOWED_EMAIL_DOMAIN || "@bakertilly.co"
};

const output = `window.CONFIDENCIALIDAD_CONFIG = ${JSON.stringify(config, null, 2)};\n`;
const outputPath = path.join(__dirname, "prototipo-confidencialidad", "env.js");

fs.writeFileSync(outputPath, output, "utf8");
console.log(`Wrote ${outputPath}`);
