// Builds the ready-to-run Windows package in windows/addin so end users only
// need Node.js (no npm install). Run with: npm run package:windows
const { execSync } = require("child_process");
const fs = require("fs");
const path = require("path");

const root = path.join(__dirname, "..");
execSync("npx webpack --mode production", {
  cwd: root,
  stdio: "inherit",
  env: { ...process.env, ADDIN_URL: "https://localhost:3000/" },
});

const out = path.join(root, "windows", "addin");
fs.rmSync(out, { recursive: true, force: true });
fs.cpSync(path.join(root, "dist"), out, {
  recursive: true,
  filter: (src) => !src.endsWith(".map") && !src.endsWith(".LICENSE.txt"),
});
console.log(`Windows package written to ${out}`);
