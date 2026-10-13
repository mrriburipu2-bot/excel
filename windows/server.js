// Zero-dependency HTTPS server for the prebuilt add-in (windows/addin).
// Usage: node server.js <path-to-localhost.pfx> <pfx-password>
"use strict";
const https = require("https");
const fs = require("fs");
const path = require("path");

const PORT = 3000;
const ROOT = path.join(__dirname, "addin");
const TYPES = {
  ".html": "text/html; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".xml": "application/xml; charset=utf-8",
  ".png": "image/png",
};

const [pfxPath, passphrase] = process.argv.slice(2);
let pfx;
try {
  pfx = fs.readFileSync(pfxPath);
} catch {
  console.error(`Certificate not found at ${pfxPath}. Run install-windows.cmd again.`);
  process.exit(1);
}

const server = https.createServer({ pfx, passphrase }, (req, res) => {
  const urlPath = decodeURIComponent(new URL(req.url, "https://localhost").pathname);
  const file = path.normalize(path.join(ROOT, urlPath === "/" ? "taskpane.html" : urlPath));
  // Only serve files inside the add-in folder.
  if (!file.startsWith(ROOT + path.sep)) {
    res.writeHead(403).end();
    return;
  }
  fs.readFile(file, (err, data) => {
    if (err) {
      res.writeHead(404).end("Not found");
      return;
    }
    res.writeHead(200, {
      "Content-Type": TYPES[path.extname(file).toLowerCase()] || "application/octet-stream",
      "Access-Control-Allow-Origin": "*",
      "Cache-Control": "no-cache",
    });
    res.end(data);
  });
});

server.on("error", (err) => {
  if (err.code === "EADDRINUSE") {
    console.error(`Port ${PORT} is already in use. The add-in server may already be running in another window.`);
  } else {
    console.error(`Server error: ${err.message}`);
  }
  process.exit(1);
});

server.listen(PORT, "127.0.0.1", () => {
  console.log("");
  console.log(`Claude AI add-in is running at https://localhost:${PORT}/`);
  console.log("Keep this window open while you use the add-in in Excel. Close it to stop.");
});
