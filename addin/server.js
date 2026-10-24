// HTTPS server for the add-in: serves ./public and proxies /api/* to Ollama.
// Same origin for both means no CORS setup and no HTTPS→HTTP mixed-content blocking.
const https = require("https");
const http = require("http");
const fs = require("fs");
const path = require("path");
const os = require("os");

const PORT = 3000;
const OLLAMA = { host: "127.0.0.1", port: 11434 };
const PUBLIC = path.join(__dirname, "public");
const CERTS = process.env.CERTS_DIR || path.join(os.homedir(), ".office-addin-dev-certs");
const TYPES = { ".html": "text/html", ".js": "application/javascript", ".json": "application/json", ".png": "image/png", ".xml": "application/xml" };

const tls = {
  key: fs.readFileSync(path.join(CERTS, "localhost.key")),
  cert: fs.readFileSync(path.join(CERTS, "localhost.crt")),
};

// LOG_PROMPTS=1 appends every prompt Excel sends to logs/prompts.log — to check what the model really sees.
const LOG_PROMPTS = process.env.LOG_PROMPTS === "1";
const LOG_FILE = path.join(__dirname, "logs", "prompts.log");
if (LOG_PROMPTS) fs.mkdirSync(path.dirname(LOG_FILE), { recursive: true });

function proxy(req, res) {
  if (LOG_PROMPTS && req.method === "POST") {
    let body = "";
    req.on("data", (c) => (body += c));
    req.on("end", () => {
      try {
        const j = JSON.parse(body);
        fs.appendFileSync(LOG_FILE, `\n===== ${new Date().toISOString()} model=${j.model}\n${j.prompt}\n`);
      } catch {}
    });
  }
  const upstream = http.request(
    { ...OLLAMA, path: req.url, method: req.method, headers: { "content-type": "application/json" } },
    (up) => {
      res.writeHead(up.statusCode, { "content-type": up.headers["content-type"] || "application/json" });
      up.pipe(res);
    }
  );
  upstream.on("error", () => {
    res.writeHead(502, { "content-type": "text/plain" });
    res.end("Ollama not reachable on localhost:11434");
  });
  req.pipe(upstream);
}

function serveStatic(req, res) {
  const urlPath = decodeURIComponent(req.url.split("?")[0]);
  const file = path.normalize(path.join(PUBLIC, urlPath === "/" ? "taskpane.html" : urlPath));
  if (!file.startsWith(PUBLIC)) {
    res.writeHead(403);
    return res.end();
  }
  fs.readFile(file, (err, buf) => {
    if (err) {
      res.writeHead(404);
      return res.end("not found");
    }
    res.writeHead(200, { "content-type": TYPES[path.extname(file)] || "application/octet-stream", "cache-control": "no-store" });
    res.end(buf);
  });
}

https
  .createServer(tls, (req, res) => {
    console.log(new Date().toISOString().slice(11, 19), req.method, req.url);
    if (req.url.startsWith("/api/")) return proxy(req, res);
    serveStatic(req, res);
  })
  .listen(PORT, "127.0.0.1", () => console.log(`Add-in server on https://localhost:${PORT} → Ollama ${OLLAMA.host}:${OLLAMA.port}`));
