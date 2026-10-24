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

// ---- /claude: run `claude -p` (Claude Code CLI) for cells whose model is "claude", "claude:haiku", …
// Tools, MCP servers, hooks/settings and session saving are all off: a cell only ever gets text back.
const { spawn } = require("child_process");
const CLAUDE_BIN = process.env.CLAUDE_BIN || "claude";
const CLAUDE_MAX_PARALLEL = 3; // Excel can fire dozens of cells at once
const CLAUDE_TIMEOUT_MS = 90000;
const CLAUDE_MODELS = { claude: "sonnet", "claude:sonnet": "sonnet", "claude:haiku": "haiku", "claude:opus": "opus" };
let claudeRunning = 0;
const claudeQueue = [];

function runClaude({ model, system, prompt }) {
  return new Promise((resolve) => {
    const start = () => {
      claudeRunning++;
      const args = ["-p", "--tools", "", "--strict-mcp-config", "--no-session-persistence", "--setting-sources", "",
                    "--output-format", "json", "--model", CLAUDE_MODELS[model] || "sonnet"];
      if (system) args.push("--system-prompt", system);
      const child = spawn(CLAUDE_BIN, args, { stdio: ["pipe", "pipe", "pipe"] });
      let out = "", err = "";
      const timer = setTimeout(() => child.kill("SIGKILL"), CLAUDE_TIMEOUT_MS);
      child.stdout.on("data", (c) => (out += c));
      child.stderr.on("data", (c) => (err += c));
      const finish = (result) => {
        clearTimeout(timer);
        claudeRunning--;
        if (claudeQueue.length) claudeQueue.shift()();
        resolve(result);
      };
      child.on("error", () => finish({ status: 503, body: { error: `'${CLAUDE_BIN}' not found — install Claude Code and log in` } }));
      child.on("close", () => {
        try {
          const parsed = JSON.parse(out);
          const events = Array.isArray(parsed) ? parsed : [parsed];
          const r = events.filter((e) => e && e.type === "result").pop();
          if (!r || r.is_error) return finish({ status: 502, body: { error: (r && r.result) || err.slice(0, 200) || "claude returned no result" } });
          finish({ status: 200, body: { response: String(r.result || "").trim(), cost_usd: r.total_cost_usd } });
        } catch {
          finish({ status: 502, body: { error: (err || out).slice(0, 200) || "claude failed (timeout or not logged in?)" } });
        }
      });
      child.stdin.end(prompt);
    };
    claudeRunning < CLAUDE_MAX_PARALLEL ? start() : claudeQueue.push(start);
  });
}

function claudeRoute(req, res) {
  let body = "";
  req.on("data", (c) => (body += c));
  req.on("end", async () => {
    let j;
    try {
      j = JSON.parse(body);
    } catch {
      res.writeHead(400, { "content-type": "application/json" });
      return res.end(JSON.stringify({ error: "bad JSON" }));
    }
    if (LOG_PROMPTS) fs.appendFileSync(LOG_FILE, `\n===== ${new Date().toISOString()} model=${j.model}\n${j.prompt}\n`);
    const { status, body: out } = await runClaude(j);
    res.writeHead(status, { "content-type": "application/json" });
    res.end(JSON.stringify(out));
  });
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
    if (req.url === "/claude" && req.method === "POST") return claudeRoute(req, res);
    serveStatic(req, res);
  })
  .listen(PORT, "127.0.0.1", () => console.log(`Add-in server on https://localhost:${PORT} → Ollama ${OLLAMA.host}:${OLLAMA.port}`));
