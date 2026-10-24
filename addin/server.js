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

// Browsers send an Origin header on cross-site requests. Only the add-in's own pages (https://localhost:3000)
// may call the model routes; local tools (curl, the test harness) send no Origin and are allowed.
// Requiring JSON forces browsers into a CORS preflight, which this server never approves.
const SELF_ORIGIN = `https://localhost:${PORT}`;
function rejectForeign(req, res) {
  const origin = req.headers.origin;
  const json = String(req.headers["content-type"] || "").toLowerCase().startsWith("application/json");
  if ((origin && origin !== SELF_ORIGIN) || (req.method === "POST" && !json)) {
    res.writeHead(403, { "content-type": "text/plain" });
    res.end("forbidden: requests must come from the add-in (same origin, JSON)");
    return true;
  }
  return false;
}

function proxy(req, res) {
  if (rejectForeign(req, res)) return;
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
const CLAUDE_MODELS = Object.assign(Object.create(null), { claude: "sonnet", "claude:sonnet": "sonnet", "claude:haiku": "haiku", "claude:opus": "opus" });
const CLAUDE_MAX_QUEUE = 50;
const CLAUDE_MAX_BODY = 2 * 1024 * 1024; // 2 MB: room for a whole-workbook prompt
let claudeRunning = 0;
const claudeQueue = [];

function runClaude({ model, system, prompt }) {
  return new Promise((resolve) => {
    const start = () => {
      claudeRunning++;
      const args = ["-p", "--tools", "", "--strict-mcp-config", "--no-session-persistence", "--setting-sources", "",
                    "--output-format", "json", "--model", CLAUDE_MODELS[model] || "sonnet"];
      if (system) args.push("--system-prompt", system);
      let done = false, timer = null, child, timedOut = false;
      const t0 = Date.now();
      const finish = (result) => {
        if (done) return; // close/error can both fire; release the slot exactly once
        done = true;
        console.log(new Date().toISOString().slice(11, 19), "claude", model, result.status, `${((Date.now() - t0) / 1000).toFixed(1)}s`,
          `${prompt.length} chars`, result.body.error ? `error: ${String(result.body.error).slice(0, 120)}` : "");
        clearTimeout(timer);
        claudeRunning--;
        if (claudeQueue.length) claudeQueue.shift()();
        resolve(result);
      };
      try {
        child = spawn(CLAUDE_BIN, args, { stdio: ["pipe", "pipe", "pipe"] });
      } catch (e) {
        return finish({ status: 503, body: { error: `could not start '${CLAUDE_BIN}'` } });
      }
      let out = "", err = "";
      timer = setTimeout(() => { timedOut = true; child.kill("SIGKILL"); }, CLAUDE_TIMEOUT_MS);
      child.stdout.on("data", (c) => (out += c));
      child.stderr.on("data", (c) => (err += c));
      child.stdin.on("error", () => {}); // EPIPE if the CLI exits early must not crash the server
      child.on("error", () => finish({ status: 503, body: { error: `'${CLAUDE_BIN}' not found — install Claude Code and log in` } }));
      child.on("close", () => {
        if (timedOut) return finish({ status: 504, body: { error: `Claude took longer than ${CLAUDE_TIMEOUT_MS / 1000}s — try again` } });
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
    if (claudeRunning < CLAUDE_MAX_PARALLEL) start();
    else if (claudeQueue.length < CLAUDE_MAX_QUEUE) claudeQueue.push(start);
    else resolve({ status: 429, body: { error: "too many Claude requests queued; try again shortly" } });
  });
}

function claudeRoute(req, res) {
  if (rejectForeign(req, res)) return;
  const reply = (status, obj) => {
    if (res.headersSent) return;
    res.writeHead(status, { "content-type": "application/json" });
    res.end(JSON.stringify(obj));
  };
  let body = "", tooBig = false;
  req.on("data", (c) => {
    if (tooBig) return; // keep reading (and discarding) so the 413 reaches the caller
    body += c;
    if (body.length > CLAUDE_MAX_BODY) {
      tooBig = true;
      body = "";
      reply(413, { error: "request too large (max 2 MB)" });
    }
  });
  req.on("end", async () => {
    if (tooBig) return;
    let j;
    try {
      j = JSON.parse(body);
    } catch {
      return reply(400, { error: "bad JSON" });
    }
    const ok = j && typeof j === "object" && typeof j.prompt === "string" && j.prompt.length > 0 &&
      (j.system === undefined || j.system === null || typeof j.system === "string") &&
      typeof j.model === "string" && Object.prototype.hasOwnProperty.call(CLAUDE_MODELS, j.model);
    if (!ok) return reply(400, { error: `expected {model: ${Object.keys(CLAUDE_MODELS).join("|")}, prompt: string, system?: string}` });
    if (LOG_PROMPTS) fs.appendFileSync(LOG_FILE, `\n===== ${new Date().toISOString()} model=${j.model}\n${j.prompt}\n`);
    try {
      const { status, body: out } = await runClaude({ model: j.model, system: j.system || undefined, prompt: j.prompt });
      reply(status, out);
    } catch (e) {
      reply(500, { error: "claude call failed" });
    }
  });
}

// ---- /log: the add-in's shared-runtime page reports diagnostics here (same-origin only, 64 KB cap)
const RUNTIME_LOG = path.join(__dirname, "logs", "runtime.log");
function logRoute(req, res) {
  if (rejectForeign(req, res)) return;
  let body = "";
  req.on("data", (c) => { if (body.length < 65536) body += c; });
  req.on("end", () => {
    fs.mkdirSync(path.dirname(RUNTIME_LOG), { recursive: true });
    fs.appendFileSync(RUNTIME_LOG, `${new Date().toISOString()} ${body.slice(0, 65536)}\n`);
    res.writeHead(204);
    res.end();
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
    if (req.url === "/log" && req.method === "POST") return logRoute(req, res);
    serveStatic(req, res);
  })
  .listen(PORT, "127.0.0.1", () => console.log(`Add-in server on https://localhost:${PORT} → Ollama ${OLLAMA.host}:${OLLAMA.port}`));
