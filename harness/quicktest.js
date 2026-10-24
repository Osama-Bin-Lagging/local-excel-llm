// Quick health check after setup: runs the real add-in code (functions.js) against local Ollama.
// 1 LOKUP.H call + 6 LOKUP.V formulas from the test sheet; each formula is evaluated on its data with the
// add-in's offline engine and compared with the known answer. No Excel needed.
// Usage: node harness/quicktest.js [model]
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const ROOT = path.join(__dirname, "..");
const MODEL = process.argv[2] || undefined; // undefined → the add-in's DEFAULT_MODEL
const EXPECTED = { F01: 355250, F02: 85750, F03: 170, F04: 5, F12: "January", F19: "Rahul Sharma" };

const ctx = {
  fetch, AbortController, console,
  CustomFunctions: {
    associate: (name, f) => (ctx["fn_" + name] = f),
    Error: class extends Error { constructor(code, msg) { super(msg); } },
    ErrorCode: { notAvailable: "N/A", invalidValue: "VALUE" },
  },
};
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.join(ROOT, "addin/public/formula-parser.js"), "utf8"), ctx);
const src = fs.readFileSync(path.join(ROOT, "addin/public/functions.js"), "utf8")
  .replace("https://localhost:3000/api", "http://127.0.0.1:11434/api"); // talk to Ollama directly
vm.runInContext(src + ";this.T={makeEngine,parseAddress,DEFAULT_MODEL}", ctx);
const { makeEngine, parseAddress, DEFAULT_MODEL } = ctx.T;

const same = (a, b) => (typeof b === "number" ? typeof a === "number" && Math.abs(a - b) <= 1e-6 * Math.max(1, Math.abs(b)) : String(a) === String(b));

(async () => {
  const model = MODEL || DEFAULT_MODEL;
  console.log(`model: ${model}`);
  let pass = 0, total = 0;

  // warm-up (loads the model into memory; not timed as a cell)
  let t = Date.now();
  const h = await ctx.fn_H("Reply with only the word OK", null, null, MODEL, { onCanceled: null });
  total++;
  const hok = /\bok\b/i.test(h);
  pass += hok;
  console.log(`${hok ? "✓" : "✗"} LOKUP.H  reply=${JSON.stringify(h.slice(0, 40))}  (${((Date.now() - t) / 1000).toFixed(1)}s incl. model load)`);

  const cases = JSON.parse(fs.readFileSync(path.join(__dirname, "cases.json"), "utf8")).filter((c) => c.id in EXPECTED);
  const times = [];
  for (const c of cases) {
    total++;
    t = Date.now();
    const at = { address: `'Formula Tests'!J${c.row}`, parameterAddresses: [null, c.contextAddress, null], onCanceled: null };
    let f;
    try {
      f = await ctx.fn_V(c.prompt, c.context, MODEL, at);
    } catch (e) {
      f = "#ERROR " + e.message;
    }
    const secs = (Date.now() - t) / 1000;
    times.push(secs);
    const data = parseAddress(c.contextAddress);
    const r = f.startsWith("=") ? makeEngine(data.sheet, data, c.context)(f, parseAddress(at.address)) : { error: "not a formula" };
    const ok = !r.unknown && !r.error && same(r.value, EXPECTED[c.id]);
    pass += ok;
    const got = r.unknown ? "(not evaluable offline)" : r.error ? r.error : JSON.stringify(r.value);
    console.log(`${ok ? "✓" : "✗"} LOKUP.V  ${c.id}  ${secs.toFixed(1)}s  ${f.slice(0, 55).padEnd(55)}  got=${got} want=${JSON.stringify(EXPECTED[c.id])}`);
  }

  times.sort((a, b) => a - b);
  try {
    const ps = await (await fetch("http://127.0.0.1:11434/api/ps")).json();
    for (const m of ps.models || []) console.log(`memory: ${m.name} uses ${(m.size / 1e9).toFixed(2)} GB while loaded`);
  } catch {}
  console.log(`\n${pass}/${total} passed · median ${times[times.length >> 1].toFixed(1)}s per formula`);
  console.log(pass >= total - 1 ? "RESULT: OK (one miss is normal for a 4B model)" : "RESULT: CHECK — more misses than expected; see lines marked ✗");
  process.exit(pass >= total - 1 ? 0 : 1);
})().catch((e) => {
  console.error("quicktest failed:", e.message, "\nIs Ollama running? Try: ./start.sh (or brew services start ollama)");
  process.exit(2);
});
