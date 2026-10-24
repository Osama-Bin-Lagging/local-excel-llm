// Run the Formula Tests through the real add-in code (functions.js) against one or more Ollama models.
// Usage: node run.js cases.json model1 [model2 ...]   → writes results-<model>.json
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const [casesFile, ...models] = process.argv.slice(2);
const cases = JSON.parse(fs.readFileSync(casesFile, "utf8"));

// Load functions.js with a fake CustomFunctions global; talk to Ollama directly instead of via server.js.
// Spec = model plus switches, e.g. "qwen3:4b-instruct+plan+dyn".
//   +think   reasoning on (bigger output budget)      +plan    one "Plan:" line before the formula
//   +dyn     3 matched examples, full reference          +compact matched examples, short reference
//   +nocheck no checker / auto sheet prefix / retry   +nogloss no definitions   +notarget no target-cell text
//   +nopp    presence_penalty 0 (override model default)
//   +noeval  don't run formulas in JS                   +novote  no extra samples  +voteall sample for every cell
function loadAddin(flags) {
  const think = flags.includes("think");
  let src = fs
    .readFileSync(path.join(__dirname, "../addin/public/functions.js"), "utf8")
    .replace("https://localhost:3000/api", "http://127.0.0.1:11434/api");
  if (think) src = src.replace("think: false", "think: true").replace("num_predict: 256", "num_predict: 4096");
  // +nopp: override a model's built-in presence penalty (qwen3.5 ships 1.5, which punishes repeating sheet names/ranges)
  if (flags.includes("nopp")) src = src.replace("num_predict: 256,", "num_predict: 256, presence_penalty: 0,");
  const opts = {
    prompt: flags.includes("compact") ? "compact" : flags.includes("dyn") ? "dyn" : "full",
    plan: flags.includes("plan"),
    check: !flags.includes("nocheck"),
    gloss: !flags.includes("nogloss"),
    target: !flags.includes("notarget"),
    evaluate: !flags.includes("noeval"),
    vote: flags.includes("novote") ? "off" : flags.includes("voteall") ? "always" : "fail",
  };
  const before = src;
  src = src.replace(/const FORMULA_OPTS = \{[^}]*\};/, `const FORMULA_OPTS = ${JSON.stringify(opts)};`);
  if (src === before) throw new Error("FORMULA_OPTS line not found in functions.js");
  const engineSrc = fs.readFileSync(path.join(__dirname, "../addin/public/formula-parser.js"), "utf8");
  const fns = {};
  const ctx = {
    fetch, AbortController, console,
    CustomFunctions: {
      associate: (name, f) => (fns[name] = f),
      Error: class extends Error { constructor(code, msg) { super(msg); } },
      ErrorCode: { notAvailable: "N/A", invalidValue: "VALUE" },
    },
  };
  vm.createContext(ctx);
  vm.runInContext(engineSrc, ctx); // same offline formula engine the add-in loads before functions.js
  vm.runInContext(src, ctx);
  return fns;
}

(async () => {
  for (const spec of models) {
    const [model, ...flags] = spec.split("+");
    const fns = loadAddin(flags); // fresh cache per model
    // warm-up so load time isn't counted against the first test
    await (fns.H || fns.ASK)("Reply with OK", null, null, model, { onCanceled: null });
    const results = [];
    for (const c of cases) {
      const invocation = {
        address: c.address || `'Formula Tests'!J${c.row}`, // real-test tasks carry their own target cell
        parameterAddresses: [null, c.contextAddress, null],
      };
      const t = Date.now();
      let out;
      try {
        // LOKUP.H / LOKUP.V (formerly LLAMA.ASK / LLAMA.FORMULA) — accept either name
        const askFn = fns.H || fns.ASK, formulaFn = fns.V || fns.FORMULA;
        out = c.fn === "ask"
          ? await askFn(c.prompt, c.context, null, model, invocation)
          : await formulaFn(c.prompt, c.context, model, invocation);
      } catch (e) {
        out = `#ERROR ${e.message}`;
      }
      const secs = (Date.now() - t) / 1000;
      results.push({ id: c.id, row: c.row, out, secs });
      process.stdout.write(`${spec} ${c.id} ${secs.toFixed(1)}s  ${out.replace(/\n/g, " ").slice(0, 90)}\n`);
    }
    fs.writeFileSync(path.join(__dirname, `results-${spec.replace(/[:/+]/g, "_")}.json`), JSON.stringify(results, null, 1));
  }
})();
