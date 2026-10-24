// How often does the offline JS engine agree with real Excel? Runs every answer-key formula through it.
const fs = require("fs"), vm = require("vm"), path = require("path");
const ctx = { console, CustomFunctions: { associate() {}, Error: class {}, ErrorCode: {} } };
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.join(__dirname, "../addin/public/formula-parser.js"), "utf8"), ctx);
vm.runInContext(fs.readFileSync(path.join(__dirname, "../addin/public/functions.js"), "utf8") + ";this.T={makeEngine,parseAddress}", ctx);
const { makeEngine, parseAddress } = ctx.T;
const tasks = JSON.parse(fs.readFileSync(process.argv[2]));
let agree = 0, unknown = 0, wrongErr = 0, wrongVal = 0;
for (const t of tasks) {
  const data = parseAddress(t.contextAddress.replace(/^'([^']+)'!/, "$1!").replace(/^([^!]+)!/, (m, s) => `'${s}'!`));
  const at = parseAddress(t.address);
  const eng = makeEngine(data.sheet, data, t.context);
  const r = eng(t.key_formula || t.ref_formula, at);
  const ref = t.ref;
  const same = (a, b) => (typeof a === "number" && typeof b === "number" ? Math.abs(a - b) <= 1e-6 * Math.max(1, Math.abs(b)) : String(a) === String(b));
  if (r.unknown) unknown++;
  else if (r.error) { wrongErr++; console.log("  FALSE ERROR", t.id, r.error, t.key_formula); }
  else if (same(r.value, ref)) agree++;
  else { wrongVal++; console.log("  DIFFERENT  ", t.id, JSON.stringify(r.value), "vs", ref, t.key_formula); }
}
console.log(`agree ${agree}  could-not-evaluate ${unknown}  false-error ${wrongErr}  different-value ${wrongVal}  (of ${tasks.length})`);
