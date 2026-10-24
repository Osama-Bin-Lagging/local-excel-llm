/* LOOOKUP — whole-workbook functions (shared runtime: custom functions can call the Excel API).
   =LOOOKUP.H(prompt, [model]) — answers in text using every sheet. Default model: Claude Sonnet. Reuses generate() etc. from functions.js. */

const WORKBOOK_DEFAULT_MODEL = "claude"; // LOOOKUP → Claude Sonnet (online); LOKUP stays on the local DEFAULT_MODEL
const WORKBOOK_MAX_CHARS = 50000; // ≈ 34k tokens: cell lists are digit-heavy (measured ~1.5 chars/token with qwen3)
const WORKBOOK_ROW_STEPS = [50, 25, 12, 6]; // too big → show only this many top rows per sheet, smallest last

// Read every sheet's used range: name, address, values, formulas. Read-only.
async function readWorkbook() {
  return Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();
    const used = sheets.items.map((ws) => {
      const r = ws.getUsedRangeOrNullObject(true); // true = values only, ignore formatting-only cells
      r.load("address,values,formulas");
      return { name: ws.name, r };
    });
    await ctx.sync();
    return used.map(({ name, r }) =>
      r.isNullObject ? { sheet: name, address: "", values: [], formulas: [] } : { sheet: name, address: r.address, values: r.values, formulas: r.formulas });
  });
}

// One block per sheet, one line per non-empty cell: "B5: 120 [=SUM(B2:B4)]". Skips the calling cell.
// maxRows: show only the first rows of each sheet (big data tables) and say how many were left out.
function serialiseWorkbook(wb, caller, maxRows = Infinity) {
  const blocks = [];
  for (const s of wb) {
    if (!s.address) continue;
    const { col: start, row } = parseAddress(s.address); // top-left of the used range
    const lines = [];
    s.values.slice(0, maxRows).forEach((vals, i) =>
      vals.forEach((v, j) => {
        if (isEmpty(v)) return;
        const addr = numToCol(start + j) + (row + i);
        if (caller && caller.sheet === s.sheet && caller.cell === addr) return;
        const f = s.formulas[i][j];
        lines.push(`${addr}: ${v}${typeof f === "string" && f.startsWith("=") ? ` [${f}]` : ""}`);
      })
    );
    const hidden = s.values.length - maxRows;
    if (hidden > 0) lines.push(`… ${hidden} more rows not shown (rows ${row + maxRows}–${row + s.values.length - 1})`);
    if (lines.length) blocks.push(`## Sheet "${s.sheet}" (${s.address.split("!").pop()})\n${lines.join("\n")}`);
  }
  return blocks.join("\n\n");
}

// Local models: a window that fits the prompt + answer, in 8k steps (each new size reloads the model).
const contextFor = (text) => Math.max(8192, Math.ceil((text.length / 1.4 + 1500) / 8192) * 8192);

function callerOf(invocation) {
  const a = parseAddress(invocation.address);
  return a ? { sheet: a.sheet, cell: numToCol(a.col) + a.row } : null;
}

// Not volatile: it runs when entered or edited, not on every change elsewhere (each run is a model call,
// and its own answer would otherwise change the workbook and trigger another run).
async function workbookAsk(prompt, model, invocation) {
  if (isEmpty(prompt)) return "";
  const m = model || WORKBOOK_DEFAULT_MODEL;
  const wb = await readWorkbook();
  const caller = callerOf(invocation);
  const max = WORKBOOK_MAX_CHARS;
  let text = serialiseWorkbook(wb, caller);
  for (const rows of WORKBOOK_ROW_STEPS) {
    if (text.length <= max) break;
    text = serialiseWorkbook(wb, caller, rows);
  }
  // Returned as text, not thrown: Excel would only show #VALUE! and hide the reason.
  if (text.length > max) return `⚠ Workbook too large (${text.length} chars even with 6 rows per sheet, max ${max}). Use LOKUP.H with a range.`;
  return generate(buildPrompt(String(prompt), text), m, invocation, undefined, undefined, { num_ctx: contextFor(text) });
}

/* ---------------------------------------------------------------- LOOOKUP.DO / LOOOKUP.UNDO (cell agent)
   DO plans cell edits and writes them once, after calculation.
   Ledger + backup live in the workbook's settings, so recalcs and reopening never repeat a run. */

// Ledger in the workbook's settings (saved in the file, no sheet): rows of [run, key, status, sheet, cell, old, new].
// A very-hidden ledger sheet created by the add-in froze Excel for Mac when the user switched sheets.
const LOG_SETTING = "LOOOKUP_log";
const DO_MAX_EDITS = 200;
const DO_MAX_RANGE_CELLS = 10000;
const DO_SYSTEM = `You edit a Microsoft Excel workbook by filling cells. Reply with JSON only, no prose:
{"edits":[{"cell":"Sheet1!B2","value":"text, a number, or a formula starting with ="},
          {"range":"Sheet1!C2:C100","value":"=A2*B2"}]}
To fill many cells with the same pattern, use ONE "range" edit: write the formula for the range's first cell; it is
copied down/across like Excel's fill handle (relative references adjust). Never list cells one by one for a pattern.
Rules: fill only empty cells; never delete anything; use the workbook's sheet names exactly; prefer formulas that
reference the data over typed-in numbers; at most ${DO_MAX_EDITS} edits.`;

// The workbook can carry injected instructions, so DO only writes formulas built from an allowlist of
// built-in functions that can't reach outside the workbook (no WEBSERVICE/IMAGE/add-ins/INDIRECT/LET/LAMBDA…),
// with no [book]/[table] brackets or DDE pipes. The formula is scanned the way Excel tokenises it.
const DO_SAFE_FUNCS = new Set([...KNOWN_FUNCS].filter((f) => !["INDIRECT", "LET", "LAMBDA"].includes(f)).concat(
  ("NPV IRR XNPV XIRR MIRR PMT PV FV RATE NPER IPMT PPMT YEARFRAC ISOWEEKNUM EXP LOG10 SUMSQ CORREL SLOPE INTERCEPT RSQ " +
   "TREND GROWTH FIXED DOLLAR NUMBERVALUE STDEV.S STDEV.P VAR.S VAR.P VAR STDEVP PERCENTILE.INC PERCENTILE.EXC " +
   "QUARTILE QUARTILE.INC QUARTILE.EXC RANK.EQ RANK.AVG MODE.SNGL CEILING.MATH FLOOR.MATH NORM.DIST NORM.S.DIST NORM.INV " +
   "FORECAST FORECAST.LINEAR NETWORKDAYS.INTL WORKDAY.INTL AVERAGEIFS MAXIFS MINIFS IFS SWITCH CONCAT TEXTJOIN XLOOKUP XMATCH").split(" ")));
const BARE_OK = /^(\$?[A-Z]{1,3}\$?\d+|\$?[A-Z]{1,3}|\$?\d+|TRUE|FALSE)$/i; // cell, column (A:A), row, boolean

// sheets: lower-cased names of this workbook's sheets; every Sheet! / 'Sheet'! prefix must be one of them.
function unsafeFormula(f, sheets) {
  const src = f.replace(/^[=+\-@]/, "");
  for (let i = 0; i < src.length; ) {
    const ch = src[i];
    if (ch === '"') { // string literal, "" escapes
      for (i++; i < src.length && !(src[i] === '"' && src[i + 1] !== '"'); i += src[i] === '"' ? 2 : 1);
      if (i >= src.length) return "unclosed string";
      i++;
    } else if (ch === "'") { // quoted sheet name, '' escapes; must be followed by ! and be a sheet of this workbook
      const start = i + 1;
      for (i++; i < src.length && !(src[i] === "'" && src[i + 1] !== "'"); i += src[i] === "'" ? 2 : 1);
      if (src[i + 1] !== "!") return "bad sheet reference";
      const name = src.slice(start, i).replace(/''/g, "'");
      if (!sheets.has(name.toLowerCase())) return `sheet ${name} not in this workbook`;
      i += 2;
    } else if (ch === "[" || ch === "]") return "brackets (link to another file / table) not allowed";
    else if (ch === "|") return "DDE link";
    else if (/[0-9]/.test(ch)) { const m = /^\d+(\.\d+)?(E[+-]?\d+)?/i.exec(src.slice(i)); i += m[0].length; }
    else if (/[A-Za-z_$\\]/.test(ch)) {
      const tok = /^[A-Za-z_$\\][A-Za-z0-9_.$\\]*/.exec(src.slice(i))[0];
      i += tok.length;
      const next = src.slice(i).trimStart()[0];
      if (next === "(") {
        const n = tok.toUpperCase().replace(/^_XLFN\./, "");
        if (!DO_SAFE_FUNCS.has(n)) return `${n} not allowed`;
      } else if (next !== "!" && !BARE_OK.test(tok)) return `name ${tok} not allowed`;
      else if (next === "!") {
        if (!sheets.has(tok.toLowerCase())) return `sheet ${tok} not in this workbook`;
        i = i + src.slice(i).indexOf("!") + 1;
      }
    } else if (/[\s+\-*/^&=<>%,;:(){}#]/.test(ch)) i++; // operators and punctuation
    else return `character ${JSON.stringify(ch)} not allowed`; // fail closed (e.g. non-ASCII names)
  }
  return null;
}

const hashKey =(s) => [...s].reduce((h, c) => (Math.imul(h, 31) + c.charCodeAt(0)) | 0, 7).toString(16);

// Model reply → [{ sheet, cell, value, skip? }]. Only cell values/formulas on existing sheets.
// Edit target "Sheet!A1" or "Sheet!A1:C9" → { sheet, cell (top-left), last (bottom-right, or null) }.
const TARGET = /^(?:'((?:[^']|'')+)'!|([^!']+)!)?\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?$/i;
const editCells = (e) => {
  if (!e.last) return [e.cell];
  const [c1, r1] = [colToNum(e.cell.replace(/\d+/, "")), Number(e.cell.replace(/[A-Z]+/, ""))];
  const [c2, r2] = [colToNum(e.last.replace(/\d+/, "")), Number(e.last.replace(/[A-Z]+/, ""))];
  const out = [];
  for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) out.push(numToCol(c) + r);
  return out;
};
// Cell count from the corners, before anything is listed (A1:XFD1048576 must not build 17 billion strings).
const editSize = (e) => !e.last ? 1 :
  (colToNum(e.last.replace(/\d+/, "")) - colToNum(e.cell.replace(/\d+/, "")) + 1) *
  (Number(e.last.replace(/[A-Z]+/, "")) - Number(e.cell.replace(/[A-Z]+/, "")) + 1);
const inSheet = (cell) => colToNum(cell.replace(/\d+/, "")) <= 16384 && Number(cell.replace(/[A-Z]+/, "")) <= 1048576 && Number(cell.replace(/[A-Z]+/, "")) >= 1;
const editLabel = (e) => `${e.sheet}!${e.cell}${e.last ? ":" + e.last : ""}`;

// Model reply → [{ sheet, cell, last, value, skip? }]. Only cell values/formulas on existing sheets.
function parsePlan(text, wb, caller) {
  const json = text.slice(text.search(/[[{]/), Math.max(text.lastIndexOf("}"), text.lastIndexOf("]")) + 1);
  let data;
  try {
    data = JSON.parse(json);
  } catch (e) {
    return null;
  }
  const raw = (Array.isArray(data) ? data : data.edits || []).slice(0, DO_MAX_EDITS);
  const names = new Map(wb.map((s) => [s.sheet.toLowerCase(), s.sheet]));
  const seen = new Set();
  let total = 0; // cells across all edits
  return raw.map((e) => {
    const target = straightQuotes(String((e && (e.range || e.cell)) || "").trim());
    const m = TARGET.exec(target);
    const sheet = m ? names.get(String(m[1] ? m[1].replace(/''/g, "'") : m[2] || caller.sheet).toLowerCase()) : null;
    let value = e && e.value;
    if (typeof value === "string" && value.startsWith("=")) value = straightQuotes(value);
    const edit = { sheet: sheet || (m && (m[1] || m[2])) || caller.sheet, cell: m ? m[3].toUpperCase() + m[4] : target, last: null, value };
    if (m && m[5]) {
      const [a, b] = [[colToNum(m[3].toUpperCase()), +m[4]], [colToNum(m[5].toUpperCase()), +m[6]]];
      edit.cell = numToCol(Math.min(a[0], b[0])) + Math.min(a[1], b[1]);
      edit.last = numToCol(Math.max(a[0], b[0])) + Math.max(a[1], b[1]);
      if (edit.last === edit.cell) edit.last = null;
    }
    const size = m ? editSize(edit) : 0;
    if (!m || !inSheet(edit.cell) || (edit.last && !inSheet(edit.last))) edit.skip = "bad cell";
    else if (!sheet) edit.skip = "no such sheet";
    else if (total + size > DO_MAX_RANGE_CELLS) edit.skip = `too many cells (max ${DO_MAX_RANGE_CELLS} per run)`;
    else if (!["string", "number", "boolean"].includes(typeof value)) edit.skip = "bad value";
    const cells = edit.skip ? [] : editCells(edit); // listed only after the size check
    if (edit.skip) return edit;
    if (cells.some((c) => seen.has(`${edit.sheet}!${c}`))) edit.skip = "duplicate";
    else if (typeof value === "string" && /^[=+\-@]/.test(value) && unsafeFormula(value, new Set(names.keys()))) edit.skip = `blocked: ${unsafeFormula(value, new Set(names.keys()))}`; // Excel treats +, -, @ as formula starts too
    if (!edit.skip) {
      cells.forEach((c) => seen.add(`${edit.sheet}!${c}`));
      total += size; // only accepted edits use up the budget
    }
    return edit;
  });
}

// The DO cell and the area its result table spills into are off limits.
function inPreviewArea(edit, caller, rows) {
  if (edit.sheet !== caller.sheet) return false;
  const cc = colToNum(caller.cell.replace(/\d+/, "")), cr = Number(caller.cell.replace(/[A-Z]+/, ""));
  return editCells(edit).some((cell) => {
    const c = colToNum(cell.replace(/\d+/, "")), r = Number(cell.replace(/[A-Z]+/, ""));
    return c >= cc && c <= cc + 2 && r >= cr && r <= cr + rows;
  });
}

// True when any target cell already has content (in the snapshot taken for planning).
function hasContent(wb, edit) {
  const s = wb.find((x) => x.sheet === edit.sheet);
  if (!s || !s.address) return false;
  const { col, row } = parseAddress(s.address);
  return editCells(edit).some((cell) => {
    const r = Number(cell.replace(/[A-Z]+/, "")) - row, c = colToNum(cell.replace(/\d+/, "")) - col;
    return ((s.formulas[r] || [])[c] ?? "") !== "";
  });
}

async function readLog(ctx) {
  const item = ctx.workbook.settings.getItemOrNullObject(LOG_SETTING);
  item.load("value");
  await ctx.sync();
  const rows = item.isNullObject ? [] : JSON.parse(item.value || "[]");
  return { rows, save: (r) => ctx.workbook.settings.add(LOG_SETTING, JSON.stringify(r)) };
}

// { done: run number if this key already ran, next: the number the next run will get }
async function findRun(key) {
  return Excel.run(async (ctx) => {
    const { rows } = await readLog(ctx);
    const r = rows.find((x) => x[1] === key);
    return { done: r ? Number(r[0]) : null, next: rows.reduce((m, x) => Math.max(m, Number(x[0]) || 0), 0) + 1 };
  });
}

// Writes the plan; skips cells that became non-empty since the preview. Returns { run, written, skipped }.
async function applyPlan(key, edits, run) {
  return Excel.run(async (ctx) => {
    const { rows, save } = await readLog(ctx);
    const prior = rows.find((x) => x[1] === key);
    if (prior) return { run: Number(prior[0]), already: true };
    const targets = edits.map((e) => {
      const ws = ctx.workbook.worksheets.getItem(e.sheet);
      const r = ws.getRange(e.last ? `${e.cell}:${e.last}` : e.cell);
      r.load("formulas");
      return { e, r, first: ws.getRange(e.cell) };
    });
    await ctx.sync();
    const todo = targets.filter(({ r }) => r.formulas.every((row) => row.every((f) => f === "")));
    todo.forEach(({ e, r, first }) => {
      first.formulas = [[e.value]];
      if (e.last) r.copyFrom(first, "Formulas"); // like the fill handle: relative references adjust per cell
    });
    todo.forEach(({ r }) => r.load("formulas")); // log what Excel stored (it normalises formulas), for UNDO's check
    await ctx.sync();
    if (todo.length) {
      const log = todo.flatMap(({ e, r }) => editCells(e).map((cell, i) =>
        [run, key, "applied", e.sheet, cell, "", String(r.formulas.flat()[i])]));
      save(rows.concat(log));
      await ctx.sync();
    }
    const cellCount = todo.reduce((n, { e }) => n + editCells(e).length, 0);
    return { run, written: cellCount, skipped: targets.length - todo.length, cells: new Set(todo.map(({ e }) => editLabel(e))) };
  });
}

const outcomes = new Map(); // key → result table of a run this session (also stops re-running a 0-edit run)
const DO_WRITE_DELAY_MS = 500; // write after Excel has finished this calculation

// Not streaming (Excel forbids stream + requiresAddress, and DO needs its own address). Plans, returns the result
// table at once, and writes the edits shortly after. It never makes Excel recalculate from inside the add-in:
// that re-entered the calculation engine and crashed Excel for Mac.
async function workbookDo(instruction, model, invocation) {
  try {
    if (isEmpty(instruction)) return [[""]];
    const m = model || WORKBOOK_DEFAULT_MODEL;
    const key = hashKey(m + "\u0000" + instruction);
    if (outcomes.has(key)) return outcomes.get(key);
    const { done, next } = await findRun(key);
    if (done) return [[`✓ Applied earlier · run ${done} · undo: =LOOOKUP.UNDO(${done})`]];
    const caller = callerOf(invocation);

    const wb = await readWorkbook();
    let text = serialiseWorkbook(wb, caller);
    for (const rows of WORKBOOK_ROW_STEPS) {
      if (text.length <= WORKBOOK_MAX_CHARS) break;
      text = serialiseWorkbook(wb, caller, rows);
    }
    if (text.length > WORKBOOK_MAX_CHARS) return [["⚠ Workbook too large for LOOOKUP.DO."]];
    const reply = await generate(buildPrompt(`Task: ${instruction}`, text), m, invocation, DO_SYSTEM, undefined,
      { num_ctx: contextFor(text), num_predict: 4096, format: "json" });
    const plan = parsePlan(reply, wb, caller);
    if (!plan) return [["⚠ The model did not return a usable plan. Try again or rephrase."]];
    if (!plan.length) return [["No edits proposed."]];
    for (const e of plan) {
      if (e.skip) continue;
      if (inPreviewArea(e, caller, plan.length + 1)) e.skip = "DO cell / result area";
      else if (hasContent(wb, e)) e.skip = "not empty";
    }
    const ok = plan.filter((e) => !e.skip);
    const n = ok.reduce((k, e) => k + editCells(e).length, 0);
    const table = [
      [`✓ ${n} cell${n === 1 ? "" : "s"} · run ${next} · undo: =LOOOKUP.UNDO(${next})`, "Value", "Status"],
      ...plan.map((e) => [editLabel(e), String(e.value ?? ""), e.skip ? `skipped: ${e.skip}` : "written"]),
    ];
    outcomes.set(key, table); // set before writing, so a recalc meanwhile can't plan/write a second time
    if (ok.length) {
      setTimeout(() => {
        applyPlan(key, ok, next).catch((e) => outcomes.set(key, [[`⚠ Could not write: ${e.message || e}`]]));
      }, DO_WRITE_DELAY_MS);
    }
    return table;
  } catch (e) {
    return [[`⚠ ${e.message || e}`]];
  }
}

// UNDO(run): restore what that run overwrote (empty cells), once. Cells changed since the run are left alone.
async function workbookUndo(run, invocation) {
  const show = (v) => invocation.setResult(v); // scalar string result
  if (isEmpty(run)) return show("Usage: =LOOOKUP.UNDO(run number shown by LOOOKUP.DO)");
  show("Undoing…");
  setTimeout(async () => {
    try {
      const msg = await Excel.run(async (ctx) => {
        const { rows, save } = await readLog(ctx);
        const mine = rows.map((x) => ({ x })).filter(({ x }) => Number(x[0]) === Number(run));
        if (!mine.length) return `No LOOOKUP.DO run ${run} in this workbook.`;
        const live = mine.filter(({ x }) => x[2] === "applied");
        if (!live.length) return `Run ${run} was already undone.`;
        const cells = live.map((l) => {
          const r = ctx.workbook.worksheets.getItem(l.x[3]).getRange(l.x[4]);
          r.load("formulas");
          return { ...l, r };
        });
        await ctx.sync();
        let restored = 0;
        for (const c of cells) {
          if (String(c.r.formulas[0][0]) === String(c.x[6])) {
            c.r.formulas = [[c.x[5]]];
            restored++;
          }
          c.x[2] = "undone";
        }
        save(rows);
        await ctx.sync();
        const kept = cells.length - restored;
        return `↩ Undid run ${run}: cleared ${restored} cell${restored === 1 ? "" : "s"}` + (kept ? ` · kept ${kept} you changed since` : "");
      });
      show(msg);
    } catch (e) {
      show(`⚠ Could not undo: ${e.message || e}`);
    }
  }, 0);
}

CustomFunctions.associate("H", workbookAsk);
CustomFunctions.associate("DO", workbookDo);
CustomFunctions.associate("UNDO", workbookUndo);
