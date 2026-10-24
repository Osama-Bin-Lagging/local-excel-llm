/* LOOOKUP — whole-workbook functions (shared runtime: custom functions can call the Excel API).
   =LOOOKUP.H(prompt, [model]) — answers in text using every sheet. Reuses generate() etc. from functions.js. */

const WORKBOOK_MAX_CHARS = 50000; // ≈ 34k tokens: cell lists are digit-heavy (measured ~1.5 chars/token with qwen3)
const WORKBOOK_ROW_STEPS = [50, 25, 12, 6]; // too big → show only this many top rows per sheet, smallest last

// Read every sheet's used range: name, address, values, formulas. Read-only.
async function readWorkbook() {
  return Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();
    const used = sheets.items.filter((ws) => ws.name !== LOG_SHEET).map((ws) => {
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
  const m = model || DEFAULT_MODEL;
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
   DO previews a plan of cell edits; with apply = TRUE it writes that plan once, after calculation.
   Ledger + backup live on a very-hidden sheet, so recalcs and reopening never repeat a run. */

const LOG_SHEET = "LOOOKUP_log"; // columns: run | key | status | sheet | cell | old | new
const DO_MAX_EDITS = 200;
const DO_SYSTEM = `You edit a Microsoft Excel workbook by filling cells. Reply with JSON only, no prose:
{"edits":[{"cell":"Sheet1!B2","value":"text, a number, or a formula starting with ="}]}
Rules: fill only empty cells; never delete anything; use the workbook's sheet names exactly; prefer formulas that
reference the data over typed-in numbers; at most ${DO_MAX_EDITS} edits.`;
const plans = new Map(); // key → edits from the last preview, so apply writes exactly what was shown

const hashKey = (s) => [...s].reduce((h, c) => (Math.imul(h, 31) + c.charCodeAt(0)) | 0, 7).toString(16);

// Model reply → [{ sheet, cell, value, skip? }]. Only cell values/formulas on existing sheets.
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
  return raw.map((e) => {
    const m = /^(?:'((?:[^']|'')+)'!|([^!']+)!)?\$?([A-Z]{1,3})\$?(\d+)$/i.exec(straightQuotes(String(e && e.cell || "").trim()));
    const sheet = m ? names.get(String(m[1] ? m[1].replace(/''/g, "'") : m[2] || caller.sheet).toLowerCase()) : null;
    const cell = m ? m[3].toUpperCase() + m[4] : String(e && e.cell);
    let value = e && e.value;
    if (typeof value === "string" && value.startsWith("=")) value = straightQuotes(value);
    const edit = { sheet: sheet || (m && (m[1] || m[2])) || caller.sheet, cell, value };
    const id = `${edit.sheet}!${cell}`;
    if (!m) edit.skip = "bad cell";
    else if (!sheet) edit.skip = "no such sheet";
    else if (!["string", "number", "boolean"].includes(typeof value)) edit.skip = "bad value";
    else if (seen.has(id)) edit.skip = "duplicate";
    seen.add(id);
    return edit;
  });
}

// The DO cell and the area its preview spills into are off limits.
function inPreviewArea(edit, caller, rows) {
  if (edit.sheet !== caller.sheet) return false;
  const c = colToNum(edit.cell.replace(/\d+/, "")), r = Number(edit.cell.replace(/[A-Z]+/, ""));
  const cc = colToNum(caller.cell.replace(/\d+/, "")), cr = Number(caller.cell.replace(/[A-Z]+/, ""));
  return c >= cc && c <= cc + 2 && r >= cr && r <= cr + rows;
}

function currentValue(wb, edit) {
  const s = wb.find((x) => x.sheet === edit.sheet);
  if (!s || !s.address) return "";
  const { col, row } = parseAddress(s.address);
  const r = Number(edit.cell.replace(/[A-Z]+/, "")) - row, c = colToNum(edit.cell.replace(/\d+/, "")) - col;
  return (s.formulas[r] || [])[c] ?? "";
}

async function readLog(ctx) {
  const ws = ctx.workbook.worksheets.getItemOrNullObject(LOG_SHEET);
  await ctx.sync();
  if (ws.isNullObject) return { ws: null, rows: [] };
  const used = ws.getUsedRange();
  used.load("values");
  await ctx.sync();
  return { ws, rows: used.values.slice(1) }; // row 1 = header
}

async function findRun(key) {
  return Excel.run(async (ctx) => {
    const r = (await readLog(ctx)).rows.find((x) => x[1] === key);
    return r ? Number(r[0]) : null;
  });
}

// Writes the plan; skips cells that became non-empty since the preview. Returns { run, written, skipped }.
async function applyPlan(key, edits) {
  return Excel.run(async (ctx) => {
    let { ws, rows } = await readLog(ctx);
    const prior = rows.find((x) => x[1] === key);
    if (prior) return { run: Number(prior[0]), already: true };
    if (!ws) {
      ws = ctx.workbook.worksheets.add(LOG_SHEET);
      const head = ws.getRange("A1:G1");
      head.numberFormat = [Array(7).fill("@")];
      head.values = [["run", "key", "status", "sheet", "cell", "old", "new"]];
      ws.visibility = Excel.SheetVisibility.veryHidden;
    }
    const run = rows.reduce((m, x) => Math.max(m, Number(x[0]) || 0), 0) + 1;
    const targets = edits.map((e) => {
      const r = ctx.workbook.worksheets.getItem(e.sheet).getRange(e.cell);
      r.load("formulas");
      return { e, r };
    });
    await ctx.sync();
    const todo = targets.filter(({ r }) => r.formulas[0][0] === "");
    todo.forEach(({ e, r }) => (r.formulas = [[e.value]]));
    todo.forEach(({ r }) => r.load("formulas")); // log what Excel stored (it normalises formulas), for UNDO's check
    await ctx.sync();
    if (todo.length) {
      const log = todo.map(({ e, r }) => [String(run), key, "applied", e.sheet, e.cell, "", String(r.formulas[0][0])]);
      const at = ws.getRange(`A${rows.length + 2}:G${rows.length + 1 + log.length}`);
      at.numberFormat = log.map((x) => x.map(() => "@")); // text, so logged formulas stay inert
      at.values = log;
      await ctx.sync();
    }
    return { run, written: todo.length, skipped: targets.length - todo.length };
  });
}

// Streaming: the cell shows progress, and the writes happen after this calculation has finished.
async function workbookDo(instruction, apply, model, invocation) {
  const show = (v) => invocation.setResult(Array.isArray(v) ? v : [[v]]);
  try {
    if (isEmpty(instruction)) return show("");
    const m = model || DEFAULT_MODEL;
    const key = hashKey(m + "\u0000" + instruction);
    const done = await findRun(key);
    if (done) return show(`✓ Applied earlier · run ${done} · undo: =LOOOKUP.UNDO(${done})`);

    if (apply && plans.has(key)) {
      show("Applying…");
      const edits = plans.get(key).filter((e) => !e.skip);
      setTimeout(async () => {
        try {
          const r = await applyPlan(key, edits);
          if (r.already) return show(`✓ Applied earlier · run ${r.run} · undo: =LOOOKUP.UNDO(${r.run})`);
          show(`✓ Applied ${r.written} edit${r.written === 1 ? "" : "s"} · run ${r.run}` +
            (r.skipped ? ` · skipped ${r.skipped} (no longer empty)` : "") + ` · undo: =LOOOKUP.UNDO(${r.run})`);
        } catch (e) {
          show(`⚠ Could not apply: ${e.message || e}`);
        }
      }, 0);
      return;
    }

    show("Planning…");
    const wb = await readWorkbook();
    const caller = callerOf(invocation);
    let text = serialiseWorkbook(wb, caller);
    for (const rows of WORKBOOK_ROW_STEPS) {
      if (text.length <= WORKBOOK_MAX_CHARS) break;
      text = serialiseWorkbook(wb, caller, rows);
    }
    if (text.length > WORKBOOK_MAX_CHARS) return show("⚠ Workbook too large for LOOOKUP.DO.");
    const reply = await generate(buildPrompt(`Task: ${instruction}`, text), m, invocation, DO_SYSTEM, undefined,
      { num_ctx: contextFor(text), num_predict: 4096, format: "json" });
    const plan = parsePlan(reply, wb, caller);
    if (!plan) return show("⚠ The model did not return a usable plan. Try again or rephrase.");
    for (const e of plan) {
      if (e.skip) continue;
      if (inPreviewArea(e, caller, plan.length + 1)) e.skip = "DO cell / preview area";
      else if (currentValue(wb, e) !== "") e.skip = "not empty";
    }
    plans.set(key, plan);
    const ok = plan.filter((e) => !e.skip).length;
    if (!plan.length) return show("No edits proposed.");
    show([
      [`Preview: ${ok} edit${ok === 1 ? "" : "s"}${apply ? " (re-enter this cell to apply)" : " (set apply to TRUE to write)"}`, "Now", "New"],
      ...plan.map((e) => [`${e.sheet}!${e.cell}`, e.skip ? "" : "(empty)", e.skip ? `skipped: ${e.skip}` : e.value]),
    ]);
  } catch (e) {
    show(`⚠ ${e.message || e}`);
  }
}

// UNDO(run): restore what that run overwrote (empty cells), once. Cells changed since the run are left alone.
async function workbookUndo(run, invocation) {
  const show = (v) => invocation.setResult([[v]]);
  if (isEmpty(run)) return show("Usage: =LOOOKUP.UNDO(run number shown by LOOOKUP.DO)");
  show("Undoing…");
  setTimeout(async () => {
    try {
      const msg = await Excel.run(async (ctx) => {
        const { ws, rows } = await readLog(ctx);
        const mine = rows.map((x, i) => ({ x, row: i + 2 })).filter(({ x }) => Number(x[0]) === Number(run));
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
          ws.getRange(`C${c.row}`).values = [["undone"]];
        }
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
