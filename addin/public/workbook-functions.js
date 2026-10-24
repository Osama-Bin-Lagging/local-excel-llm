/* LOOOKUP — whole-workbook functions (shared runtime: custom functions can call the Excel API).
   =LOOOKUP.H(prompt, [model]) — answers in text using every sheet. Reuses generate() etc. from functions.js. */

const WORKBOOK_MAX_CHARS = { local: 20000, claude: 400000 }; // local: fits num_ctx 8192 with room for the answer
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
  const max = WORKBOOK_MAX_CHARS[isClaude(m) ? "claude" : "local"];
  let text = serialiseWorkbook(wb, caller);
  for (const rows of WORKBOOK_ROW_STEPS) {
    if (text.length <= max) break;
    text = serialiseWorkbook(wb, caller, rows);
  }
  // Returned as text, not thrown: Excel would only show #VALUE! and hide the reason.
  if (text.length > max) return `⚠ Workbook too large for ${m} (${text.length} chars, max ${max}). Use LOKUP.H with a range${isClaude(m) ? "" : ', or model "claude"'}.`;
  return generate(buildPrompt(String(prompt), text), m, invocation);
}

CustomFunctions.associate("H", workbookAsk);
