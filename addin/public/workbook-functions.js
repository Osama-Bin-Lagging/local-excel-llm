/* LOOOKUP — whole-workbook functions (shared runtime: custom functions can call the Excel API).
   =LOOOKUP.H(prompt, [model]) — answers in text using every sheet. Reuses generate() etc. from functions.js. */

const WORKBOOK_MAX_CHARS = { local: 20000, claude: 400000 }; // local: fits num_ctx 8192 with room for the answer

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
function serialiseWorkbook(wb, caller) {
  const blocks = [];
  for (const s of wb) {
    if (!s.address) continue;
    const { col: start, row } = parseAddress(s.address); // top-left of the used range
    const lines = [];
    s.values.forEach((vals, i) =>
      vals.forEach((v, j) => {
        if (isEmpty(v)) return;
        const addr = numToCol(start + j) + (row + i);
        if (caller && caller.sheet === s.sheet && caller.cell === addr) return;
        const f = s.formulas[i][j];
        lines.push(`${addr}: ${v}${typeof f === "string" && f.startsWith("=") ? ` [${f}]` : ""}`);
      })
    );
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
  const text = serialiseWorkbook(await readWorkbook(), callerOf(invocation));
  const max = WORKBOOK_MAX_CHARS[isClaude(m) ? "claude" : "local"];
  if (text.length > max) {
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.invalidValue,
      `Workbook too large (${text.length} chars, max ${max} for ${m}). Use LOKUP.H with a range${isClaude(m) ? "" : ', or model "claude"'}.`
    );
  }
  return generate(buildPrompt(String(prompt), text), m, invocation);
}

CustomFunctions.associate("H", workbookAsk);
