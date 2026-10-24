# Product Requirements Document: Local LLM in Excel

## Product Overview
**Product Vision:** Call a locally running LLM from an Excel cell (`=LOKUP.H(prompt, range)`) with no data leaving the machine.
**Target Users:** Single user (owner), Excel 16.x on macOS (Apple M4, 16 GB).
**Success Metrics:**
- ≥ 80% PASS on graded rows of `llm_test_sheet.xlsx` with the chosen default model
- ≤ 1.5 s median per cell for short answers (warm model)
- Works with Wi-Fi off

## Feature Requirements
| Feature | Description | Priority | Acceptance Criteria |
|---|---|---|---|
| `LOKUP.H(prompt, [context], [headers], [model])` | Cell function returning model text | Must | Returns answer in a cell; T01–T15 and F01–F18 runnable from the Tests tab |
| `LOKUP.V(request, [data], [model])` | Returns an Excel formula as text; sends the data range's address + column map so references are real | Must | ≥ 70% PASS on Formula Tests (F01–F18) with the default model |
| Context from ranges | Range args serialised into the prompt | Must | Row (T10–T11), table (T12–T14) and formula (F01–F18) tests produce grounded answers |
| Offline | Model via Ollama on localhost; add-in files self-hosted (no CDN) | Must | Function works with network disabled |
| Result cache | Same prompt+context+model → no new model call on recalc | Must | Recalc of an unchanged sheet returns instantly |
| Empty input → blank | No model call for empty prompt/context | Must | T04 returns blank |
| Clear errors | Ollama down / model missing → readable `#N/A` message, not a hang | Must | Stop Ollama → cell shows error within 5 s |
| Date handling in context | Excel date serials rendered as dates | Should | T11 passes |
| Prompt-injection guard | Context delimited and marked as data | Should | T05 passes |
| Context size cap | Large ranges rejected with message | Should | 10k-row range → error, not a freeze |
| Model picker / settings pane | Task pane UI for default model | Could | — |
| Streaming into cell | Partial output while generating | Won't (v0) | — |
| Windows support | — | Won't (v0) | — |

## User Flows
1. Start Ollama (auto at login) → start add-in server → open Excel → add-in loaded → type `=LOKUP.H("…", A2)`.
   - Ollama down → cell `#N/A` "Ollama not reachable at localhost:11434".
   - Model not pulled → `#N/A` "model 'x' not found — run: ollama pull x".

## Non-Functional Requirements
- **Performance:** temperature 0, thinking disabled, short max-token cap; model kept warm (`keep_alive`).
- **Security:** server binds to localhost only; context treated as untrusted data.
- **Compatibility:** Excel for Mac 16.x (Microsoft 365). Sideloaded developer add-in.

## Technical Specifications
- Office add-in (XML manifest) with JS custom functions; Node HTTPS server on `https://localhost:3000` serving add-in files and proxying `/api/*` → Ollama `http://localhost:11434` (same origin: avoids CORS and HTTPS→HTTP mixed-content blocking).
- Dev TLS cert via `office-addin-dev-certs`.
- Office runtime self-hosted: the whole `dist/` of `@microsoft/office-js@1.1.110` (minus debug builds) is copied into `public/`. `custom-functions-runtime.js` alone is not enough — on Mac it also loads `excel-mac-16.00-core.js` and `en-us/office_strings.js` from its own folder, and hangs at `#BUSY` if they 404.

## Release Planning
- **v0 (prototype):** ASK function, cache, errors, basic context serialisation. Validate on the test sheet.
- **v1:** chosen default model from benchmark, date handling, injection guard, size cap.

## Open Questions & Assumptions
- ✅ Confirmed 2026-10-23: sideloading via `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef` works on Excel 16.111 (Home → Add-ins → Developer Add-ins), even with the unlicensed "Subscription Required" banner.
- ✅ Confirmed 2026-10-23: Excel's add-in web view trusts the `office-addin-dev-certs` CA; `=LOKUP.H` returned answers end to end.
- Assumption: `requiresParameterAddresses` is supported by Excel 16.111 for Mac (needed by LOKUP.V). If not, the function still runs but without the column map.
- Open: formula grading is by required parts (PASS/CHECK/FAIL); true evaluation of the generated formula would need writing it into a cell — possible later via a test harness.
- Open: context serialisation format (markdown vs CSV vs key:value) — decide using T10–T14 and F01–F18 results.
- Open: default model. Results 2026-10-23 (harness, 30 formula cases, prompt v4; F01–F18 tuned-on, F19–F30 cleaning held out):
  qwen2.5-coder:7b 18/18 + 7/12 (4.65 GB, 1.2 s/cell on M4) · phi4-mini 16/18 + 4/12 (2.5 GB) · qwen3:4b-instruct 14/18 + 5/12 (2.9 GB) · gemma3:4b 12/18 + 5/12 · qwen2.5-coder:3b 12/18 + 1/12 · llama3.2:3b 7/18 + 1/12. qwen3:8b / 14b pending.
  Lesson: scores on tuned cases overstate quality; always keep a held-out set when changing FORMULA_SYSTEM.
- Note: Ollama `qwen3:4b` is the 2507 thinking-only build and ignores `think:false` — use `qwen3:4b-instruct`.
- Open: should cache persist across Excel restarts? (v0: in-memory only.)

## Results log — 2026-10-23 (final round)
Setup: checker + auto sheet names + 1 retry; matched examples (`dyn` for ≥7B, `compact` for 4B); num_ctx 8192; small tables shown whole.
Scored on: held-out v3 (10 unseen formula tasks) and the two downloaded assessments (57 cells, isolated per-cell scoring in Excel; key formulas score 57/57).

| Model | RAM | Held-out v3 | IB (34) | DA (23) | Real total | s/cell (real, M4) |
|---|---|---|---|---|---|---|
| qwen3:8b +dyn | 5.2 GB | 8/10 | 20 | 18 | 38/57 | 7.7 |
| qwen2.5-coder:7b +dyn | 4.8 GB | 7/10 | 19 | 17 | 36/57 | 7.4 |
| qwen3:4b-instruct +compact | 3.3 GB | 6/10 | 18 | 19 | 37/57 | 4.8 |
| gemma4:e4b +dyn | 14.7 GB | 6/10 | 16 | 21 | 37/57 | 6.3 |
| gemma3:4b +compact | 2.4 GB | 5/10 | 4 | 12 | 16/57 | 4.6 |
| baseline (no checker/examples): 4B 17/57, 7B 18/57 | | | | | | |

- Strong: cells with written guidance (Deal Model 9–11/12), links (Summary), data metrics (DA).
- Weak: label-only finance cells (Trading Comps 0–2/10, Operating Data 0–2/4): models look up the empty target column (circular) instead of computing price × shares, EBITDA / revenue, etc.
- Bugs found via the real files and fixed: Ollama default 4096-token context silently truncated prompts; 8-row cap hid model-sheet inputs; circular self-references opened Excel dialogs; multi-word sheet names.
- qwen3:14b: held-out v3 7/10 but ~40–60 s/cell on real sheets → impractical; real-test run stopped.
- The two assessments are now dev data (failures inspected). Any further tuning needs a fresh real workbook as the test.

## v1.1 — add-ons for qwen3:4b-instruct (2026-10-23)
Target model: qwen3:4b-instruct (3.3 GB in memory → fits an 8 GB Mac). Each add-on has a switch in `FORMULA_OPTS`.

| Feature | Description | Priority | Acceptance criteria |
|---|---|---|---|
| Glossary (83 terms) | Standard finance/accounting/SaaS/ops/stats definitions; the ≤6 matching the task/headers go into the prompt | Must | No test-specific wording (scanned); helps on unseen workbooks (v4: 2→14/23) |
| P1 Target-cell awareness | Prompt names the cell being filled, its header/row label and the row's inputs; forbids self-reference | Must | Fewer circular-reference answers |
| P2 Run formula in JS | Offline engine (fast-formula-parser + formula.js, `formula-parser.js`) evaluates candidates on the sheet data; Excel errors → retry | Must | 0 false errors / 0 wrong values on 96 key formulas; array-style formulas skipped, never penalised |
| P3 Expected result type | count / fraction / Yes-No / text inferred from the task's subject; mismatch → retry | Should | 0 correct answers flagged across 127 tasks |
| P4 Vote | On failure, 3 seeded samples (T=0.7, top_p 0.8, top_k 20); most common evaluated value wins | Should | Deterministic (seeded); only costs time when the first answer fails |

- Held-out v5 (payroll, campaigns, shipments; 16 tasks) frozen before the add-ons were written; run once at the end.
- Lesson: every check is validated against the answer keys before use — first versions of P2 and P3 both had false alarms (array maths; "× interest rate" read as a percentage).

### v1.1 results (qwen3:4b-instruct, compact prompt; grader counts hard-coded / circular / unknown-sheet formulas as wrong)
| Setup | Dev 128 | v5 16* | s/cell (M4) |
|---|---|---|---|
| Glossary only | 96 | 10 | 4.2 |
| Add-ons without P1 | 97 | 10 | 4.7 |
| All add-ons, vote on failure | 99 | 11 | 3.7 |
| **All add-ons, vote every cell (default)** | **102** | **12** | 6.0 |
*v5's first blind run: 10 (glossary only) vs 11 (add-ons, P1 v1, scored honestly); the table row is a second look after P1 was fixed.
- P1 v1 listed row values → model hard-coded numbers (7/80 answers). Fixed: addresses + headers only, plus a hard-coding check.
- Defaults set: DEFAULT_MODEL = qwen3:4b-instruct; FORMULA_OPTS prompt "compact", vote "always".

## Rename (2026-10-24)
- `LLAMA.ASK` → `LOKUP.H`, `LLAMA.FORMULA` → `LOKUP.V` (namespace `LOKUP`; Excel requires a namespace, so a bare `=HLOKUP()` is not possible). Manifest version bumped to 0.2.0.0 to bust Excel's add-in cache.

## qwen3.5:4b evaluation (2026-10-24)
Same harness, 144 tasks (dev 128 + v5 16). Uses ~12.7 GB in memory with real prompts (vs 4.1 GB on a one-line test) → not viable on 8 GB.
| Setup | Dev 128 | v5 16 | s/cell |
|---|---|---|---|
| qwen3:4b-instruct, add-ons (default) | **102** | **12** | 6.0 |
| qwen3.5:4b, add-ons, presence_penalty 0 | 79 | 9 | 12.3 |
| qwen3.5:4b, add-ons, built-in penalty 1.5 | 76 | 8 | 8.6 |
| qwen3.5:4b, no add-ons, penalty 0 | 70 | 8 | 8.3 |
| qwen3.5:4b + thinking | stopped: 105–708 s per cell | | |
Decision: keep qwen3:4b-instruct. Lesson: check built-in model parameters (`ollama show`) and memory under real prompts before judging a model.

## v1.2 — Claude backend, whole-workbook functions, cell agent (proposed 2026-10-24)

### User needs
- Use a stronger cloud model for hard cells without leaving Excel (accept that those cells go online).
- Ask about / write formulas across the **whole workbook**, not one range.
- Get "GPT-for-Excel"-style actions (build sheets, fill formulas) from a **cell** instead of a sidebar.

### Feature requirements
| Feature | Description | Priority | Acceptance criteria |
|---|---|---|---|
| Claude backend | `model` = `"claude"` (Sonnet), `"claude:haiku"`, `"claude:sonnet"`, `"claude:opus"` routes the call through `claude -p` via the add-in server | Must | Works in LOKUP.H/V; tools disabled (`--tools ""`), no MCP/hooks/settings/session; ≤ 3 concurrent; 90 s timeout; clear `#N/A` if `claude` missing or logged out |
| Claude benchmark | Sonnet + Haiku on the 144 tasks | Must | Scores vs qwen3:4b-instruct (102/128, 12/16) recorded here |
| Workbook functions | `=LOOOKUP.H(prompt,[model])` (text) and `=LOOOKUP.V(request,[model])` (formula) using every sheet — second add-in, shared runtime | Should | Snapshot at call time; big sheets shown as top rows + "N more rows not shown" (never a bare #VALUE!); works with local and Claude models. **H built 2026-10-24**, V to do |
| Cell agent | `=LOOOKUP.DO(instruction,[model])` fills other cells with values/formulas (no preview: user decision 2026-10-24); `=LOOOKUP.UNDO(run)` reverses one run | Could | (1) Writes the plan **once**, then shows a spilled result table (cell · value · written/skipped + reason). (2) a ledger in the workbook settings (`LOOOKUP_log`; a very-hidden ledger sheet froze Excel for Mac on sheet switch) makes recalcs/reopens never repeat it; cell shows "✓ Applied N · run k". (3) Only sets cell values/formulas (≤ 200); never deletes, never creates/renames sheets, never writes into non-empty cells, its own cell or its preview area, or sheets that don't exist. (4) Old contents backed up in the ledger; UNDO(k) restores them once, skipping cells changed since. (5) Writes happen after calculation (deferred), result streamed into the cell |
| Default models | **LOKUP** stays local by default (qwen3:4b-instruct); **LOOOKUP** defaults to Claude Sonnet (user decision 2026-10-24: whole-workbook prompts are ~10–35k tokens, ~2 min on the local 4B vs seconds on Sonnet) | Must | LOKUP cells never go online unless they name a `claude…` model; LOOOKUP cells go online unless they name a local model; README states this |

### Risks / open questions
- Shared runtime on Excel 16.111 for Mac: supported in docs, untested here — validate with a spike before building on it.
- Writing to the workbook from a custom function's run is discouraged by Microsoft; edits are deferred until after calculation. Excel's Undo does not cover add-in edits → cell agent must preview first.
- Claude calls count toward the user's Claude usage limits; prompts and data leave the machine for those cells.
- Naming decided: second add-in `LOOOKUP` (.H, .V, .DO, .UNDO).
- Excel for Mac caches add-in metadata and ignores version bumps → `./reload.sh` + re-insert after changing functions.json.
- DO with a 4B local model on big workbooks is unproven; preview-first is the guard against bad plans.
