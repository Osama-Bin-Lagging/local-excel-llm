# LOKUP — a local AI model inside Excel (Mac)

Two Excel functions backed by a model running on your own Mac (Ollama). Nothing leaves the machine.

| Function | Use |
|---|---|
| `=LOKUP.V(request, [data], [model])` | Writes an Excel **formula** for the request, using the data range's real addresses |
| `=LOKUP.H(prompt, [context], [headers], [model])` | Free-text answer (classify, extract, explain) |

Example: `=LOKUP.V("Total revenue for the North region", Sales!A1:F13)` → `=SUMIFS(Sales!F2:F13,Sales!A2:A13,"North")`
The formula comes back as text — paste it into a cell to use it. Describing the calculation helps
("Equity value = share price × diluted shares").

## Requirements
- Mac with Apple Silicon, 8 GB RAM or more (the default model uses ~3.3 GB)
- Microsoft Excel for Mac (16.x) · [Homebrew](https://brew.sh)

## Setup (once)
```sh
git clone <this repo> lokup && cd lokup
./setup.sh
```
Installs Ollama + Node, downloads `qwen3:4b-instruct`, trusts a localhost-only HTTPS certificate
(asks for your password), registers the add-in with Excel, then runs `./test.sh`.

## Every day
```sh
./start.sh        # Ollama + add-in server; keep the window open
```
Excel → **Home → Add-ins → Developer Add-ins → Local LLM** (once per Excel launch).

## Check it works
```sh
./test.sh         # ~1–2 min: Ollama, model, HTTPS, server, LOKUP.H + 6 LOKUP.V formulas with known answers
```

## Using Claude for hard cells (optional, online)
Put a Claude model in the `model` argument — the cell's prompt and data then go to Anthropic via the
[Claude Code](https://claude.com/claude-code) CLI on this Mac (must be installed and logged in; counts toward your Claude usage):
```
=LOKUP.V("Pro forma EPS", 'Deal Model'!A4:E29, "claude")        ← Sonnet
=LOKUP.H("Summarise this review", A2, , "claude:haiku")          ← Haiku (fastest)
```
`"claude"` = Sonnet; also `"claude:haiku"`, `"claude:sonnet"`, `"claude:opus"`. Tools are disabled — Claude only returns text.
Cells without a `claude…` model stay fully local.

## Whole-workbook questions: LOOOKUP (second add-in)
Insert **Local LLM Workbook** too (Home → Add-ins). It reads every sheet itself — no ranges to pass.
**LOOOKUP uses Claude Sonnet by default, so the workbook's contents go to Anthropic** (via the Claude Code CLI).
Pass a local model name to keep it offline:
```
=LOOOKUP.H("Which region had the highest margin?")                    ← Claude Sonnet
=LOOOKUP.H("Any inconsistencies between Summary and Deal Model?", "qwen3:4b-instruct")   ← local
=LOOOKUP.DO("fill Gross Sales for the first 5 orders in Raw Data")    ← writes once (empty cells only)
=LOOOKUP.UNDO(1)                                                       ← undo run 1
```
Runs when entered or edited (not on every change elsewhere — each run is a model call). Up to ~50k characters
of cell content; bigger sheets are cut to their top rows. Local models are slow on big workbooks (~2 min for 25k chars on an M-series Air).

## Settings
In `addin/public/functions.js`:
- `DEFAULT_MODEL` — `qwen3:4b-instruct` (best for 8 GB). On 16 GB, `qwen3:8b` is a bit more accurate.
- `FORMULA_OPTS.vote` — `"always"` (most accurate, ~6 s/cell on M4) or `"fail"` (faster, ~3.7 s/cell).
- `EXCEL_365` — set `true` once Excel is signed in to Microsoft 365 (enables XLOOKUP etc.).

## Layout
```
addin/            Excel add-in: manifest, server.js (HTTPS + proxy to Ollama), public/ (functions, engine, Office runtime)
harness/          Benchmark: run.js (runs the add-in code), grade.py / real/grade_real.py (grade in Excel), test cases
llm_test_sheet.xlsx  Test workbook (48 formula/text cases)
prd.md            Requirements, design decisions and all benchmark results
```

## Troubleshooting
- Cells show `#BUSY` forever → the add-in server isn't running (`./start.sh`).
- `#N/A … Ollama not reachable` → `brew services start ollama`.
- Excel runs old add-in code or shows old function names (it caches add-ins and ignores version bumps) →
  quit Excel, `./reload.sh`, reopen Excel, then Home → Add-ins → insert the add-in(s) again.
- Slow on 8 GB → close Chrome/Docker; set `FORMULA_OPTS.vote = "fail"`.
