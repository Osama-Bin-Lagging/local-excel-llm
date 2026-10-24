"""Grade harness results inside real Excel.

For each results-<model>.json: copy the test workbook, put every model formula into
'Formula Tests' as text (J → sheet's own PASS/CHECK/FAIL) and as a live formula (column O),
let Excel calculate, then compare O's value with the reference value in G.

Usage: python3 grade.py ../llm_test_sheet.xlsx results-a.json [results-b.json ...]
"""
import json, os, re, shutil, subprocess, sys, tempfile

SRC, result_files = sys.argv[1], sys.argv[2:]
SEP = "\x1f"
KNOWN_SHEETS = {"README", "Tests", "Formula Tests", "Summary", "Reviews", "Contacts", "Sales", "Messy", "Context Design"}


def q(s):  # AppleScript string literal
    return '"' + s.replace("\\", "\\\\").replace('"', '\\"') + '"'


def same(got, ref):
    try:
        a, b = float(got.replace(",", "")), float(ref.replace(",", ""))
        return abs(a - b) <= 1e-6 * max(1, abs(b))
    except ValueError:
        return got.strip() == ref.strip()  # case-sensitive: "AR" vs "Ar" matters for cleaning tasks


def verdict(r):
    if r["value_ok"] is None:
        return "✓" if r["parts"] == "PASS" else "✗"
    if r["value_ok"]:
        return "✓" if r["parts"] == "PASS" else "~"
    return "✗"


def grade(results_path):
    results = json.load(open(results_path))
    tmpdir = tempfile.mkdtemp()
    # unique name per run: a copy left stuck in Excel by an earlier run can't make the name ambiguous
    name = f"grade_{os.getpid()}_{os.path.basename(results_path)[8:-5]}.xlsx"
    copy = os.path.join(tmpdir, name)
    shutil.copy(SRC, copy)
    lines = [
        "with timeout of 180 seconds",
        'tell application "Microsoft Excel"',
        f"open POSIX file {q(copy)}",
        # Excel can return from `open` before the workbook exists; retry briefly
        f"repeat 20 times\ntry\nset wb to workbook {q(name)}\nexit repeat\non error\ndelay 0.5\nend try\nend repeat",
        'set ft to worksheet "Formula Tests" of wb',
    ]
    for r in results:
        out, row = r["out"], r["row"]
        out = " ".join(out.split("\n"))[:250]  # Excel/AppleScript reject multi-line or very long cell text
        r["out"] = out
        lines.append(f'try\nset value of range "J{row}" of ft to {q(chr(39) + out)}\non error\nset value of range "J{row}" of ft to "(unwritable output)"\nend try')
        bad_sheets = {m.strip("'") for m in re.findall(r"('[^']+'|[A-Za-z_][\w.]*)!", out)} - KNOWN_SHEETS
        if bad_sheets:
            # Excel opens a blocking "Cannot find X. Copy from:" file picker for unknown sheets
            r["bad_sheets"] = sorted(bad_sheets)
            lines.append(f'set value of range "O{row}" of ft to "#BADSHEET"')
        elif out.startswith("="):
            lines.append(f'try\nset formula of range "O{row}" of ft to {q(out)}\non error\nset value of range "O{row}" of ft to "#SYNTAX"\nend try')
    # dates: Excel auto-formats DATE() results as dates; compare raw serials like the reference column
    lines += ['set number format of range "O2:O60" of ft to "General"', "calculate", 'set acc to ""']
    for r in results:
        row = r["row"]
        lines.append(
            f'set acc to acc & (string value of range "A{row}" of ft) & {q(SEP)} & (string value of range "M{row}" of ft) & {q(SEP)}'
            f' & (string value of range "G{row}" of ft) & {q(SEP)} & (string value of range "O{row}" of ft) & linefeed')
    lines += ["close wb saving no", "return acc", "end tell", "end timeout"]
    script = "\n".join(lines)
    open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "last_grade.applescript"), "w").write(script)
    out = subprocess.run(["osascript", "-"], input=script, capture_output=True, text=True)
    shutil.rmtree(tmpdir, ignore_errors=True)
    if out.returncode:
        sys.exit(out.stderr)
    by_id = {r["id"]: r for r in results}
    rows = []
    for line in out.stdout.strip().splitlines():
        fid, parts, ref, got = (line.split(SEP) + ["", "", ""])[:4]
        r = by_id[fid]
        if not r["out"].startswith("="):
            value_ok = None if r["id"] == "F16" else False  # F16 is a text answer
        elif ref in ("", "-"):
            value_ok = None  # no reference value (F18): judged on parts only
        else:
            value_ok = same(got, ref)
        rows.append({"id": fid, "parts": parts, "ref": ref, "got": got, "value_ok": value_ok, "secs": r["secs"], "out": r["out"]})
    return rows


summary = []
for f in result_files:
    model = os.path.basename(f)[8:-5]
    rows = grade(f)
    print(f"\n=== {model}")
    for r in rows:
        r["v"] = verdict(r)
        print(f"  {r['id']} {r['v']} {r['parts']:<6} ref={r['ref'][:12]:<12} got={r['got'][:12]:<12} {r['out'][:60]}")
    secs = sorted(r["secs"] for r in rows)
    summary.append((model, sum(r["v"] == "✓" for r in rows), sum(r["v"] == "~" for r in rows), len(rows), secs[len(secs) // 2]))

print("\n✓ = right value AND required parts present   ~ = right value, parts differ (look at it)   ✗ = wrong")
print(f"{'model':<22}{'✓ correct':>11}{'~ check':>9}{'median s':>10}")
for m, ok, chk, n, s in summary:
    print(f"{m:<22}{f'{ok}/{n}':>11}{chk:>9}{s:>10.1f}")
