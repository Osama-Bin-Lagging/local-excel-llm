"""Score LOKUP.V answers on the real assessments, one cell at a time, in real Excel.

For each task: copy of the answer key → put the model's formula in the target cell → calculate →
read the value → put the key formula back. Isolated, so one wrong cell can't fail the cells after it.
Usage: python3 -I grade_real.py tasks.json keys/ results-a.json [results-b.json ...]
"""
import json, os, re, shutil, subprocess, sys, tempfile

TASKS, KEYDIR, result_files = sys.argv[1], sys.argv[2], sys.argv[3:]
tasks = {t["id"]: t for t in json.load(open(TASKS))}
TESTS = sorted({t["test"] for t in tasks.values()})
SEP = "\x1f"
SHEETS = {"IB": {"Read Me", "Operating Data", "Trading Comps", "Deal Model", "Summary", "Answer Key"},
          "DA": {"Read Me", "Raw Data", "Analysis", "Answer Key"},
          "V4": {"SaaS Metrics", "PnL", "Loan Calc", "Budget", "Stock Data", "Inventory", "Project", "Scores"},
          "V5": {"Payroll", "Campaigns", "Shipments"}}


def q(s):
    return '"' + s.replace("\\", "\\\\").replace('"', '\\"') + '"'


def num(x):
    try:
        return float(str(x).replace(",", ""))
    except ValueError:
        return None


def same(got, ref):
    a, b = num(got), num(ref)
    if a is not None and b is not None:
        return abs(a - b) <= 1e-6 * max(1.0, abs(b))
    return str(got).strip() == str(ref).strip()


def col_num(c):
    n = 0
    for ch in c:
        n = n * 26 + ord(ch) - 64
    return n


def refers_to_itself(f, sheet, cell):
    """True if formula f (typed into sheet!cell) references a range containing that cell → circular."""
    m0 = re.match(r"([A-Z]+)(\d+)", cell)
    tc, tr = col_num(m0[1]), int(m0[2])
    masked = re.sub(r'"(?:[^"]|"")*"', '""', f)
    for m in re.finditer(r"(?:'([^']+)'!|([A-Za-z_][\w.]*)!)?\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?", masked):
        sh = m[1] or m[2] or sheet
        if sh != sheet:
            continue
        c1, r1 = col_num(m[3]), int(m[4])
        c2, r2 = (col_num(m[5]), int(m[6])) if m[5] else (c1, r1)
        if min(c1, c2) <= tc <= max(c1, c2) and min(r1, r2) <= tr <= max(r1, r2):
            return True
    return False


def evaluate(test, rows):
    name = f"grade_{os.getpid()}_{test}.xlsx"
    path = os.path.join(tempfile.mkdtemp(), name)
    shutil.copy(os.path.join(KEYDIR, f"{test}_key.xlsx"), path)
    L = ["with timeout of 300 seconds", 'tell application "Microsoft Excel"', f"open POSIX file {q(path)}",
         f"repeat 20 times\ntry\nset wb to workbook {q(name)}\nexit repeat\non error\ndelay 0.5\nend try\nend repeat",
         'set acc to ""']
    for r in rows:
        t = tasks[r["id"]]
        cell = f'range "{t["cell"]}" of worksheet {q(t["sheet"])} of wb'
        L.append(
            f"set keyf to formula of {cell}\n"
            f"try\nset formula of {cell} to {q(r['out'])}\ncalculate\nset v to (value of {cell}) as text\n"
            f"on error\nset v to \"#ERROR\"\nend try\n"
            f"set acc to acc & {q(r['id'])} & {q(SEP)} & v & linefeed\n"
            f"set formula of {cell} to keyf")
    L += ["close wb saving no", "return acc", "end tell", "end timeout"]
    out = subprocess.run(["osascript", "-"], input="\n".join(L), capture_output=True, text=True)
    shutil.rmtree(os.path.dirname(path), ignore_errors=True)
    if out.returncode:
        sys.exit(out.stderr)
    return dict(line.split(SEP, 1) for line in out.stdout.splitlines() if SEP in line)


summary = []
for f in result_files:
    model = os.path.basename(f)[8:-5]
    res = json.load(open(f))
    verdicts = {}
    for test in TESTS:
        rows, skipped = [], []
        for r in res:
            t = tasks.get(r["id"])
            if not t or t["test"] != test:
                continue
            out = " ".join(r["out"].split("\n"))[:250]
            bad = {m.strip("'") for m in re.findall(r"('[^']+'|[A-Za-z_][\w.]*)!", out)} - SHEETS[test]
            if not out.startswith("="):
                verdicts[r["id"]] = ("✗", "not a formula", out)
            elif bad:  # unknown sheet → Excel would open a blocking file picker
                verdicts[r["id"]] = ("✗", f"unknown sheet {sorted(bad)}", out)
            elif refers_to_itself(out, t["sheet"], t["cell"]):  # Excel would show a circular-reference dialog
                verdicts[r["id"]] = ("✗", "circular ref", out)
            elif not re.search(r"(?:'[^']+'!|[A-Za-z_][\w.]*!)?\$?[A-Z]{1,3}\$?\d+", re.sub(r'"(?:[^"]|"")*"', '""', out)):
                verdicts[r["id"]] = ("✗", "hard-coded", out)  # right number, but typed in: won't update with the data
            else:
                rows.append({**r, "out": out})
        got = evaluate(test, rows) if rows else {}
        for r in rows:
            v = got.get(r["id"], "#ERROR")
            ok = same(v, tasks[r["id"]]["ref"])
            verdicts[r["id"]] = ("✓" if ok else "✗", v, r["out"])
    print(f"\n=== {model}")
    for tid in tasks:
        if tid in verdicts:
            mark, v, out = verdicts[tid]
            print(f"  {mark} {tid:<14} ref={str(tasks[tid]['ref'])[:12]:<12} got={str(v)[:12]:<12} {out[:70]}")
    for test in TESTS:
        ids = [i for i in verdicts if tasks[i]["test"] == test]
        summary.append((model, test, sum(verdicts[i][0] == "✓" for i in ids), len(ids)))
    secs = sorted(r["secs"] for r in res)
    summary.append((model, "median s", secs[len(secs) // 2], None))

print(f"\n{'model':<30}" + "".join(f"{t:>8}" for t in TESTS) + f"{'median s':>10}")
by = {}
for m, k, a, b in summary:
    by.setdefault(m, {})[k] = (a, b)
for m, d in by.items():
    print(f"{m:<30}" + "".join(f"{f'{d[t][0]}/{d[t][1]}':>8}" for t in TESTS) + f"{d['median s'][0]:>10.1f}")
