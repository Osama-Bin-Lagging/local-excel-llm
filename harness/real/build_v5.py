"""Held-out v5: a fresh business workbook (not derived from any test file) to check that the
glossary generalises. Builds the answer-key workbook, has Excel compute the references, writes tasks.

Usage: python3 -I build_v4.py <out tasks_v4.json> <keys dir>
"""
import json, os, shutil, subprocess, sys, tempfile
from openpyxl import Workbook
from openpyxl.utils import range_boundaries, get_column_letter

OUT, KEYDIR = sys.argv[1], sys.argv[2]
SEP = "\x1f"
wb = Workbook()

# ---- Payroll (title row 1, header row 2, data rows 3-10)
ws = wb.active
ws.title = "Payroll"
ws["A1"] = "March payroll (INR)"
ws.append(["Employee", "Department", "Base salary", "Hours worked", "Overtime hours", "Overtime rate per hour",
           "Bonus %", "Gross pay", "Tax deducted", "Net pay", "Overtime share of gross"])
rows = [("Anil", "Ops", 52000, 176, 12, 450, 0.05), ("Bina", "Sales", 61000, 168, 0, 0, 0.12), ("Charu", "Ops", 48000, 180, 20, 400, 0),
        ("Deepak", "IT", 90000, 160, 6, 900, 0.08), ("Esha", "Sales", 58000, 172, 4, 500, 0.10), ("Farid", "HR", 45000, 165, 0, 0, 0.03),
        ("Gita", "IT", 82000, 170, 10, 850, 0.06), ("Hari", "Ops", 50000, 184, 24, 420, 0.02)]
for i, (n, d, b, h, o, r, bp) in enumerate(rows, 3):
    ws.append([n, d, b, h, o, r, bp, f"=C{i}*(1+G{i})+E{i}*F{i}", f"=H{i}*0.1", f"=H{i}-I{i}", f"=E{i}*F{i}/H{i}"])
ws["M2"], ws["N2"] = "Department", "Total gross pay"
for j, dep in enumerate(["Ops", "Sales", "IT", "HR"], 3):
    ws[f"M{j}"], ws[f"N{j}"] = dep, f"=SUMIFS($H$3:$H$10,$B$3:$B$10,M{j})"

# ---- Campaigns (header row 1, data rows 2-7)
ws = wb.create_sheet("Campaigns")
ws.append(["Campaign", "Channel", "Spend (INR)", "Impressions", "Clicks", "Orders", "Revenue (INR)", "CTR", "CPC", "ROAS", "Cost per order"])
for i, row in enumerate([("Diwali Sale", "Search", 250000, 1200000, 36000, 900, 1350000), ("New Year", "Social", 180000, 2100000, 25200, 420, 630000),
                         ("Summer Drop", "Search", 120000, 640000, 15360, 300, 510000), ("Brand Week", "Display", 90000, 3000000, 9000, 60, 96000),
                         ("Clearance", "Social", 60000, 900000, 13500, 270, 405000), ("Loyalty", "Email", 15000, 80000, 4000, 160, 280000)], 2):
    ws.append(list(row) + [f"=E{i}/D{i}", f"=C{i}/E{i}", f"=G{i}/C{i}", f"=C{i}/F{i}"])
ws["A9"], ws["B9"] = "Total spend", "=SUM(C2:C7)"
ws["A10"], ws["B10"] = "Blended ROAS", "=SUM(G2:G7)/SUM(C2:C7)"
ws["A11"], ws["B11"] = "Campaigns with ROAS above 4", '=COUNTIF(J2:J7,">4")'

# ---- Shipments (data table, header row 4)
ws = wb.create_sheet("Shipments")
ws["A1"] = "Shipment log"
ws.append([]); ws.append([])
ws.append(["Shipment ID", "Origin", "Destination", "Ship date", "Delivery date", "Weight (kg)", "Freight cost (INR)", "Status"])
import datetime as dt
ship = [("S-101", "Mumbai", "Delhi", 5, 8, 120, 9600, "Delivered"), ("S-102", "Pune", "Chennai", 6, 11, 80, 7200, "Delivered"),
        ("S-103", "Delhi", "Kolkata", 6, None, 200, 15000, "In transit"), ("S-104", "Mumbai", "Bengaluru", 7, 10, 65, 5850, "Delivered"),
        ("S-105", "Chennai", "Mumbai", 8, 14, 150, 12750, "Delayed"), ("S-106", "Pune", "Delhi", 9, 12, 40, 4000, "Delivered"),
        ("S-107", "Delhi", "Mumbai", 10, None, 95, 8075, "Lost"), ("S-108", "Kolkata", "Pune", 11, 16, 110, 9900, "Delivered")]
for sid, o, d, s, e, w, c, st in ship:
    ws.append([sid, o, d, dt.date(2026, 3, s), dt.date(2026, 3, e) if e else None, w, c, st])
for r in range(5, 13):
    ws[f"D{r}"].number_format = ws[f"E{r}"].number_format = "dd-mmm-yyyy"
ws["J4"], ws["K4"] = "Metric", "Value"
ws["J5"], ws["K5"] = "Delivered share", '=COUNTIF(H5:H12,"Delivered")/COUNTA(A5:A12)'
ws["J6"], ws["K6"] = "Average transit days (delivered)", '=SUMPRODUCT((H5:H12="Delivered")*(E5:E12-D5:D12))/COUNTIF(H5:H12,"Delivered")'
ws["J7"], ws["K7"] = "Freight cost per kg (all)", "=SUM(G5:G12)/SUM(F5:F12)"
ws["J8"], ws["K8"] = "Shipments from Mumbai", '=COUNTIF(B5:B12,"Mumbai")'
ws["J9"], ws["K9"] = "Heaviest shipment ID", "=INDEX(A5:A12,MATCH(MAX(F5:F12),F5:F12,0))"
ws["L5"] = "(per-row) Transit days"
ws["L6"] = "=E6-D6"

key = os.path.join(KEYDIR, "V5_key.xlsx")
wb.save(key)

PR, CA, SH = "'Payroll'!A2:N10", "'Campaigns'!A1:K11", "'Shipments'!A4:L12"
tasks = [
    ("Payroll", "H5", "Gross pay for Charu (row 5): base salary plus bonus % of base, plus overtime hours × overtime rate.", PR),
    ("Payroll", "J7", "Net pay for Esha (row 7): gross pay minus tax deducted.", PR),
    ("Payroll", "I4", "Tax deducted for Bina (row 4): 10% of gross pay.", PR),
    ("Payroll", "K10", "Overtime share of gross pay for Hari (row 10).", PR),
    ("Payroll", "N5", "Total gross pay for the Sales department (name in M5).", PR),
    ("Campaigns", "H3", "CTR for the New Year campaign (row 3).", CA),
    ("Campaigns", "I4", "CPC for Summer Drop (row 4).", CA),
    ("Campaigns", "J2", "ROAS for Diwali Sale (row 2).", CA),
    ("Campaigns", "K6", "Cost per order for Clearance (row 6).", CA),
    ("Campaigns", "B10", "Blended ROAS across all campaigns.", CA),
    ("Campaigns", "B11", "How many campaigns have a ROAS above 4?", CA),
    ("Shipments", "K5", "Share of shipments that were delivered.", SH),
    ("Shipments", "K7", "Freight cost per kg across all shipments.", SH),
    ("Shipments", "K8", "How many shipments started from Mumbai?", SH),
    ("Shipments", "K9", "Shipment ID of the heaviest shipment.", SH),
    ("Shipments", "L6", "Transit days for shipment S-102 (row 6): delivery date minus ship date.", SH),
]


def q(s):
    return '"' + s.replace("\\", "\\\\").replace('"', '\\"') + '"'


# Excel computes every cell we need (targets + formula cells inside each context)
from openpyxl import load_workbook
kf = load_workbook(key)
wanted = {}
for sh, cell, _, ctx in tasks:
    csh, rng = ctx.split("!")
    csh = csh.strip("'")
    c1, r1, c2, r2 = range_boundaries(rng)
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            wanted.setdefault(csh, set()).add(f"{get_column_letter(c)}{r}")
name = f"ref_{os.getpid()}_V5.xlsx"
tmp = os.path.join(tempfile.mkdtemp(), name)
shutil.copy(key, tmp)
L_ = ["with timeout of 180 seconds", 'tell application "Microsoft Excel"', f"open POSIX file {q(tmp)}",
      f"repeat 20 times\ntry\nset wb to workbook {q(name)}\nexit repeat\non error\ndelay 0.5\nend try\nend repeat",
      "calculate", 'set acc to ""']
for sh, cells in wanted.items():
    for a in sorted(cells):
        L_.append(f'set acc to acc & {q(sh + "!" + a)} & {q(SEP)} & ((value of range "{a}" of worksheet {q(sh)} of wb) as text) & linefeed')
L_ += ["close wb saving no", "return acc", "end tell", "end timeout"]
out = subprocess.run(["osascript", "-"], input="\n".join(L_), capture_output=True, text=True)
if out.returncode:
    sys.exit(out.stderr)
vals = {}
for line in out.stdout.splitlines():
    if SEP in line:
        k_, v = line.split(SEP, 1)
        vals[tuple(k_.split("!", 1))] = v


def val(v):
    if v in ("", "missing value"):
        return ""
    try:
        return float(v.replace(",", ""))
    except ValueError:
        return v


result = []
for sh, cell, prompt, ctx in tasks:
    csh, rng = ctx.split("!")
    csh = csh.strip("'")
    c1, r1, c2, r2 = range_boundaries(rng)
    grid = [["" if (csh == sh and f"{get_column_letter(c)}{r}" == cell) else val(vals[(csh, f"{get_column_letter(c)}{r}")])
             for c in range(c1, c2 + 1)] for r in range(r1, r2 + 1)]
    result.append({"test": "V5", "sheet": sh, "cell": cell, "prompt": prompt, "contextAddress": ctx, "context": grid,
                   "ref": val(vals[(sh, cell)]), "key_formula": kf[sh][cell].value, "address": f"'{sh}'!{cell}",
                   "id": f"V5-{sh.split()[0][:5]}-{cell}", "fn": "formula"})
json.dump(result, open(OUT, "w"), indent=1)
print(len(result), "tasks")
for t in result:
    print(f"  {t['id']:<16} ref={str(t['ref'])[:14]:<14} key={t['key_formula']}")
