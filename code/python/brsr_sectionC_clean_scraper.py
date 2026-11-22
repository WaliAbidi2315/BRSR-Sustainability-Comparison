"""
BRSR Section C scraper (clean, readable output) - Python version

Reads a BRSR / annual-report PDF, isolates "SECTION C: PRINCIPLE WISE
PERFORMANCE DISCLOSURE" and writes an Excel file with 5 sheets:

  Read Me         - what each column means
  Section C Data  - one row per number:
                    Principle | Q No | Indicator | Essential/Leadership |
                    Item | Period | Column | Col No | Value | Unit | As Reported
  Wide Tables     - big multi-column tables kept as one row per table row
  Qualitative     - the written answer for every indicator
  Raw Text        - cleaned Section C text, for checking

Usage:
    pip install openpyxl pandas pdfplumber      # pdfplumber only needed if pdftotext is missing
    python brsr_sectionC_clean_scraper.py Bajaj_Finserv.pdf
    python brsr_sectionC_clean_scraper.py Bajaj_Finserv.pdf out.xlsx "BAJAJ FINSERV LIMITED"

Needs the poppler `pdftotext` command (preferred) because column positions
matter for assigning numbers to the right year/column.
"""
import re
import shutil
import subprocess
import sys

import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ------------------------------------------------------------------ settings
PDF_PATH = sys.argv[1] if len(sys.argv) > 1 else "Bajaj_Finserv.pdf"
OUT_PATH = sys.argv[2] if len(sys.argv) > 2 else "Bajaj_Finserv_BRSR_SectionC_clean.xlsx"
COMPANY_FOOTER = sys.argv[3] if len(sys.argv) > 3 else "BAJAJ FINSERV LIMITED"  # running page header

# ------------------------------------------------------------------ regexes
FOOT_RE = re.compile(
    r"(^\s*\d{1,3}\s+\d{1,2}(st|nd|rd|th)\s+Annual Report)|"
    r"(Business Responsibility and Sustainability Report\s+\d+)|"
    + "(" + re.escape(COMPANY_FOOTER) + ")|"
    r"(^\s*Section A\s+Section B\s+Section C)|"
    r"(^\s*Corporate\s+Statutory\s+Financial)|"
    r"(^\s*Overview\s+Reports\s+Statements)|"
    r"(^\s*\d{1,3}\s*$)|(^\s*Section [ABC]\s*$)|"
    r"(^\s*(Financial\s+)?Statements\s*$)|(^\s*(Statutory\s+)?Reports\s*$)|"
    r"(^\s*Corporate\s*$)|(^\s*Overview\s*$)",
    re.I,
)
# a table cell is a VALUE if it is a number, a %, a number + unit word, or NIL / N.A. / dash
VAL_RE = re.compile(
    r"^(-?[\d,]+(\.\d+)?%?|"
    r"-?[\d,]+(\.\d+)?\s*(lakh|crore|million|MW|KW|KWh)|"
    r"NIL|Nil|N\.A\.?|NA|-)$"
)
NUM_WITH_UNIT_RE = re.compile(r"^(-?[\d,]+(?:\.\d+)?)\s*(lakh|crore|million|MW|KW|KWh)$")
PLAIN_NUM_RE = re.compile(r"^-?[\d,]+(\.\d+)?%?$")
NILISH_RE = re.compile(r"^(NIL|Nil|N\.A\.?|NA|-)$")
FY_RE = re.compile(r"^FY\s?\d{4}", re.I)
NARR_NUM_RE = re.compile(
    r"\d[\d,]*(?:\.\d+)?\s*(?:Metric Tonnes|metric tonnes|million units|"
    r"million hours|million|MW|KL|crore|lakh|per cent|%)"
)

UNIT_RULES = [
    ("Gigajoules", r"gigajoule"),
    ("Kilolitres", r"kilolitre|\bKL\b"),
    ("Metric tonnes CO2e", r"CO2"),
    ("Metric tonnes", r"metric tonne|\bMT\b"),
    ("INR crore", r"crore|\bCr\b"),
    ("INR lakh", r"lakh"),
    ("Days", r"days of"),
]
UNIT_WORD_MAP = {"lakh": "INR lakh", "crore": "INR crore", "million": "Million",
                 "mw": "MW", "kw": "KW", "kwh": "KWh"}


# ------------------------------------------------------------------ helpers
def get_lines(path):
    if shutil.which("pdftotext"):
        out = subprocess.run(["pdftotext", "-layout", path, "-"],
                             capture_output=True, text=True, errors="replace").stdout
        return out.split("\n")
    try:
        import pdfplumber
    except ImportError:
        sys.exit("Install poppler (pdftotext) or `pip install pdfplumber`.")
    lines = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            lines.extend((page.extract_text(layout=True) or "").split("\n"))
    return lines


def clean_line(l):
    """Blank out the P1..P9 margin tab but keep character positions."""
    l = l.replace("\x07", " ").replace("\t", " ").replace("\x0c", "")
    l = re.sub(r"^P[1-9](?=\s|$)", lambda m: " " * len(m.group(0)), l)
    return l.rstrip()


def get_cells(l):
    """Split a line into cells separated by 2+ spaces, remembering positions."""
    return [(m.start(), m.end(), m.group(0)) for m in re.finditer(r"\S+(?: \S+)*", l)]


def parse_val(txt):
    m = NUM_WITH_UNIT_RE.match(txt)
    if m:
        return float(m.group(1).replace(",", "")), m.group(2)
    if PLAIN_NUM_RE.match(txt):
        try:
            return float(txt.replace(",", "").replace("%", "")), ""
        except ValueError:
            return None, ""
    return None, ""


def infer_unit(txt, unit_word, ctx):
    if txt.endswith("%"):
        return "%"
    if unit_word:
        return UNIT_WORD_MAP.get(unit_word.lower(), unit_word)
    if NILISH_RE.match(txt):
        return ""
    for name, pat in UNIT_RULES:
        if re.search(pat, ctx, re.I):
            return name
    return "Number"


# ------------------------------------------------------------------ 1. isolate Section C
lines = get_lines(PDF_PATH)
s = next((i for i, l in enumerate(lines)
          if re.search(r"SECTION C\s*:\s*PRINCIPLE WISE", l, re.I)), None)
if s is None:
    sys.exit("Could not find the 'SECTION C' header in the PDF text.")
e = next((i for i in range(s + 1, len(lines))
          if re.search(r"Independent Assurance Statement", lines[i], re.I)), len(lines))
sec = [clean_line(l) for l in lines[s:e]]

# ------------------------------------------------------------------ 2. walk through Section C
long_rows, wide_rows = [], []
qmeta, qtext, narr, qorder = {}, {}, {}, []
princ = kind = qkey = None
hdr, seen = [], False

for l in sec:
    if not l.strip() or FOOT_RE.search(l):
        continue

    mp = re.search(r"PRINCIPLE\s*(\d)\s*:", l, re.I)
    if mp:
        princ, kind, qkey = "P" + mp.group(1), None, None
        continue
    if re.match(r"^\s*Essential Indicators", l, re.I):
        kind, qkey = "Essential", None
        continue
    if re.match(r"^\s*Leadership Indicators", l, re.I):
        kind, qkey = "Leadership", None
        continue
    if princ is None or kind is None:
        continue

    # new numbered question
    mq = re.match(r"^\s{0,8}(\d{1,2})\.\s+([^\d].*)$", l)
    if mq:
        qn = int(mq.group(1))
        qkey = (princ, kind, qn)
        qmeta[qkey] = {"princ": princ, "kind": kind, "qn": qn}
        qtext[qkey] = mq.group(2).strip()
        narr[qkey] = []
        qorder.append(qkey)
        hdr, seen = [], False
        continue
    if qkey is None:
        continue

    cs = get_cells(l)
    if not cs:
        continue
    isv = [bool(VAL_RE.match(c[2])) for c in cs]

    # no numbers (or a lone cell): header row or plain text
    if not any(isv) or len(cs) < 2:
        if not any(isv) and len(cs) >= 2:
            if seen:                      # a new table starts
                hdr, seen = [], False
            hdr.append(cs)
            continue
        t = l.strip()
        if not narr[qkey] and len(qtext[qkey]) < 400 and not re.search(r"[?:.)]$", qtext[qkey]):
            qtext[qkey] += " " + t        # wrapped question text
        else:
            narr[qkey].append(t)          # answer text
        continue

    # data row: label + values
    seen = True
    vals = [c for c, v in zip(cs, isv) if v]
    first = vals[0][0]
    left = [c for c, v in zip(cs, isv) if (not v) and c[1] <= first]
    # first cell is the row label (whatever its length); extra left cells only if short
    left = left[:1] + [c for c in left[1:] if len(c[2]) <= 45]
    label = " ".join(c[2] for c in left)[:90]
    n = len(vals)

    entries = []
    for k, c in enumerate(vals, 1):
        ctr = (c[0] + c[1]) / 2
        per, parts = "", []
        for hl in hdr:
            d = [abs((h[0] + h[1]) / 2 - ctr) for h in hl]
            j = d.index(min(d))
            h = hl[j]
            if d[j] <= 13 or (h[0] - 6 <= ctr <= h[1] + 6):
                if FY_RE.match(h[2]):
                    per = h[2]
                elif len(h[2]) <= 30:
                    parts.append(h[2])
        num, uw = parse_val(c[2])
        entries.append({
            "Principle": princ, "QN": qmeta[qkey]["qn"], "EL": kind,
            "Item": label, "Period": per, "Column": " ".join(parts),
            "ColNo": k, "Value": num,
            "Unit": infer_unit(c[2], uw, l + " " + qtext[qkey]),
            "Reported": c[2], "QKey": qkey,
        })

    any_period = any(x["Period"] for x in entries)
    if n <= 5 or any_period:
        long_rows.extend(entries)
    else:
        heads = []
        for hl in hdr:
            for h in sorted(hl):
                if h[2] not in heads and len(h[2]) <= 30:
                    heads.append(h[2])
        wide_rows.append({
            "Principle": princ, "QN": qmeta[qkey]["qn"], "EL": kind, "Item": label,
            "Headers": " | ".join(heads), "Values": " | ".join(c[2] for c in vals),
            "QKey": qkey,
        })

# ------------------------------------------------------------------ 3. numbers inside written answers
for qk in qorder:
    txt = " ".join(narr[qk])
    for m in NARR_NUM_RE.finditer(txt):
        hit = m.group(0)
        num = float(re.sub(r"\s*[A-Za-z% ].*$", "", hit).replace(",", ""))
        u = re.sub(r"^\d[\d,]*(?:\.\d+)?\s*", "", hit).strip().lower()
        u = {"mw": "MW", "kl": "KL", "%": "%"}.get(u, u[:1].upper() + u[1:])
        snip = txt[max(0, m.start() - 70):m.start()].strip()
        long_rows.append({
            "Principle": qmeta[qk]["princ"], "QN": qmeta[qk]["qn"], "EL": qmeta[qk]["kind"],
            "Item": "(from written answer) ..." + snip, "Period": "", "Column": "",
            "ColNo": 1, "Value": num, "Unit": u, "Reported": hit, "QKey": qk,
        })


# ------------------------------------------------------------------ 4. build output tables
def ind_label(qk, width=220):
    t = qtext[qk]
    if len(t) > width:
        t = t[:width] + "..."
    return f"Q{qmeta[qk]['qn']}. {t}"


long_df = pd.DataFrame([{
    "Principle": x["Principle"], "Q No": x["QN"], "Indicator": ind_label(x["QKey"]),
    "Essential/Leadership": x["EL"], "Item": x["Item"], "Period": x["Period"],
    "Column": x["Column"], "Col No": x["ColNo"], "Value": x["Value"],
    "Unit": x["Unit"], "As Reported": x["Reported"],
} for x in long_rows])
# document order is already P1..P9, Essential -> Leadership; enforce it (stable sort)
long_df["_p"] = long_df["Principle"].str[1:].astype(int)
long_df["_l"] = (long_df["Essential/Leadership"] != "Essential").astype(int)
long_df = (long_df.sort_values(["_p", "_l", "Q No"], kind="stable")
           .drop(columns=["_p", "_l"]).reset_index(drop=True))

if wide_rows:
    wide_df = pd.DataFrame([{
        "Principle": x["Principle"], "Q No": x["QN"], "Indicator": ind_label(x["QKey"]),
        "Essential/Leadership": x["EL"], "Table Row": x["Item"],
        "Column Headings (approx.)": x["Headers"], "Values (in table order)": x["Values"],
    } for x in wide_rows])
else:
    wide_df = pd.DataFrame({"Note": ["No wide tables found"]})

qual_rows = []
for qk in qorder:
    ans = " ".join(narr[qk])
    if not ans:
        ans = "(table / numbers only - see 'Section C Data' and 'Wide Tables')"
    if len(ans) > 2500:
        ans = ans[:2500] + "..."
    qual_rows.append({
        "Principle": qmeta[qk]["princ"], "Q No": qmeta[qk]["qn"],
        "Essential/Leadership": qmeta[qk]["kind"],
        "Indicator": f"Q{qmeta[qk]['qn']}. {qtext[qk]}", "Response": ans,
    })
qual_df = pd.DataFrame(qual_rows)

readme = pd.DataFrame({
    "Column": ["Principle", "Q No", "Indicator", "Essential/Leadership", "Item", "Period",
               "Column", "Col No", "Value", "Unit", "As Reported"],
    "Meaning": [
        "NGRBC principle P1-P9 the indicator belongs to",
        "Question number within that principle's Essential or Leadership list",
        "The BRSR question being answered (shortened)",
        "Essential = mandatory indicator; Leadership = voluntary indicator",
        "Row label of the table the number came from (e.g. Male, Total, a metric name)",
        "Financial year column the number sits under (FY2024, FY2023, ...), if the table has one",
        "Other column heading(s) detected above the number (best effort, from page layout)",
        "Position of the number within its table row (1 = first number)",
        "The number as a numeric value (blank for NIL / N.A.)",
        "Unit inferred from the number (%, INR crore, Gigajoules, ...) or 'Number' for counts",
        "Exactly what the report printed (keeps NIL, N.A., 8.50 lakh etc.)",
    ],
})
raw_df = pd.DataFrame({"Cleaned Section C text": sec})

# ------------------------------------------------------------------ 5. write Excel
HEAD_FILL = PatternFill("solid", fgColor="1F4E78")
HEAD_FONT = Font(bold=True, color="FFFFFF")
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def write_sheet(writer, name, df, widths):
    df.to_excel(writer, sheet_name=name, index=False)
    ws = writer.sheets[name]
    for j in range(1, df.shape[1] + 1):
        c = ws.cell(row=1, column=j)
        c.fill, c.font, c.border = HEAD_FILL, HEAD_FONT, BORDER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        w = widths[(j - 1) % len(widths)]
        ws.column_dimensions[get_column_letter(j)].width = w
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=df.shape[1]):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
            c.border = BORDER
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


with pd.ExcelWriter(OUT_PATH, engine="openpyxl") as writer:
    write_sheet(writer, "Read Me", readme, [24, 100])
    write_sheet(writer, "Section C Data", long_df, [10, 7, 55, 14, 34, 10, 26, 7, 12, 18, 14])
    write_sheet(writer, "Wide Tables", wide_df, [10, 7, 50, 14, 30, 40, 60])
    write_sheet(writer, "Qualitative", qual_df, [10, 7, 14, 60, 110])
    write_sheet(writer, "Raw Text", raw_df, [150])

print("Done:", OUT_PATH)
print("  Section C Data rows:", len(long_df))
print("  Wide table rows    :", 0 if "Note" in wide_df.columns else len(wide_df))
print("  Indicators found   :", len(qorder))
