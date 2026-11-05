"""Core logic: read M-Pesa branch statements, drop duplicates, build one workbook."""
import csv
import hashlib
import io
import re
import zipfile
from collections import Counter
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

ALLOWED = {".xlsx", ".xlsm", ".xls", ".csv"}
MAX_UNZIPPED = 300 * 1024 * 1024
NCOLS = 14
WIDTHS = [15.8, 33.6, 18.1, 60, 18.1, 12, 11.8, 12, 18.7, 47.6, 42.7, 21, 7.7, 9.1]
ARIAL = Font(name="Arial", size=10)
ARIAL_B = Font(name="Arial", size=10, bold=True)
GREY = PatternFill("solid", fgColor="C0C0C0")
DT = "%d-%m-%Y %H:%M:%S"


@dataclass
class Statement:
    file: str
    short_code: str
    holder: str
    period: tuple
    report_date: object
    rows: list
    hdr: int
    sheet_name: str = ""
    is_new: bool = False
    check: str = "n/a"

    @property
    def n_txn(self):
        return len(self.rows) - self.hdr - 1


def _num(v):
    if isinstance(v, (int, float)):
        return v
    try:
        return float(str(v).replace(",", "").strip())
    except ValueError:
        return None


def _read_rows(name, data):
    ext = Path(name).suffix.lower()
    if ext in (".xlsx", ".xlsm"):
        ws = load_workbook(io.BytesIO(data), data_only=True, read_only=True).worksheets[0]
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
    elif ext == ".xls":
        import xlrd
        sh = xlrd.open_workbook(file_contents=data).sheet_by_index(0)
        rows = [sh.row_values(i) for i in range(sh.nrows)]
    else:
        text = data.decode("utf-8-sig", errors="replace")
        rows = [list(r) for r in csv.reader(io.StringIO(text))]
    rows = [(r + [None] * NCOLS)[:NCOLS] for r in rows]
    rows = [[None if (c is None or c == "") else c for c in r] for r in rows]
    while rows and all(c is None for c in rows[-1]):
        rows.pop()
    return rows, ext == ".csv"


def _labels(rows):
    """{'Short Code:': value, ...} for every 'Label:' cell in the header block."""
    out = {}
    for r in rows:
        for i, c in enumerate(r[:-1]):
            if isinstance(c, str) and c.strip().endswith(":"):
                out.setdefault(c.strip(), r[i + 1])
    return out


def parse(name, data):
    rows, is_csv = _read_rows(name, data)
    hdr = next((i for i, r in enumerate(rows)
                if isinstance(r[0], str) and r[0].strip() == "Receipt No."), None)
    if hdr is None:
        raise ValueError("no 'Receipt No.' header row, so it doesn't look like an M-Pesa statement")
    meta = _labels(rows[:hdr])
    code = meta.get("Short Code:")
    if code is None:
        raise ValueError("no Short Code found")
    code = str(int(code)) if isinstance(code, float) and code.is_integer() else str(code).strip()
    if is_csv:  # csv has no types: restore numbers like the original sheets have
        for r in rows[:hdr]:
            for i in range(1, NCOLS):
                if isinstance(r[i - 1], str) and r[i - 1].endswith(":") and "Code" not in r[i - 1]:
                    if _num(r[i]) is not None:
                        r[i] = _num(r[i])
        for r in rows[hdr + 1:]:
            for i in (5, 6, 7):
                if r[i] is not None and _num(r[i]) is not None:
                    r[i] = _num(r[i])
    tp = next((r for r in rows[:hdr] if r[0] == "Time Period:"), [None] * NCOLS)
    try:
        rep = datetime.strptime(str(meta.get("Date of Report:")), DT)
    except ValueError:
        rep = None
    st = Statement(name, code, str(meta.get("Account Holder:") or "").strip(),
                   (str(tp[2]), str(tp[4])), rep, rows, hdr)
    # control: header totals vs sum of the transaction columns
    pi, wd = _num(meta.get("Total Paid In:")), _num(meta.get("Total Withdrawn:"))
    if pi is not None and wd is not None:
        sp = sum(_num(r[5]) or 0 for r in rows[hdr + 1:])
        sw = sum(_num(r[6]) or 0 for r in rows[hdr + 1:])
        st.check = "OK" if abs(sp - pi) < 0.01 and abs(sw - wd) < 0.01 else "Totals don't match"
    return st


def suggest_name(holder, code):
    n = re.sub(r"^\W*new\s+muthokinju\s*(hardware)?\s*(ltd)?\W*", "", holder, flags=re.I)
    n = re.sub(r"\s+branch\s*$", "", n, flags=re.I)
    n = re.sub(r"[\[\]:*?/\\.]", "", n).strip(" -").upper()
    return (n or code)[:31]


def load_branches(path):
    with open(path, newline="", encoding="utf-8") as f:
        return {r["short_code"].strip(): r["sheet_name"].strip() for r in csv.DictReader(f)}


def mapping_csv(statements, mapping):
    merged = dict(mapping)
    merged.update({s.short_code: s.sheet_name for s in statements})
    out = io.StringIO()
    w = csv.writer(out)
    w.writerow(["short_code", "sheet_name"])
    for k in sorted(merged):
        w.writerow([k, merged[k]])
    return out.getvalue()


def process_zip(data, mapping):
    kept, skipped, rejected, ignored, seen = {}, [], [], [], {}
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        infos = [i for i in z.infolist() if not i.is_dir()]
        if sum(i.file_size for i in infos) > MAX_UNZIPPED:
            raise ValueError("zip is too large")
        for info in sorted(infos, key=lambda i: i.filename):
            name, base = info.filename, Path(info.filename).name
            if "__MACOSX" in name or base.startswith((".", "~$")):
                continue
            if Path(base).suffix.lower() not in ALLOWED:
                ignored.append(name)
                continue
            raw = z.read(info)
            digest = hashlib.md5(raw).hexdigest()
            if digest in seen:
                skipped.append((name, f"Identical file to {seen[digest]}"))
                continue
            seen[digest] = name
            try:
                st = parse(name, raw)
            except Exception as e:
                rejected.append((name, str(e)))
                continue
            key = (st.short_code, st.period)
            old = kept.get(key)
            if old is None:
                kept[key] = st
                continue
            rank = lambda s: (s.report_date or datetime.min, s.n_txn)
            win, lose = (st, old) if rank(st) > rank(old) else (old, st)
            kept[key] = win
            skipped.append((lose.file, f"Same branch and period as {win.file} (kept the newer/larger one)"))
    sts = sorted(kept.values(), key=lambda s: (s.short_code, s.period))
    for s in sts:
        if s.short_code in mapping:
            s.sheet_name = mapping[s.short_code]
        else:
            s.sheet_name, s.is_new = suggest_name(s.holder, s.short_code), True
    return {"statements": sts, "skipped": skipped, "rejected": rejected, "ignored": ignored}


def output_name(sts):
    try:
        d = datetime.strptime(Counter(s.period[0] for s in sts).most_common(1)[0][0], DT)
        suf = "TH" if 10 <= d.day % 100 <= 20 else {1: "ST", 2: "ND", 3: "RD"}.get(d.day % 10, "TH")
        return f"{d.day}{suf}_{d.strftime('%b').upper()}_{d.year}_MPESA.xlsx"
    except Exception:
        return "MPESA_CONSOLIDATED.xlsx"


def build_workbook(sts, blank_first=True):
    wb = Workbook()
    wb.active.title = "Sheet1"
    if not blank_first:
        wb.remove(wb.active)
    used = {"sheet1"} if blank_first else set()
    for s in sts:
        base = re.sub(r"[\[\]:*?/\\]", "", s.sheet_name).strip()[:31] or s.short_code
        name, n = base, 2
        while name.lower() in used:
            name = f"{base[:28]} {n}"
            n += 1
        used.add(name.lower())
        ws = wb.create_sheet(name)
        for ri, row in enumerate(s.rows):
            for ci, v in enumerate(row):
                c = ws.cell(ri + 1, ci + 1, v)
                c.font = ARIAL
                if ri == s.hdr:
                    c.font, c.fill = ARIAL_B, GREY
                    c.alignment = Alignment(horizontal="left")
                elif isinstance(v, (int, float)) and not isinstance(v, bool):
                    c.number_format = "#,##0.00"
                    c.alignment = Alignment(horizontal="right")
                else:
                    c.alignment = Alignment(horizontal="left")
        for i, w in enumerate(WIDTHS):
            ws.column_dimensions[chr(65 + i)].width = w
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
