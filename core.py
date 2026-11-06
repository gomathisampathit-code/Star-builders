"""Star Builders - data layer (SQLite), carry-forward logic, reports and Excel import/export."""
import io
import os
import re
import sqlite3
from collections import Counter
from contextlib import closing
from datetime import date, datetime

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

DB_PATH = os.environ.get(
    "STAR_DB", os.path.join(os.path.dirname(os.path.abspath(__file__)), "star_builders.db")
)
HEAD = ["Date", "Paid for", "Description of Expense", "Nos", "Paid To", "Amount (₹)", "Advance", "Balance"]
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

SCHEMA = """
CREATE TABLE IF NOT EXISTS sites(
    id INTEGER PRIMARY KEY,
    name TEXT NOT NULL UNIQUE COLLATE NOCASE
);
CREATE TABLE IF NOT EXISTS entries(
    id INTEGER PRIMARY KEY,
    site_id INTEGER NOT NULL REFERENCES sites(id) ON DELETE CASCADE,
    date TEXT, cat TEXT DEFAULT '', description TEXT DEFAULT '', nos TEXT DEFAULT '',
    paid_to TEXT DEFAULT '', amount REAL DEFAULT 0, advance REAL DEFAULT 0
);
CREATE TABLE IF NOT EXISTS breakups(
    id INTEGER PRIMARY KEY,
    entry_id INTEGER NOT NULL REFERENCES entries(id) ON DELETE CASCADE,
    date TEXT, cat TEXT DEFAULT '', nos TEXT DEFAULT '', paid_to TEXT DEFAULT '', amount REAL DEFAULT 0
);
CREATE INDEX IF NOT EXISTS ix_entries_site ON entries(site_id, date);
"""


# ----------------------------------------------------------------- helpers
def norm_cat(c):
    """Trim; merge every spelling of 'Labour Payment' into one item."""
    c = str(c or "").strip()
    return "Labour Payment" if re.fullmatch(r"labou?rs?\s*payments?", c, re.I) else c


def is_personal(cat):
    return bool(re.search(r"personal\s*advance", str(cat or ""), re.I))


def fd(d):
    try:
        return datetime.strptime(str(d)[:10], "%Y-%m-%d").strftime("%d-%b-%y")
    except Exception:
        return d or ""


def inr(n):
    """Indian digit grouping: 1234567 -> ₹ 12,34,567"""
    try:
        n = float(n or 0)
    except Exception:
        return ""
    neg = n < 0
    n = abs(n)
    s = f"{n:.2f}" if abs(n - round(n)) > 0.004 else f"{int(round(n))}"
    whole, _, frac = s.partition(".")
    if len(whole) > 3:
        head, tail = whole[:-3], whole[-3:]
        parts = []
        while len(head) > 2:
            parts.insert(0, head[-2:])
            head = head[:-2]
        if head:
            parts.insert(0, head)
        whole = ",".join(parts + [tail])
    return ("-" if neg else "") + "₹ " + whole + ("." + frac if frac else "")


def _num(v):
    try:
        return float(str(v).strip().replace(",", ""))
    except Exception:
        return 0.0


def month_label(k):
    return f"{MONTHS[int(k[5:7]) - 1]} {k[:4]}" if re.fullmatch(r"\d{4}-\d{2}", k) else k


# -------------------------------------------------------------- database
def _connect():
    c = sqlite3.connect(DB_PATH)
    c.row_factory = sqlite3.Row
    c.execute("PRAGMA foreign_keys=ON")
    return c


def init_db():
    with closing(_connect()) as c:
        c.executescript(SCHEMA)
        c.commit()


def _all(sql, args=()):
    with closing(_connect()) as c:
        return [dict(r) for r in c.execute(sql, args).fetchall()]


def _run(sql, args=()):
    with closing(_connect()) as c:
        cur = c.execute(sql, args)
        c.commit()
        return cur.lastrowid


def sites():
    return _all("SELECT id, name FROM sites ORDER BY id")


def add_site(name):
    name = name.strip()
    if not name:
        raise ValueError("Please enter a site name")
    try:
        return _run("INSERT INTO sites(name) VALUES(?)", (name,))
    except sqlite3.IntegrityError:
        raise ValueError("A site with this name already exists")


def rename_site(site_id, name):
    name = name.strip()
    if not name:
        raise ValueError("Please enter a site name")
    try:
        _run("UPDATE sites SET name=? WHERE id=?", (name, site_id))
    except sqlite3.IntegrityError:
        raise ValueError("A site with this name already exists")


def delete_site(site_id):
    _run("DELETE FROM sites WHERE id=?", (site_id,))


def distinct(field):
    assert field in ("cat", "paid_to")
    rows = _all(f"SELECT {field} AS v FROM entries UNION SELECT {field} FROM breakups")
    return sorted({r["v"].strip() for r in rows if r["v"] and r["v"].strip()}, key=str.lower)


def load_entries(site_id=None):
    """Entries (oldest first) with their break-up lines under e['br']."""
    where, args = ("WHERE e.site_id=?", (site_id,)) if site_id else ("", ())
    ents = _all(
        "SELECT e.*, s.name AS site FROM entries e JOIN sites s ON s.id=e.site_id "
        + where + " ORDER BY e.date, e.id", args)
    by_id = {e["id"]: e for e in ents}
    for e in ents:
        e["br"] = []
    for b in _all(
        "SELECT b.* FROM breakups b JOIN entries e ON e.id=b.entry_id " + where + " ORDER BY b.date, b.id", args
    ):
        e = by_id[b["entry_id"]]
        b["date"] = b["date"] or e["date"]
        e["br"].append(b)
    return ents


def save_entry(data, brs, entry_id=None):
    cols = ("date", "cat", "description", "nos", "paid_to", "amount", "advance")
    vals = [data[c] for c in cols]
    with closing(_connect()) as c:
        if entry_id:
            c.execute(
                "UPDATE entries SET date=?,cat=?,description=?,nos=?,paid_to=?,amount=?,advance=? WHERE id=?",
                vals + [entry_id])
            c.execute("DELETE FROM breakups WHERE entry_id=?", (entry_id,))
        else:
            entry_id = c.execute(
                "INSERT INTO entries(date,cat,description,nos,paid_to,amount,advance,site_id) VALUES(?,?,?,?,?,?,?,?)",
                vals + [data["site_id"]]).lastrowid
        c.executemany(
            "INSERT INTO breakups(entry_id,date,cat,nos,paid_to,amount) VALUES(?,?,?,?,?,?)",
            [(entry_id, b["date"], norm_cat(b["cat"]), b["nos"], b["paid_to"], b["amount"]) for b in brs])
        c.commit()
    return entry_id


def delete_entry(entry_id):
    _run("DELETE FROM entries WHERE id=?", (entry_id,))


# ------------------------------------------------- advances & carry-forward
def adv_info(entries):
    """Per site advance: previous balance of the same person is carried into the next advance."""
    info, last = {}, {}
    for e in sorted((e for e in entries if e["advance"]), key=lambda e: (e["date"] or "", e["id"])):
        key = (e["site_id"], (e["paid_to"] or "").strip().lower())
        prev = last.get(key)
        carry = prev["bal"] if prev else 0.0
        used = sum(b["amount"] for b in e["br"])
        avail = e["advance"] + carry
        if prev:
            prev["next"] = True
        info[e["id"]] = last[key] = {
            "carry": carry, "has": prev is not None, "avail": avail,
            "used": used, "bal": avail - used, "next": False}
    return info


def totals(entries):
    given = sum(e["advance"] for e in entries)
    used = sum(b["amount"] for e in entries if e["advance"] for b in e["br"])
    spent = sum(e["amount"] for e in entries) + used
    return {"given": given, "used": used, "spent": spent, "balance": given - used}


def flat_rows(entries):
    """One row per payment: normal expenses + break-up lines (+ advance rows with amount 0)."""
    rows = []
    for e in entries:
        base = dict(id=e["id"], site=e["site"])
        if e["advance"]:
            rows.append(dict(base, date=e["date"] or "", cat=e["cat"] or "Site Advance",
                             description=e["description"], nos=e["nos"], paid_to=e["paid_to"],
                             amount=0.0, advance=e["advance"]))
            for b in e["br"]:
                rows.append(dict(base, date=b["date"] or e["date"] or "", cat=b["cat"],
                                 description="From site advance", nos=b["nos"],
                                 paid_to=b["paid_to"] or e["paid_to"], amount=b["amount"], advance=0.0))
        else:
            rows.append(dict(base, date=e["date"] or "", cat=e["cat"], description=e["description"],
                             nos=e["nos"], paid_to=e["paid_to"], amount=e["amount"], advance=0.0))
    return rows


# ----------------------------------------------------------------- reports
def _pct(part, total):
    return round(part / total * 100, 1) if total else 0.0


def _group(rows, key):
    out = {}
    for r in rows:
        k = key(r) or "(not specified)"
        g = out.setdefault(k, [0.0, 0.0])
        g[0] += r["amount"]
        g[1] += _num(r["nos"])
    return out


def build_report(entries, ai, scope_all):
    """entries = the filtered entries; ai = adv_info() over ALL entries (so carry-forward is complete)."""
    rows = flat_rows(entries)
    exp = [r for r in rows if r["amount"] > 0]
    T = totals(entries)
    total = T["spent"]
    tabs = []

    def add(title, cols, data, kinds, open_=False, hl=False, total_row=True):
        tabs.append(dict(title=title, df=pd.DataFrame(data, columns=cols), kinds=kinds,
                         open=open_, hl=hl, total_row=total_row and bool(data)))

    if scope_all:
        data = []
        for n in sorted({e["site"] for e in entries}):
            sub = [e for e in entries if e["site"] == n]
            t = totals(sub)
            data.append([n, len(sub), t["spent"], _pct(t["spent"], total), t["given"], t["balance"]])
        data.sort(key=lambda r: -r[2])
        if data:
            data.append(["All sites", len(entries), total, 100.0 if total else 0.0, T["given"], T["balance"]])
        add("Site-wise expense", ["Site", "Entries", "Spent", "Share %", "Site advance", "Advance balance"], data,
            {"Spent": "money", "Share %": "pct", "Site advance": "money", "Advance balance": "money"}, True)

    g = _group(exp, lambda r: norm_cat(r["cat"]))
    data = [[k, v[1] or None, v[0], _pct(v[0], total)] for k, v in sorted(g.items(), key=lambda x: -x[1][0])]
    if data:
        data.append(["Total", sum(v[1] for v in g.values()) or None, total, 100.0])
    add("By item (Paid for)", ["Item", "Nos", "Amount", "Share %"], data,
        {"Nos": "nos", "Amount": "money", "Share %": "pct"}, True)

    g = _group(exp, lambda r: r["paid_to"])
    data = [[k, v[0], _pct(v[0], total)] for k, v in sorted(g.items(), key=lambda x: -x[1][0])]
    if data:
        data.append(["Total", total, 100.0])
    add("By person (Paid to)", ["Paid to", "Amount", "Share %"], data, {"Amount": "money", "Share %": "pct"})

    g = _group(exp, lambda r: (r["date"] or "")[:7] or "(no date)")
    data = [[month_label(k), v[0], _pct(v[0], total)] for k, v in sorted(g.items())]
    if data:
        data.append(["Total", total, 100.0])
    add("By month", ["Month", "Amount", "Share %"], data, {"Amount": "money", "Share %": "pct"})

    if scope_all:
        names = sorted({r["site"] for r in exp})
        mat = {}
        for r in exp:
            k = norm_cat(r["cat"]) or "(not specified)"
            mat.setdefault(k, {})
            mat[k][r["site"]] = mat[k].get(r["site"], 0.0) + r["amount"]
        data = [[k] + [mat[k].get(n) for n in names] + [sum(mat[k].values())]
                for k in sorted(mat, key=lambda k: -sum(mat[k].values()))]
        if data:
            data.append(["Total"] + [sum(r["amount"] for r in exp if r["site"] == n) for n in names] + [total])
        add("Item x Site", ["Item"] + names + ["Total"], data, {c: "money" for c in names + ["Total"]})

    advs = sorted((e for e in entries if e["advance"]), key=lambda e: (e["date"] or "", e["id"]))
    data = [[e["date"], e["site"], e["paid_to"] or "-", e["advance"],
             ai[e["id"]]["carry"] if ai[e["id"]]["has"] else None,
             ai[e["id"]]["avail"], ai[e["id"]]["used"], ai[e["id"]]["bal"]] for e in advs]
    if data:
        g_, u_ = T["given"], T["used"]
        data.append(["Total", "", "", g_, None, None, u_, g_ - u_])
    add("Site advances", ["Date", "Site", "Person", "Given", "Previous balance", "Total", "Used", "Balance"], data,
        {"Date": "date", "Given": "money", "Previous balance": "money", "Total": "money", "Used": "money",
         "Balance": "money"}, bool(advs))

    data = [[b["date"] or e["date"], e["site"], e["paid_to"] or "-", b["cat"] or "-", b["nos"] or None, b["amount"]]
            for e in advs for b in e["br"]]
    if data:
        data.append(["Total", "", "", "", sum(_num(r[4]) for r in data) or None, sum(r[5] for r in data)])
    add("Advance break-up lines", ["Date", "Site", "Person", "Item", "Nos", "Amount"], data,
        {"Date": "date", "Nos": "nos", "Amount": "money"})

    pers = sorted((r for r in exp if is_personal(r["cat"])), key=lambda r: (r["date"], r["id"]))
    data = [[r["date"], r["site"], r["paid_to"] or "-", r["description"], r["amount"]] for r in pers]
    if data:
        data.append(["Total", "", "", "", sum(r[4] for r in data)])
    add("Personal advances", ["Date", "Site", "Person", "Note", "Amount"], data,
        {"Date": "date", "Amount": "money"}, bool(pers), hl=True)

    top = sorted(exp, key=lambda r: -r["amount"])[:10]
    data = [[r["date"], r["site"], r["cat"] or "-", r["paid_to"] or "-", r["amount"]] for r in top]
    if data:
        data.append(["Total (top 10)", "", "", "", sum(r[4] for r in data)])
    add("Top 10 expenses", ["Date", "Site", "Item", "Paid to", "Amount"], data, {"Date": "date", "Amount": "money"})

    data = [[r["date"], r["site"], r["cat"], r["description"], r["paid_to"], r["nos"] or None,
             r["amount"] or None, r["advance"] or None]
            for r in sorted(rows, key=lambda r: (r["date"], r["id"]))]
    if data:
        data.append(["Total", "", "", "", "", sum(_num(r[5]) for r in data) or None,
                     sum(r[6] or 0 for r in data), sum(r[7] or 0 for r in data)])
    add("Full statement", ["Date", "Site", "Item", "Description", "Paid to", "Nos", "Amount", "Advance"], data,
        {"Date": "date", "Nos": "nos", "Amount": "money", "Advance": "money"})
    return tabs, T


def display_df(t):
    """Text version of a report table (Indian number format) for on-screen display."""
    df = t["df"].copy().astype(object)
    for col, kind in t["kinds"].items():
        def f(v, kind=kind):
            if v is None or (isinstance(v, float) and pd.isna(v)) or v == "":
                return ""
            if kind == "money":
                return inr(v)
            if kind == "pct":
                return f"{v:g}%"
            if kind == "nos":
                return f"{_num(v):g}"
            if kind == "date":
                return fd(v) if re.match(r"\d{4}-", str(v)) else v
            return v
        df[col] = df[col].map(f)
    return df.fillna("")


# ------------------------------------------------------------ Excel export
def _safe_title(name, used):
    t = re.sub(r"[\\/?*\[\]:]", "", name).strip()[:31] or "Site"
    base, i = t, 2
    while t.lower() in used:
        t = f"{base[:28]}_{i}"
        i += 1
    used.add(t.lower())
    return t


def _style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row, c)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="0F6B5C")
        cell.alignment = Alignment(horizontal="center")


def _to_date(v):
    try:
        return datetime.strptime(str(v)[:10], "%Y-%m-%d").date()
    except Exception:
        return None


def export_workbook(entries, site_list):
    """One sheet per site in the same layout as the original P&L Excel, plus a Summary sheet."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    used = set()
    summary = [["Site", "Advance", "Total spent", "Balance"]]
    for s in site_list:
        ents = [e for e in entries if e["site_id"] == s["id"]]
        ws = wb.create_sheet(_safe_title(s["name"], used))
        ws["A1"] = s["name"]
        ws["A1"].font = Font(bold=True, size=14)
        ws.append(HEAD)
        _style_header(ws, 2, len(HEAD))
        bal, nos_total = 0.0, 0.0
        for e in sorted(ents, key=lambda e: (e["date"] or "", e["id"])):
            if e["advance"]:
                ai = adv_info(ents)[e["id"]]
                bal += e["advance"]
                note = e["description"] or ""
                if ai["has"]:
                    note += (" · " if note else "") + f"Prev. balance {ai['carry']:g} carried, total {ai['avail']:g}"
                ws.append([_to_date(e["date"]), e["cat"] or "Site Advance", note or None, _num(e["nos"]) or e["nos"] or None,
                           e["paid_to"] or None, None, e["advance"], bal])
                nos_total += _num(e["nos"])
                for b in sorted(e["br"], key=lambda b: (b["date"] or "", b["id"])):
                    bal -= b["amount"]
                    nos_total += _num(b["nos"])
                    ws.append([_to_date(b["date"] or e["date"]), b["cat"] or None, "From site advance",
                               _num(b["nos"]) or b["nos"] or None, b["paid_to"] or e["paid_to"] or None,
                               b["amount"], None, bal])
            else:
                nos_total += _num(e["nos"])
                ws.append([_to_date(e["date"]), e["cat"] or None, e["description"] or None,
                           _num(e["nos"]) or e["nos"] or None, e["paid_to"] or None, e["amount"] or None, None, None])
        t = totals(ents)
        ws.append([])
        ws.append(["Total", None, None, nos_total or None, None, t["spent"], t["given"], t["balance"]])
        for c in ws[ws.max_row]:
            c.font = Font(bold=True)
        for row in ws.iter_rows(min_row=3):
            row[0].number_format = "dd-mmm-yy"
            for c in row[5:8]:
                c.number_format = "#,##0"
        for i, w in enumerate([12, 22, 26, 7, 30, 13, 12, 13], 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A3"
        summary.append([s["name"], t["given"], t["spent"], t["balance"]])
    if len(site_list) > 1:
        T = totals(entries)
        summary.append(["All sites", T["given"], T["spent"], T["balance"]])
    ws = wb.create_sheet("Summary")
    for r in summary:
        ws.append(r)
    _style_header(ws, 1, 4)
    for row in ws.iter_rows(min_row=2):
        for c in row[1:]:
            c.number_format = "#,##0"
    for i, w in enumerate([22, 14, 14, 14], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def report_workbook(tabs, meta):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Summary"
    for r in [["Star Builders - Report"], *[[k, v] for k, v in meta.items()]]:
        ws.append(r)
    ws["A1"].font = Font(bold=True, size=14)
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 18
    used = {"summary"}
    for t in tabs:
        ws = wb.create_sheet(_safe_title(t["title"], used))
        cols = list(t["df"].columns)
        ws.append(cols)
        _style_header(ws, 1, len(cols))
        for rec in t["df"].itertuples(index=False):
            row = []
            for col, v in zip(cols, rec):
                kind = t["kinds"].get(col)
                if v is None or (isinstance(v, float) and pd.isna(v)):
                    v = None
                elif kind == "date" and re.match(r"\d{4}-", str(v)):
                    v = _to_date(v)
                elif kind == "nos":
                    v = _num(v) if str(v).strip() else None
                row.append(v)
            ws.append(row)
        for r in ws.iter_rows(min_row=2):
            for col, c in zip(cols, r):
                kind = t["kinds"].get(col)
                c.number_format = {"date": "dd-mmm-yy", "money": "#,##0", "pct": "0.0", "nos": "0.##"}.get(kind, "General")
        if t["total_row"]:
            for c in ws[ws.max_row]:
                c.font = Font(bold=True)
        for i, col in enumerate(cols, 1):
            ws.column_dimensions[get_column_letter(i)].width = 24 if t["kinds"].get(col) is None else 14
        ws.freeze_panes = "A2"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ------------------------------------------------------------ Excel import
ALIASES = {
    "date": {"date", "dt", "day", "on"},
    "cat": {"paidfor", "category", "item", "head", "purpose", "particulars"},
    "description": {"description", "descriptionofexpense", "details", "desc", "narration", "remarks"},
    "nos": {"nos", "no", "qty", "quantity", "number"},
    "paid_to": {"paidto", "payee", "vendor", "party", "name", "to"},
    "amount": {"amount", "amountrs", "expense", "expenses", "debit", "paid", "cost", "spent"},
    "advance": {"advance", "advancereceived", "credit", "received"},
}


def _map_col(h):
    k = re.sub(r"[^a-z]", "", str(h or "").lower())
    return next((f for f, names in ALIASES.items() if k in names), None)


def _parse_date(v):
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    s = str(v or "").strip()
    m = re.fullmatch(r"(\d{1,2})[-/ ]([A-Za-z]{3})[a-z]*[-/ ](\d{2,4})", s)
    if m:
        y = m[3] if len(m[3]) == 4 else "20" + m[3]
        mon = [x.lower() for x in MONTHS].index(m[2].lower()) + 1 if m[2].lower() in [x.lower() for x in MONTHS] else 0
        return f"{y}-{mon:02d}-{int(m[1]):02d}" if mon else ""
    m = re.fullmatch(r"(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})", s)
    if m:
        y = m[3] if len(m[3]) == 4 else "20" + m[3]
        return f"{y}-{int(m[2]):02d}-{int(m[1]):02d}"
    return s[:10] if re.match(r"\d{4}-\d{2}-\d{2}", s) else ""


def _text(v):
    if v is None:
        return ""
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v).strip()


def import_workbook(file_obj):
    """Every sheet becomes a site. Headings are matched to the standard ones automatically.
    Rows already present (same site/date/item/person/amount) are skipped, so re-importing is safe."""
    wb = openpyxl.load_workbook(file_obj, data_only=True)
    results = []
    with closing(_connect()) as c:
        for ws in wb.worksheets:
            data = list(ws.iter_rows(values_only=True))
            hi = next((i for i, r in enumerate(data[:15]) if len({_map_col(x) for x in r} - {None}) >= 3), None)
            if hi is None:
                results.append(dict(sheet=ws.title, added=0, skipped=0, note="no recognisable headings - skipped"))
                continue
            cm = [_map_col(x) for x in data[hi]]
            has_cat = "cat" in cm
            row = c.execute("SELECT id FROM sites WHERE name=? COLLATE NOCASE", (ws.title,)).fetchone()
            sid = row["id"] if row else c.execute("INSERT INTO sites(name) VALUES(?)", (ws.title,)).lastrowid
            # rows already in the app BEFORE this import (counted, so genuine repeats inside the file are kept)
            seen = Counter((r["date"], r["cat"].lower(), r["paid_to"].lower(), round(r["amount"], 2), round(r["advance"], 2))
                           for r in c.execute("SELECT date,cat,paid_to,amount,advance FROM entries WHERE site_id=?", (sid,)))
            added = skipped = 0
            for r in data[hi + 1:]:
                o = {f: r[i] for i, f in enumerate(cm) if f and i < len(r)}
                d = _parse_date(o.get("date"))
                amount, advance = _num(o.get("amount")), _num(o.get("advance"))
                if not d or not (amount or advance):
                    continue  # blank rows and "Total" rows
                cat, desc = _text(o.get("cat")), _text(o.get("description"))
                if not has_cat:
                    cat, desc = desc, ""
                cat = norm_cat(cat)
                if cat.lower() == "site advance" and amount > 0 and not advance:
                    advance, amount = amount, 0.0  # money given to a site person
                key = (d, cat.lower(), _text(o.get("paid_to")).lower(), round(amount, 2), round(advance, 2))
                if seen[key] > 0:
                    seen[key] -= 1
                    skipped += 1
                    continue
                c.execute(
                    "INSERT INTO entries(site_id,date,cat,description,nos,paid_to,amount,advance) VALUES(?,?,?,?,?,?,?,?)",
                    (sid, d, cat, desc, _text(o.get("nos")), _text(o.get("paid_to")), amount, advance))
                added += 1
            results.append(dict(sheet=ws.title, added=added, skipped=skipped, note=""))
        c.commit()
    return results
