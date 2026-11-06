"""Star Builders - site expense book (Streamlit).  Run:  streamlit run app.py"""
import hmac
import html
from datetime import date, datetime, timedelta

import pandas as pd
import streamlit as st

import core

st.set_page_config(page_title="Star Builders", page_icon="🏗️", layout="centered")
core.init_db()
inr, fd = core.inr, core.fd

st.markdown(
    """<style>
.card{border:1px solid rgba(128,128,128,.35);border-radius:10px;padding:8px 12px;margin-bottom:8px}
.card .t{font-weight:600}.card .s{opacity:.7;font-size:.85em}.card .r{text-align:right;font-weight:600;white-space:nowrap}
.card .row{display:flex;justify-content:space-between;gap:10px}
.card .stats{display:flex;gap:18px;flex-wrap:wrap;margin-top:6px}.card .stats small{display:block;opacity:.7}
.pa{background:rgba(245,166,35,.18);border-color:#e8a317}
.br{margin-left:18px;border-left:3px solid #0f6b5c}
.tag{background:#e8a317;color:#1d1f23;border-radius:6px;padding:1px 6px;font-size:11px;font-weight:700}
.neg{color:#d9534f}.pos{color:#2e9e6f}
</style>""",
    unsafe_allow_html=True,
)


# ------------------------------------------------------------------ login
def secret(name, default):
    """Passwords come from .streamlit/secrets.toml (see secrets.toml.example); defaults below."""
    try:
        return str(st.secrets[name])
    except Exception:
        return default


APP_PASSWORD = secret("app_password", "star123")
ADMIN_PASSWORD = secret("admin_password", "adminstar")
ss = st.session_state
ss.setdefault("role", None)
ss.setdefault("ver", 0)


def role_for(pw):
    if hmac.compare_digest(pw.encode(), ADMIN_PASSWORD.encode()):
        return "admin"
    if hmac.compare_digest(pw.encode(), APP_PASSWORD.encode()):
        return "user"
    return None


if not ss.role:
    st.title("🏗️ Star Builders")
    with st.form("login"):
        pw = st.text_input("Enter app password", type="password")
        if st.form_submit_button("Open", type="primary"):
            r = role_for(pw)
            if r:
                ss.role = r
                st.rerun()
            st.error("Wrong password")
    st.stop()

admin = ss.role == "admin"
if ss.pop("flash", None):
    st.toast("Saved ✔")


def admin_login_form(key):
    with st.form(key):
        pw = st.text_input("Admin password", type="password")
        if st.form_submit_button("Admin login"):
            if role_for(pw) == "admin":
                ss.role = "admin"
                st.rerun()
            st.error("Wrong admin password")


# ---------------------------------------------------------------- sidebar
site_list = core.sites()
with st.sidebar:
    st.markdown("### 🏗️ Star Builders")
    st.caption("Signed in as: **" + ("Admin" if admin else "User") + "**")
    site = None
    if site_list:
        name = st.selectbox("Site", [s["name"] for s in site_list], key="site_sel")
        site = next(s for s in site_list if s["name"] == name)
    if not admin:
        with st.expander("🔑 Admin login"):
            admin_login_form("adm_side")
    if st.button("🔒 Lock app"):
        ss.role = None
        st.rerun()

tab_e, tab_a, tab_r, tab_s = st.tabs(["📋 Entries", "➕ Add / Edit", "📊 Reports", "⚙️ Sites & Excel"])
ents = core.load_entries(site["id"]) if site else []
ai = core.adv_info(ents)
if admin:
    all_e = core.load_entries()
    xl_all = core.export_workbook(all_e, site_list)
    xl_site = core.export_workbook(ents, [site]) if site else None
    csv_site = core.entries_csv(ents) if site else None
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# ---------------------------------------------------------------- entries
with tab_e:
    if not site:
        st.info("No sites yet. Go to **Sites & Excel** to add a site or import your Excel.")
    else:
        T = core.totals(ents)
        c1, c2, c3 = st.columns(3)
        c1.metric("Site advance given", inr(T["given"]))
        c2.metric("Total spent", inr(T["spent"]))
        c3.metric("Advance balance", inr(T["balance"]))
        if admin:
            d1, d2, d3 = st.columns(3)
            d1.download_button("⬇ Excel – this site", xl_site, file_name=f"{site['name']}.xlsx", mime=XLSX, key="dl_site_e")
            d2.download_button("⬇ Excel – all sites", xl_all, file_name="Star_Builders_P_L_Report.xlsx", mime=XLSX, key="dl_all_e")
            d3.download_button("⬇ CSV – this site", csv_site, file_name=f"{site['name']}.csv", mime="text/csv", key="dl_csv_e")
        q = st.text_input("Search", placeholder="item, person, description…").strip().lower()
        h = []
        for e in sorted(ents, key=lambda e: (e["date"] or "", e["id"]), reverse=True):
            esc = html.escape
            if e["advance"]:
                i = ai[e["id"]]
                if q and q not in " ".join([e["paid_to"], e["description"], e["cat"], "site advance"] +
                                           [b["cat"] + b["paid_to"] for b in e["br"]]).lower():
                    continue
                st_ = lambda k, v, c="": f'<div><small>{k}</small><b class="{c}">{inr(v)}</b></div>'
                h.append(
                    f'<div class="card"><div class="t">{esc(e["cat"] or "Site Advance")} <span class="s">(site advance) · #{e["id"]}</span></div>'
                    f'<div class="s">{fd(e["date"])} · {esc(e["paid_to"] or "-")}{" · " + esc(e["description"]) if e["description"] else ""}</div>'
                    f'<div class="stats">{st_("Advance", e["advance"], "pos")}'
                    + (st_("Previous balance", i["carry"]) + st_("Total", i["avail"], "pos") if i["has"] else "")
                    + st_("Used", i["used"]) + st_("Balance", i["bal"], "neg" if i["bal"] < 0 else "")
                    + "</div>" + ('<div class="s">Balance carried forward to the next advance</div>' if i["next"] else "") + "</div>")
                for b in sorted(e["br"], key=lambda b: (b["date"] or "", b["id"])):
                    h.append(
                        f'<div class="card br"><div class="row"><div><div class="t">{esc(b["cat"] or "(no item)")}</div>'
                        f'<div class="s">{fd(b["date"])}{" · " + esc(b["paid_to"]) if b["paid_to"] else ""}'
                        f'{" · Nos " + esc(b["nos"]) if b["nos"] else ""} · from site advance ({esc(e["paid_to"] or "advance")})</div></div>'
                        f'<div class="r">{inr(b["amount"])}</div></div></div>')
            else:
                if q and q not in " ".join([e["paid_to"], e["description"], e["cat"], e["nos"]]).lower():
                    continue
                pa = core.is_personal(e["cat"])
                h.append(
                    f'<div class="card{" pa" if pa else ""}"><div class="row"><div><div class="t">{esc(e["cat"] or "(no item)")}'
                    f'{" <span class=tag>PERSONAL ADVANCE</span>" if pa else ""} <span class="s">#{e["id"]}</span></div>'
                    f'<div class="s">{fd(e["date"])}{" · " + esc(e["paid_to"]) if e["paid_to"] else ""}'
                    f'{" · " + esc(e["description"]) if e["description"] else ""}{" · Nos " + esc(e["nos"]) if e["nos"] else ""}</div></div>'
                    f'<div class="r">{inr(e["amount"])}</div></div></div>')
        st.markdown("".join(h) if h else "No entries yet. Use **Add / Edit**.", unsafe_allow_html=True)

# --------------------------------------------------------------- add / edit
NEW = "➕ New…"


def pick(label, options, current, key):
    opts = sorted({o for o in options if o} | ({current} if current else set()), key=str.lower)
    choices = ["(select)"] + opts + [NEW]
    sel = st.selectbox(label, choices, index=choices.index(current) if current else 0, key=key)
    if sel == NEW:
        return st.text_input(f"New {label.lower()}", key=key + "_new").strip()
    return "" if sel == "(select)" else sel


def to_date(v):
    try:
        return datetime.strptime(str(v)[:10], "%Y-%m-%d").date()
    except Exception:
        return date.today()


with tab_a:
    if not site:
        st.info("Add a site first (Sites & Excel tab).")
    else:
        options = {"➕ New entry": None}
        for e in sorted(ents, key=lambda e: (e["date"] or "", e["id"]), reverse=True)[:300]:
            options[f"#{e['id']} · {fd(e['date'])} · {e['cat'] or 'Site Advance'} · {e['paid_to'] or '-'} · {inr(e['advance'] or e['amount'])}"] = e
        ver = ss.ver
        choice = st.selectbox("Add new, or pick an entry to edit", list(options), key=f"pick_{site['id']}_{ver}")
        cur = options[choice]
        eid = cur["id"] if cur else 0
        k = lambda n: f"{n}_{site['id']}_{eid}_{ver}"

        kinds = ["Expense", "Site advance", "Personal advance"]
        dk = 0 if not cur else (1 if cur["advance"] else (2 if core.is_personal(cur["cat"]) else 0))
        kind = st.radio("Type", kinds, index=dk, horizontal=True, key=k("kind"))
        st.caption({
            "Expense": "Money paid out for site work (materials, labour, etc.).",
            "Site advance": "Money given to a site person to spend on site. Add a break-up below.",
            "Personal advance": "Advance given to a person for personal use (e.g. salary advance). Shown highlighted.",
        }[kind])
        c1, c2 = st.columns(2)
        d = c1.date_input("Date", value=to_date(cur["date"]) if cur else date.today(), key=k("date"), format="DD/MM/YYYY")
        amt = c2.number_input("Amount (₹)", min_value=0.0, step=100.0,
                              value=float((cur["advance"] or cur["amount"]) if cur else 0.0), key=k("amt"))
        if kind == "Expense":
            cat = pick("Paid for", core.distinct("cat"), core.norm_cat(cur["cat"]) if cur and not cur["advance"] else "", k("cat"))
        else:
            cat = "Personal advance" if kind == "Personal advance" else "Site Advance"
        desc = st.text_input("Description (optional)", value=cur["description"] if cur else "", key=k("desc"))
        c3, c4 = st.columns(2)
        nos = c3.text_input("Nos / qty", value=cur["nos"] if cur else "", key=k("nos"))
        with c4:
            paid_to = pick("Paid to", core.distinct("paid_to"), cur["paid_to"] if cur else "", k("to"))

        brs = []
        if kind == "Site advance":
            st.markdown("**Break-up of this advance (how it was spent)**")
            st.caption("To add a line, click the empty “+” row at the bottom of the table and fill Date, Item, Nos, Amount and Paid to. To remove a line, tick its row at the left and press Delete. Then press Save entry.")
            cols = ["Date", "Item", "Nos", "Amount", "Paid to"]
            if cur and cur["br"]:
                base = pd.DataFrame([[pd.Timestamp(to_date(b["date"])), b["cat"], b["nos"], b["amount"], b["paid_to"]]
                                     for b in cur["br"]], columns=cols)
            else:
                base = pd.DataFrame({"Date": pd.Series(dtype="datetime64[ns]"), "Item": pd.Series(dtype="str"),
                                     "Nos": pd.Series(dtype="str"), "Amount": pd.Series(dtype="float"),
                                     "Paid to": pd.Series(dtype="str")})
            edited = st.data_editor(
                base, num_rows="dynamic", hide_index=True, key=k("br"),
                column_config={"Date": st.column_config.DateColumn(format="DD-MMM-YY"),
                               "Amount": st.column_config.NumberColumn(min_value=0.0, format="%.2f")})
            for _, r in edited.iterrows():
                a = r["Amount"]
                if pd.isna(a) or a <= 0:
                    continue
                brs.append(dict(
                    date=(r["Date"].date() if not pd.isna(r["Date"]) else d).isoformat(),
                    cat=core.norm_cat("" if pd.isna(r["Item"]) else r["Item"]),
                    nos="" if pd.isna(r["Nos"]) else str(r["Nos"]).strip(),
                    paid_to="" if pd.isna(r["Paid to"]) else str(r["Paid to"]).strip(), amount=float(a)))
            tmp = dict(id=eid or 10**12, site_id=site["id"], date=d.isoformat(), paid_to=paid_to,
                       advance=amt or 1e-9, br=brs)
            info = core.adv_info([e for e in ents if e["id"] != eid] + [tmp])[tmp["id"]]
            m = st.columns(4 if info["has"] else 3)
            j = 0
            if info["has"]:
                m[0].metric("Previous balance", inr(info["carry"]))
                j = 1
            m[j].metric("Total available", inr(info["avail"]))
            m[j + 1].metric("Used", inr(info["used"]))
            m[j + 2].metric("Balance", inr(info["bal"]))
            if info["bal"] < 0:
                st.warning("The break-up is more than the available advance.")

        if st.button("💾 Save entry", type="primary", key=k("save")):
            if amt <= 0:
                st.error("Please enter the amount.")
            elif kind == "Expense" and not cat:
                st.error("Please choose what it was paid for.")
            else:
                adv = kind == "Site advance"
                core.save_entry(
                    dict(site_id=site["id"], date=d.isoformat(), cat=core.norm_cat(cat), description=desc.strip(),
                         nos=nos.strip(), paid_to=paid_to, amount=0.0 if adv else amt, advance=amt if adv else 0.0),
                    brs if adv else [], eid or None)
                ss.ver += 1
                ss.flash = True
                st.rerun()
        if cur and admin:
            with st.expander("Delete this entry"):
                if st.checkbox("Yes, delete it permanently", key=k("delok")) and st.button("🗑 Delete", key=k("del")):
                    core.delete_entry(eid)
                    ss.ver += 1
                    st.rerun()

# ---------------------------------------------------------------- reports
with tab_r:
    if not admin:
        st.info("Reports are for the admin only. Enter the admin password:")
        admin_login_form("adm_rep")
    else:
        all_ents = core.load_entries()
        ai_all = core.adv_info(all_ents)
        scope = st.selectbox("Site", ["All sites"] + [s["name"] for s in site_list], key="rep_site")
        period = st.radio("Period", ["All time", "This month", "Last 30 days", "Custom"], horizontal=True)
        today = date.today()
        frm = to = None
        if period == "This month":
            frm = today.replace(day=1)
        elif period == "Last 30 days":
            frm = today - timedelta(days=29)
        elif period == "Custom":
            c1, c2 = st.columns(2)
            frm = c1.date_input("From", value=today.replace(day=1), format="DD/MM/YYYY", key="rep_from")
            to = c2.date_input("To", value=today, format="DD/MM/YYYY", key="rep_to")
        sel = [e for e in all_ents
               if (scope == "All sites" or e["site"] == scope)
               and (not frm or (e["date"] or "") >= frm.isoformat())
               and (not to or (e["date"] or "") <= to.isoformat())]
        tabs, T = core.build_report(sel, ai_all, scope == "All sites")
        c1, c2 = st.columns(2)
        c1.metric("Total spent", inr(T["spent"]))
        c2.metric("Entries", len(sel))
        c1.metric("Site advance given", inr(T["given"]))
        c2.metric("Advance balance", inr(T["balance"]))
        for t in tabs:
            with st.expander(t["title"], expanded=t["open"]):
                if t["df"].empty:
                    st.caption("No data")
                    continue
                if t["hl"]:
                    st.markdown('<span class="tag">PERSONAL ADVANCE</span>', unsafe_allow_html=True)
                dd = core.display_df(t)
                n = len(dd)

                def style(row, t=t, n=n):
                    css = []
                    pa = core.is_personal(row.iloc[0]) or t["hl"]
                    for _ in row:
                        css.append(("background-color: rgba(245,166,35,.25);" if pa and row.name != n - 1 else "")
                                   + ("font-weight:700;" if t["total_row"] and row.name == n - 1 else ""))
                    return css

                st.dataframe(dd.style.apply(style, axis=1), hide_index=True)
        meta = {"Site": scope, "From": str(frm or "start"), "To": str(to or "latest"),
                "Total spent": T["spent"], "Site advance given": T["given"], "Advance balance": T["balance"],
                "Entries": len(sel)}
        st.download_button("⬇ Download this report (Excel)", core.report_workbook(tabs, meta),
                           file_name="Star_Builders_Report.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ---------------------------------------------------------- sites & excel
with tab_s:
    st.subheader("Sites")
    c1, c2 = st.columns([3, 1])
    new_name = c1.text_input("New site name", key=f"newsite_{ss.ver}", label_visibility="collapsed", placeholder="New site name")
    if c2.button("＋ Add site"):
        try:
            core.add_site(new_name)
            ss.ver += 1
            st.rerun()
        except ValueError as ex:
            st.error(str(ex))
    if site:
        c1, c2 = st.columns([3, 1])
        ren = c1.text_input("Rename this site", value=site["name"], key=f"ren_{site['id']}_{ss.ver}", label_visibility="collapsed")
        if c2.button("Rename"):
            try:
                core.rename_site(site["id"], ren)
                st.rerun()
            except ValueError as ex:
                st.error(str(ex))

    if admin:
        st.subheader("Import Excel (admin)")
        st.caption("Every sheet becomes a site and the headings are matched automatically. "
                   "Rows already in the app are skipped, so importing the same file twice is safe.")
        up = st.file_uploader("Excel file (.xlsx)", type=["xlsx"])
        if up and st.button("Import this file", type="primary"):
            try:
                for r in core.import_workbook(up):
                    st.write(f"**{r['sheet']}** – added {r['added']}, skipped {r['skipped']} {r['note']}")
            except Exception as ex:  # unreadable file
                st.error(f"Could not read that file: {ex}")

        st.subheader("Export (admin)")
        st.download_button("⬇ Export all sites to Excel", xl_all, file_name="Star_Builders_P_L_Report.xlsx", mime=XLSX, key="dl_all_s")
        if site:
            st.download_button(f"⬇ Export “{site['name']}” to Excel", xl_site, file_name=f"{site['name']}.xlsx", mime=XLSX, key="dl_site_s")
            st.download_button(f"⬇ Export “{site['name']}” as CSV", csv_site, file_name=f"{site['name']}.csv", mime="text/csv", key="dl_csv_s")
            with st.expander("Delete this site"):
                st.warning(f"“{site['name']}” and its {len(ents)} entries will be deleted permanently.")
                if st.checkbox("Yes, delete this site") and st.button("🗑 Delete site"):
                    core.delete_site(site["id"])
                    st.rerun()
    else:
        st.caption("Import, export and deleting a site need the admin login (sidebar).")
