"""SaaS Billing & Revenue Hub: all the calculation logic from the my casestudy , in pandas.

Each function mirrors a tab or a helper column in the Excel template, so the app and the
workbook always agree. Column names are the same plain labels used in the workbook.
"""
import re
import numpy as np
import pandas as pd

HUB_REQUIRED = ["Deal ID", "Stage", "Start Date", "Contract Value", "End Date"]
HUB_COLUMNS = ["Deal ID", "Customer", "Stage", "Start Date", "Contract Value", "End Date", "Product"]
STR_REQUIRED = ["Invoice ID", "Invoice Date", "Due Date", "Amount", "Amount Paid", "Status"]
STR_COLUMNS = ["Invoice ID", "Customer", "Invoice Date", "Due Date", "Amount", "Amount Paid", "Status", "Deal ID"]
AGING_ORDER = ["Current", "1 to 30", "31 to 60", "61 to 90", "Over 90"]
STATUS_ORDER = ["Paid", "Partially paid", "Past due", "Open", "Duplicate", "Void", "Draft"]
CHECK_ORDER = ["OK", "Missing Deal ID", "No matching deal", "Deal not won"]


def _norm(c):
    return re.sub(r"[\s_]+", " ", str(c)).strip().lower()


def _find_sheet(names, wanted):
    for n in names:
        if _norm(n) == wanted:
            return n
    return None


def _rename(df, columns):
    lookup = {_norm(c): c for c in columns}
    return df.rename(columns={c: lookup[_norm(c)] for c in df.columns if _norm(c) in lookup})


def load_raw(file):
    """Read Raw_HubSpot and Raw_Stripe from one Excel file. Returns (hubspot, stripe, notes)."""
    xl = pd.ExcelFile(file)
    hs_name = _find_sheet(xl.sheet_names, "raw hubspot")
    st_name = _find_sheet(xl.sheet_names, "raw stripe")
    if not hs_name or not st_name:
        raise ValueError("The file needs two sheets named Raw_HubSpot and Raw_Stripe. "
                         f"Sheets found: {', '.join(xl.sheet_names)}.")
    hub = _rename(xl.parse(hs_name), HUB_COLUMNS)
    stp = _rename(xl.parse(st_name), STR_COLUMNS)
    # Stripe exports sometimes call the status column "Stripe Status"
    if "Status" not in stp.columns:
        stp = stp.rename(columns={c: "Status" for c in stp.columns if _norm(c) == "stripe status"})
    miss_h = [c for c in HUB_REQUIRED if c not in hub.columns]
    miss_s = [c for c in STR_REQUIRED if c not in stp.columns]
    if miss_h or miss_s:
        parts = []
        if miss_h: parts.append("Raw_HubSpot is missing: " + ", ".join(miss_h))
        if miss_s: parts.append("Raw_Stripe is missing: " + ", ".join(miss_s))
        raise ValueError(". ".join(parts) + ".")
    return clean_inputs(hub, stp)


def clean_inputs(hub, stp):
    notes = []
    hub = hub.copy(); stp = stp.copy()
    for c in HUB_COLUMNS:
        if c not in hub.columns: hub[c] = ""
    for c in STR_COLUMNS:
        if c not in stp.columns: stp[c] = ""
    hub = hub[HUB_COLUMNS]; stp = stp[STR_COLUMNS]
    hub = hub[hub["Deal ID"].notna() & (hub["Deal ID"].astype(str).str.strip() != "")].reset_index(drop=True)
    stp = stp[stp["Invoice ID"].notna() & (stp["Invoice ID"].astype(str).str.strip() != "")].reset_index(drop=True)
    for df, cols in ((hub, ["Start Date", "End Date"]), (stp, ["Invoice Date", "Due Date"])):
        for c in cols:
            as_text = df[c].map(lambda v: isinstance(v, str)).sum()
            if as_text:
                notes.append(f"{as_text} text dates in {c} were converted to real dates.")
            df[c] = pd.to_datetime(df[c], errors="coerce").dt.normalize()
            bad = int(df[c].isna().sum())
            if bad:
                notes.append(f"{bad} blank or unreadable dates in {c}.")
    hub["Contract Value"] = pd.to_numeric(hub["Contract Value"], errors="coerce").fillna(0.0)
    stp["Amount"] = pd.to_numeric(stp["Amount"], errors="coerce").fillna(0.0)
    stp["Amount Paid"] = pd.to_numeric(stp["Amount Paid"], errors="coerce").fillna(0.0)
    for df in (hub, stp):
        df["Deal ID"] = df["Deal ID"].fillna("").astype(str).str.strip()
        df["Customer"] = df["Customer"].fillna("").astype(str)
    hub["Product"] = hub["Product"].fillna("").astype(str)
    hub["Stage"] = hub["Stage"].astype(str).str.strip().map(lambda v: v.title() if v.lower().startswith("closed") else v)
    stp["Status"] = stp["Status"].astype(str).str.strip().map(lambda v: v.capitalize() if v.lower() in ("void", "draft", "paid", "open") else v)
    return hub, stp, notes


def default_dates(hub, stp):
    """Month-end of the latest invoice, and January of that year, as starting values."""
    last = stp["Invoice Date"].max()
    asof = (last + pd.offsets.MonthEnd(0)).normalize() if pd.notna(last) else pd.Timestamp.today().normalize()
    return asof, pd.Timestamp(year=asof.year, month=1, day=1)


# ------------------------------------------------------------------ Stripe_Working helper columns
def stripe_working(stp, hub, asof):
    s = stp.copy().reset_index(drop=True)
    s["Clean Customer"] = s["Customer"].map(lambda v: re.sub(r"\s+", " ", str(v)).strip().lower())
    s["_month"] = s["Invoice Date"].dt.to_period("M")
    nonvoid = s["Status"] != "Void"
    has_id = s["Deal ID"] != ""
    key = s["Clean Customer"] + "|" + s["Amount"].round(2).astype(str) + "|" + s["Deal ID"] + "|" + s["_month"].astype(str)
    run = s[nonvoid & has_id].groupby(key[nonvoid & has_id]).cumcount()
    s["Duplicate Check"] = ""
    s.loc[run.index[run >= 1], "Duplicate Check"] = "Duplicate"
    s["Valid Flag"] = np.where(s["Status"].isin(["Void", "Draft"]) | (s["Duplicate Check"] == "Duplicate"), "Exclude", "Valid")
    valid = s["Valid Flag"] == "Valid"
    s["Open Balance"] = np.where(valid, s["Amount"] - s["Amount Paid"], 0.0)
    days = (asof - s["Due Date"]).dt.days.clip(lower=0).fillna(0)
    s["Days Past Due"] = np.where(s["Open Balance"] > 0, days, 0).astype(int)
    d = s["Days Past Due"]
    s["Aging Bucket"] = np.select(
        [~valid, s["Open Balance"] <= 0, d == 0, d <= 30, d <= 60, d <= 90],
        ["Excluded", "Paid", "Current", "1 to 30", "31 to 60", "61 to 90"], default="Over 90")
    s["Status (calc)"] = np.select(
        [s["Status"] == "Void", s["Status"] == "Draft", ~valid, s["Open Balance"] <= 0, s["Amount Paid"] > 0, d > 0],
        ["Void", "Draft", "Duplicate", "Paid", "Partially paid", "Past due"], default="Open")
    first = hub.drop_duplicates("Deal ID").set_index("Deal ID")
    s["HubSpot Contract Amount"] = s["Deal ID"].map(first["Contract Value"])
    s["HubSpot Stage"] = s["Deal ID"].map(first["Stage"]).fillna("No Matching Deal")
    s["Invoice Check"] = np.select(
        [~valid, s["Deal ID"] == "", s["HubSpot Stage"] == "No Matching Deal", s["HubSpot Stage"] != "Closed Won"],
        ["Excluded", "Missing Deal ID", "No matching deal", "Deal not won"], default="OK")
    return s.drop(columns=["_month"]).rename(columns={"Status": "Stripe Status", "Status (calc)": "Status"})


# ------------------------------------------------------------------ HubSpot_Working helper columns
def hubspot_working(hub, sw, asof):
    h = hub.copy().reset_index(drop=True)
    billed = sw[sw["Valid Flag"] == "Valid"].groupby("Deal ID")["Amount"].sum()
    won = h["Stage"] == "Closed Won"
    h["Stripe Billed"] = h["Deal ID"].map(billed).fillna(0.0)
    h["Unbilled Variance"] = np.where(won, h["Contract Value"] - h["Stripe Billed"], 0.0)
    v = h["Unbilled Variance"]
    h["Billing Status"] = np.select(
        [~won & (h["Stripe Billed"] > 0), ~won, h["Stripe Billed"] == 0, v > 0.005, v < -0.005],
        ["Billed but not won", "Not Won", "Unbilled", "Underbilled", "Overbilled"], default="Matched")
    h.loc[won & (h["Stripe Billed"] == 0), "Billing Status"] = "Unbilled"
    h["Contract Days"] = (h["End Date"] - h["Start Date"]).dt.days + 1
    elapsed = (np.minimum(h["End Date"], asof) - h["Start Date"]).dt.days + 1
    h["Revenue Recognized To Date"] = np.where(won, elapsed.clip(lower=0) / h["Contract Days"] * h["Contract Value"], 0.0)
    h["Billed Less Recognized"] = np.where(won, h["Stripe Billed"] - h["Revenue Recognized To Date"], 0.0)
    h["Won #"] = np.where(won, won.cumsum(), np.nan)
    return h


# ------------------------------------------------------------------ RevRec_Waterfall and Deferred_Rev
def revrec(hw, first_month, n_months=12):
    months = pd.date_range(first_month, periods=n_months, freq="MS")
    won = hw[hw["Stage"] == "Closed Won"].dropna(subset=["Start Date", "End Date"])
    rows = []
    last_end = months[-1] + pd.offsets.MonthEnd(0)
    for r in won.itertuples(index=False):
        cd = (r[5] - r[3]).days + 1  # End Date - Start Date + 1
        start, end, val = r[3], r[5], r[4]
        row = {"Deal ID": r[0], "Customer": r[1], "Start Date": start, "End Date": end, "Total Contract Value": val, "Contract Days": cd}
        row["Revenue Before Window"] = max(0, (min(end, months[0] - pd.Timedelta(days=1)) - start).days + 1) / cd * val
        for m in months:
            m_end = m + pd.offsets.MonthEnd(0)
            if start > m_end or end < m:
                row[m] = 0.0
            else:
                row[m] = ((min(end, m_end) - max(start, m)).days + 1) / cd * val
        row["Revenue After Window"] = max(0, (end - max(start, last_end + pd.Timedelta(days=1))).days + 1) / cd * val
        row["Check"] = round(val - row["Revenue Before Window"] - sum(row[m] for m in months) - row["Revenue After Window"], 2)
        rows.append(row)
    cols = ["Deal ID", "Customer", "Start Date", "End Date", "Total Contract Value", "Contract Days", "Revenue Before Window"] + list(months) + ["Revenue After Window", "Check"]
    return pd.DataFrame(rows, columns=cols), months


def deferred(sw, rr, months):
    valid = sw[sw["Valid Flag"] == "Valid"]
    before_total = rr["Revenue Before Window"].sum() if len(rr) else 0.0
    opening = valid.loc[valid["Invoice Date"] < months[0], "Amount"].sum() - before_total
    out, begin, cum_rev = [], opening, 0.0
    for m in months:
        m_end = m + pd.offsets.MonthEnd(0)
        bill = valid.loc[(valid["Invoice Date"] >= m) & (valid["Invoice Date"] <= m_end), "Amount"].sum()
        rev = rr[m].sum() if len(rr) else 0.0
        end = begin + bill - rev
        cum_rev += rev
        billed_to_date = valid.loc[valid["Invoice Date"] <= m_end, "Amount"].sum()
        out.append({"Month": m, "Beginning Deferred Revenue": begin, "New Billings": bill, "Revenue Recognized": rev,
                    "Ending Deferred Revenue": end, "Check": "Ties" if abs(end - (billed_to_date - (before_total + cum_rev))) < 0.01 else "Check"})
        begin = end
    return pd.DataFrame(out)


# ------------------------------------------------------------------ Summary_Dashboard
def summary(sw, hw, df_def, asof):
    valid = sw["Valid Flag"] == "Valid"
    won = hw["Stage"] == "Closed Won"
    m = {}
    m["contracted"] = hw.loc[won, "Contract Value"].sum()
    m["unbilled"] = hw.loc[hw["Unbilled Variance"] > 0, "Unbilled Variance"].sum()
    m["overbilled"] = -hw.loc[hw["Unbilled Variance"] < 0, "Unbilled Variance"].sum()
    m["billed"] = sw.loc[valid, "Amount"].sum()
    m["billed_other"] = m["billed"] - hw.loc[won, "Stripe Billed"].sum()
    m["collected"] = sw.loc[valid, "Amount Paid"].sum()
    m["open_ar"] = sw["Open Balance"].sum()
    dup = sw["Duplicate Check"] == "Duplicate"
    m["dup_amt"] = sw.loc[dup, "Amount"].sum(); m["dup_n"] = int(dup.sum())
    m["waterfall_ties"] = abs(m["contracted"] - m["unbilled"] + m["overbilled"] + m["billed_other"] - m["billed"]) < 0.01
    m["ar_ties"] = abs(m["billed"] - m["collected"] - m["open_ar"]) < 0.01
    ag = sw.groupby("Aging Bucket").agg(ar=("Open Balance", "sum"), n=("Invoice ID", "count")).reindex(AGING_ORDER).fillna(0)
    ag["pct"] = ag["ar"] / ag["ar"].sum() if ag["ar"].sum() else 0.0
    m["aging"] = ag
    m["aging_ties"] = abs(ag["ar"].sum() - m["open_ar"]) < 0.01
    m["past_due"] = sw.loc[sw["Days Past Due"] > 0, "Open Balance"].sum()
    m["over60"] = ag.loc[["61 to 90", "Over 90"], "ar"].sum()
    m_start = pd.Timestamp(year=asof.year, month=asof.month, day=1)
    cur = df_def[df_def["Month"] == m_start]
    m["rev_month"] = cur["Revenue Recognized"].sum() if len(cur) else 0.0
    ytd = df_def[(df_def["Month"] <= asof) & (df_def["Month"] >= pd.Timestamp(year=asof.year, month=1, day=1))]
    m["rev_ytd"] = ytd["Revenue Recognized"].sum()
    m["bill_month"] = cur["New Billings"].sum() if len(cur) else 0.0
    m["deferred_end"] = cur["Ending Deferred Revenue"].iloc[0] if len(cur) else np.nan
    m["deferred_deals"] = hw.loc[hw["Billed Less Recognized"] > 0, "Billed Less Recognized"].sum()
    m["unbilled_ar"] = -hw.loc[hw["Billed Less Recognized"] < 0, "Billed Less Recognized"].sum()
    m["rollforward_ties"] = (not np.isnan(m["deferred_end"])) and abs(m["deferred_end"] - (m["deferred_deals"] - m["unbilled_ar"] + m["billed_other"])) < 0.01
    def cnt(mask, col="Amount"): return int(mask.sum()), float(sw.loc[mask, col].sum())
    m["checks"] = pd.DataFrame([
        ("Duplicate Invoices",) + cnt(dup),
        ("Void Invoices",) + cnt(sw["Stripe Status"] == "Void"),
        ("Draft Invoices",) + cnt(sw["Stripe Status"] == "Draft"),
        ("Invoices Missing Deal ID",) + cnt(sw["Invoice Check"] == "Missing Deal ID"),
        ("Invoices, No Matching Deal",) + cnt(sw["Invoice Check"] == "No matching deal"),
        ("Invoices on Deals Not Won",) + cnt(sw["Invoice Check"] == "Deal not won"),
        ("Closed Won, Not Invoiced", int((hw["Billing Status"] == "Unbilled").sum()), float(hw.loc[hw["Billing Status"] == "Unbilled", "Unbilled Variance"].sum())),
        ("Closed Won, Underbilled", int((hw["Billing Status"] == "Underbilled").sum()), float(hw.loc[hw["Billing Status"] == "Underbilled", "Unbilled Variance"].sum())),
        ("Closed Won, Overbilled", int((hw["Billing Status"] == "Overbilled").sum()), float(-hw.loc[hw["Billing Status"] == "Overbilled", "Unbilled Variance"].sum())),
    ], columns=["Check", "Count", "Amount ($)"]).set_index("Check")
    return m


def takeaways(m):
    c = m["checks"]; g = lambda k: int(c.loc[k, "Count"]); a = lambda k: float(c.loc[k, "Amount ($)"])
    money = lambda v: f"${v:,.0f}"
    t1 = (f"1. Collect {money(m['open_ar'])} in open AR. {money(m['past_due'])} is past due and {money(m['over60'])} is over 60 days. "
          f"Bill {money(a('Closed Won, Not Invoiced') + a('Closed Won, Underbilled'))} on {g('Closed Won, Not Invoiced') + g('Closed Won, Underbilled')} Closed Won deals not fully invoiced. "
          f"{g('Invoices Missing Deal ID')} invoice(s) have no Deal ID, so check those first.")
    t2 = (f"2. Void or refund {g('Duplicate Invoices')} duplicate invoices ({money(a('Duplicate Invoices'))}). Also look at "
          f"{g('Invoices, No Matching Deal') + g('Invoices on Deals Not Won')} invoices ({money(a('Invoices, No Matching Deal') + a('Invoices on Deals Not Won'))}) "
          f"with no matching deal or on a deal not won, and {g('Closed Won, Overbilled')} overbilled deal(s) ({money(a('Closed Won, Overbilled'))}).")
    n = g("Duplicate Invoices") + g("Invoices Missing Deal ID") + g("Invoices, No Matching Deal") + g("Invoices on Deals Not Won") + g("Closed Won, Overbilled")
    t3 = (f"3. Rule going forward: Stripe should not send an invoice unless the Deal ID exists in HubSpot, the deal is Closed Won, and the amount equals the Contract Value. "
          f"That rule would catch {n} of the issues here.")
    return [t1, t2, t3]


# ------------------------------------------------------------------ Dashboard (Support tab) filters
def dashboard(sw, month=None, status=None, aging=None, check=None, months=None):
    """month is a Timestamp (first of month) or None for All. Each table skips its own filter."""
    def mask(skip=()):
        k = pd.Series(True, index=sw.index)
        if month is not None and "month" not in skip:
            k &= (sw["Invoice Date"] >= month) & (sw["Invoice Date"] < month + pd.offsets.MonthBegin(1))
        if status and "status" not in skip: k &= sw["Status"] == status
        if aging and "aging" not in skip: k &= sw["Aging Bucket"] == aging
        if check and "check" not in skip: k &= sw["Invoice Check"] == check
        return k
    valid = sw["Valid Flag"] == "Valid"
    k = mask()
    out = {"rows": sw[k]}
    out["kpi"] = {"invoices": int(k.sum()), "billed": sw.loc[k & valid, "Amount"].sum(), "collected": sw.loc[k & valid, "Amount Paid"].sum(),
                  "open_ar": sw.loc[k, "Open Balance"].sum(), "past_due": sw.loc[k & (sw["Days Past Due"] > 0), "Open Balance"].sum()}
    ka = mask({"aging"})
    out["aging"] = sw[ka].groupby("Aging Bucket")["Open Balance"].sum().reindex(AGING_ORDER).fillna(0)
    ks = mask({"status"})
    out["status"] = sw[ks].groupby("Status")["Invoice ID"].count().reindex(STATUS_ORDER).fillna(0)
    kc = mask({"check"})
    out["check"] = sw[kc].groupby("Invoice Check")["Amount"].sum().reindex(CHECK_ORDER).fillna(0)
    km = mask({"month"})
    sub = sw[km & valid]
    by_m = sub.groupby(sub["Invoice Date"].dt.to_period("M").dt.to_timestamp()).agg(Billed=("Amount", "sum"), Collected=("Amount Paid", "sum"))
    if months is not None: by_m = by_m.reindex(months).fillna(0)
    out["by_month"] = by_m
    billed = out["kpi"]["billed"]
    out["collection_rate"] = out["kpi"]["collected"] / billed if billed else 0.0
    out["past_due_share"] = out["kpi"]["past_due"] / out["kpi"]["open_ar"] if out["kpi"]["open_ar"] else 0.0
    return out


def run_all(hub, stp, asof, first_month, n_months=12):
    sw = stripe_working(stp, hub, asof)
    hw = hubspot_working(hub, sw, asof)
    rr, months = revrec(hw, first_month, n_months)
    df_def = deferred(sw, rr, months)
    m = summary(sw, hw, df_def, asof)
    return dict(sw=sw, hw=hw, rr=rr, months=months, deferred=df_def, summary=m, takeaways=takeaways(m))
