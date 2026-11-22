"""SaaS Billing & Revenue Hub
Upload one Excel file with 2 sheets (Raw_HubSpot and Raw_Stripe) and the app rebuilds the
Case Study 3 workbook: order to cash, AR aging, revenue schedule, deferred revenue, dashboard.
"""
import io
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

import saas_billing_logic as L
from excel_export import build_workbook

NAVY, GREEN, GREEN_L, RED, AMBER, GREY = "#1F2A44", "#2E7D32", "#81C784", "#D32F2F", "#F9A825", "#B8C2CF"
SAMPLE = Path(__file__).parent / "Sample_Raw_Data.xlsx"
money = lambda v: f"${v:,.0f}"

st.set_page_config(page_title="SaaS Billing & Revenue Hub", page_icon=":bar_chart:", layout="wide")
st.markdown("""
<style>
.block-container {padding-top: 2rem;}
div[data-testid="stMetric"] {background:#F3F5F8; border-left:4px solid #2E7D32; padding:10px 14px; border-radius:6px;}
.takeaway {background:#E8F5E9; border-left:4px solid #2E7D32; padding:10px 14px; margin-bottom:8px; border-radius:4px;}
.ties {color:#2E7D32; font-weight:600;} .check {color:#D32F2F; font-weight:600;}
</style>""", unsafe_allow_html=True)


# ------------------------------------------------------------------ helpers
@st.cache_data(show_spinner=False)
def load(data: bytes):
    return L.load_raw(io.BytesIO(data))


@st.cache_data(show_spinner="Calculating...")
def compute(hub, stp, asof, first_month):
    return L.run_all(hub, stp, pd.Timestamp(asof), pd.Timestamp(first_month))


@st.cache_data(show_spinner="Building the Excel workbook...")
def excel_with_formulas(hub, stp, asof, first_month):
    return build_workbook(hub, stp, asof, first_month)


def style(fig, h=330):
    fig.update_layout(template="plotly_white", height=h, margin=dict(l=10, r=10, t=50, b=10), font=dict(family="Arial", size=12),
                      title=dict(font=dict(size=15)), legend=dict(orientation="h", y=-0.15))
    return fig


def tie(ok):
    return '<span class="ties">Ties</span>' if ok else '<span class="check">Check</span>'


def month_label(df):
    return df.rename(columns={c: c.strftime("%b %y") for c in df.columns if isinstance(c, pd.Timestamp)})


def values_workbook(r):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        m = r["summary"]
        pd.DataFrame({"Line": ["Total Contracted (Closed Won)", "(-) Unbilled or Underbilled Deals", "(+) Overbilled vs Contract", "(+) Billed, No Closed Won Deal",
                               "(=) Valid Total Billed (Stripe)", "(-) Total Cash Collected", "(=) Total Open AR"],
                      "Amount": [m["contracted"], m["unbilled"], m["overbilled"], m["billed_other"], m["billed"], m["collected"], m["open_ar"]]}).to_excel(xw, sheet_name="Order to Cash", index=False)
        m["aging"].reset_index().to_excel(xw, sheet_name="AR Aging", index=False)
        m["checks"].reset_index().to_excel(xw, sheet_name="Data Checks", index=False)
        month_label(r["rr"]).to_excel(xw, sheet_name="RevRec_Waterfall", index=False)
        r["deferred"].assign(Month=r["deferred"]["Month"].dt.strftime("%b %y")).to_excel(xw, sheet_name="Deferred_Rev", index=False)
        r["sw"].to_excel(xw, sheet_name="Stripe_Working", index=False)
        r["hw"].to_excel(xw, sheet_name="HubSpot_Working", index=False)
    return buf.getvalue()


# ------------------------------------------------------------------ sidebar
st.sidebar.title("SaaS Billing & Revenue Hub")
st.sidebar.markdown("**1. Upload your data**")
upload = st.sidebar.file_uploader("Choose a file", type=["xlsx"], help="One Excel file with two sheets: Raw_HubSpot and Raw_Stripe.")
if SAMPLE.exists():
    st.sidebar.download_button("Download the sample file", SAMPLE.read_bytes(), file_name="Sample_Raw_Data.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    use_sample = st.sidebar.checkbox("Use the sample file", value=False, disabled=upload is not None)
else:
    use_sample = False

data = upload.getvalue() if upload else (SAMPLE.read_bytes() if use_sample else None)

st.title("SaaS Billing & Revenue Hub")
if data is None:
    st.write("A finance operations tool for SaaS companies. Upload a HubSpot deals export and a Stripe invoices export and get order to cash, invoice data quality checks, AR aging, daily-prorated revenue recognition (ASC 606 style) and a deferred revenue rollforward, in one click.")
    st.markdown("""
1. Put your HubSpot deals on a sheet named **Raw_HubSpot** and your Stripe invoices on a sheet named **Raw_Stripe**, in one `.xlsx` file.
2. Click **Browse files** on the left and choose that file.
3. Read the Summary, explore the Dashboard, and download the Excel workbook with live formulas.

**Raw_HubSpot columns:** Deal ID, Customer, Stage, Start Date, Contract Value, End Date, Product  
**Raw_Stripe columns:** Invoice ID, Customer, Invoice Date, Due Date, Amount, Amount Paid, Status, Deal ID

No file yet? Download the sample file from the sidebar and upload it back, or tick **Use the sample file**.""")
    st.stop()

try:
    hub, stp, notes = load(data)
except Exception as e:  # friendly message instead of a traceback
    st.error(str(e))
    st.stop()

default_asof, default_first = L.default_dates(hub, stp)
st.sidebar.markdown("**2. Settings**")
asof = st.sidebar.date_input("As of date (use a month end)", value=default_asof.date())
first_month = st.sidebar.date_input("Revenue schedule starts", value=default_first.date(), help="The schedule shows 12 months from this date.")
first_month = pd.Timestamp(first_month).replace(day=1)
st.sidebar.caption("Every aging and revenue number reads these two dates.")

r = compute(hub, stp, pd.Timestamp(asof), first_month)
m, sw, hw, rr, dfd, months = r["summary"], r["sw"], r["hw"], r["rr"], r["deferred"], r["months"]
st.caption(f"{len(hub)} deals and {len(stp)} invoices loaded. As of {pd.Timestamp(asof):%b %d, %Y}.")
for n in notes:
    st.warning(n)
if len(stp) > 3000:
    st.info("Large file: the Excel workbook download uses running-count formulas and may calculate slowly. The app itself is fine.")

tab_sum, tab_dash, tab_rev, tab_def, tab_hw, tab_sw, tab_sup, tab_raw, tab_dl = st.tabs(
    ["Summary_Dashboard", "Dashboard", "RevRec_Waterfall", "Deferred_Rev", "HubSpot_Working", "Stripe_Working", "Support", "Raw Data", "Download"])

# ------------------------------------------------------------------ Summary
with tab_sum:
    c = st.columns(5)
    c[0].metric("Total Contracted", money(m["contracted"])); c[1].metric("Valid Billed", money(m["billed"])); c[2].metric("Cash Collected", money(m["collected"]))
    c[3].metric("Open AR", money(m["open_ar"])); c[4].metric("Past Due AR", money(m["past_due"]))

    st.subheader("Executive takeaways")
    for t in r["takeaways"]:
        st.markdown(f'<div class="takeaway">{t}</div>', unsafe_allow_html=True)

    left, right = st.columns([1, 1.25])
    with left:
        st.subheader("A. Order to Cash")
        a = pd.DataFrame({"Line": ["Total Contracted (Closed Won)", "(-) Unbilled or Underbilled Deals", "(+) Overbilled vs Contract", "(+) Billed, No Closed Won Deal",
                                   "(=) Valid Total Billed (Stripe)", "(-) Total Cash Collected", "(=) Total Open AR", "Memo: Duplicate Invoices ($)", "Memo: Duplicate Invoices (#)"],
                          "Amount": [m["contracted"], m["unbilled"], m["overbilled"], m["billed_other"], m["billed"], m["collected"], m["open_ar"], m["dup_amt"], m["dup_n"]]})
        a["Amount"] = [money(v) for v in a["Amount"][:8]] + [f"{int(a['Amount'].iloc[8])}"]
        st.dataframe(a, hide_index=True, width="stretch")
        st.markdown(f"Waterfall check: {tie(m['waterfall_ties'])} &nbsp;|&nbsp; Billed minus collected equals open AR: {tie(m['ar_ties'])}", unsafe_allow_html=True)
    with right:
        fig = go.Figure(go.Waterfall(
            orientation="v", measure=["absolute", "relative", "relative", "relative", "total", "relative", "total"],
            x=["Contracted", "Unbilled", "Overbilled", "Billed, no won deal", "Valid billed", "Collected", "Open AR"],
            y=[m["contracted"], -m["unbilled"], m["overbilled"], m["billed_other"], 0, -m["collected"], 0],
            increasing=dict(marker=dict(color=GREEN_L)), decreasing=dict(marker=dict(color=RED)), totals=dict(marker=dict(color=NAVY)),
            connector=dict(line=dict(color=GREY))))
        fig.update_yaxes(tickprefix="$", tickformat=",.0f")
        st.plotly_chart(style(fig.update_layout(title="Contract to cash"), 360), width="stretch")

    left, right = st.columns([1, 1.25])
    with left:
        st.subheader("B. AR Aging")
        ag = m["aging"].copy(); tot = ag[["ar", "n"]].sum()
        show = pd.DataFrame({"Bucket": list(ag.index) + ["Total Open AR"], "Open AR ($)": [money(v) for v in ag["ar"]] + [money(tot["ar"])],
                             "Invoices": [int(v) for v in ag["n"]] + [int(tot["n"])], "% of total": [f"{v:.1%}" for v in ag["pct"]] + ["100.0%" if tot["ar"] else "0.0%"]})
        st.dataframe(show, hide_index=True, width="stretch")
        st.markdown(f"Matches open AR above: {tie(m['aging_ties'])}", unsafe_allow_html=True)
    with right:
        fig = go.Figure(go.Bar(x=list(ag.index), y=ag["ar"], marker_color=NAVY, text=[money(v) for v in ag["ar"]], textposition="outside"))
        fig.update_yaxes(tickprefix="$", tickformat=",.0f")
        st.plotly_chart(style(fig.update_layout(title="Open AR by days past due"), 300), width="stretch")

    left, right = st.columns(2)
    with left:
        st.subheader("C. Revenue and Deferred Revenue")
        cdf = pd.DataFrame({"Line": ["Revenue This Month", "Revenue Year to Date", "Billings This Month", "Deferred Revenue (Rollforward)", "Deferred Revenue by Deal", "Unbilled AR by Deal"],
                            "Amount": [money(m["rev_month"]), money(m["rev_ytd"]), money(m["bill_month"]), money(m["deferred_end"]) if m["deferred_end"] == m["deferred_end"] else "n/a",
                                       money(m["deferred_deals"]), money(m["unbilled_ar"])]})
        st.dataframe(cdf, hide_index=True, width="stretch")
        st.markdown(f"Rollforward matches the deal-level view: {tie(m['rollforward_ties'])}", unsafe_allow_html=True)
        if m["deferred_end"] != m["deferred_end"]:
            st.info("The as-of month is outside the 12-month revenue schedule. Change the schedule start date in the sidebar.")
    with right:
        st.subheader("D. Data Checks")
        chk = m["checks"].copy(); chk["Amount ($)"] = [money(v) for v in chk["Amount ($)"]]
        st.dataframe(chk, width="stretch")

# ------------------------------------------------------------------ Dashboard
with tab_dash:
    data_months = pd.period_range(sw["Invoice Date"].min(), sw["Invoice Date"].max(), freq="M").to_timestamp() if sw["Invoice Date"].notna().any() else pd.DatetimeIndex([])
    f = st.columns(4)
    sel_month = f[0].selectbox("Month", ["All"] + [d.strftime("%b %y") for d in data_months])
    sel_status = f[1].selectbox("Status", ["All"] + L.STATUS_ORDER)
    sel_aging = f[2].selectbox("Aging Bucket", ["All", "Paid"] + L.AGING_ORDER + ["Excluded"])
    sel_check = f[3].selectbox("Deal Check", ["All"] + L.CHECK_ORDER + ["Excluded"])
    month = None if sel_month == "All" else data_months[[d.strftime("%b %y") for d in data_months].index(sel_month)]
    d = L.dashboard(sw, month, None if sel_status == "All" else sel_status, None if sel_aging == "All" else sel_aging,
                    None if sel_check == "All" else sel_check, months=data_months)
    k = d["kpi"]; c = st.columns(5)
    c[0].metric("# of Invoices", f"{k['invoices']:,}"); c[1].metric("Total Billed", money(k["billed"])); c[2].metric("Total Collected", money(k["collected"]))
    c[3].metric("Open AR", money(k["open_ar"])); c[4].metric("Past Due AR", money(k["past_due"]))

    c1, c2, c3 = st.columns(3)
    fig = go.Figure(go.Bar(x=L.AGING_ORDER, y=d["aging"], marker_color=[ "#1B5E20", "#388E3C", "#66BB6A", "#A5D6A7", "#C8E6C9"], text=[money(v) for v in d["aging"]], textposition="outside"))
    fig.update_yaxes(tickprefix="$", tickformat=",.0f"); c1.plotly_chart(style(fig.update_layout(title="Open AR by Aging Bucket")), width="stretch")
    fig = go.Figure(go.Pie(labels=L.STATUS_ORDER, values=d["status"], hole=0.6, sort=False,
                           marker=dict(colors=[GREEN, AMBER, RED, GREEN_L, "#8D6E63", "#9E9E9E", "#BDBDBD"])))
    c2.plotly_chart(style(fig.update_layout(title="Number of Invoices by Status")), width="stretch")
    fig = go.Figure(go.Bar(x=L.CHECK_ORDER, y=d["check"], marker_color=[GREEN, AMBER, RED, "#8D6E63"], text=[money(v) for v in d["check"]], textposition="outside"))
    fig.update_yaxes(tickprefix="$", tickformat=",.0f"); c3.plotly_chart(style(fig.update_layout(title="Billed by Deal Check")), width="stretch")

    c4, c5 = st.columns([1, 2])
    fig = go.Figure(go.Pie(labels=["Collected", "Still open"], values=[d["collection_rate"], max(0, 1 - d["collection_rate"])], hole=0.65, sort=False, marker=dict(colors=[GREEN, "#E0E0E0"]), textinfo="percent"))
    c4.plotly_chart(style(fig.update_layout(title="Collection Rate (%)")), width="stretch")
    bm = d["by_month"]
    fig = go.Figure([go.Bar(name="Billed", x=[x.strftime("%b %y") for x in bm.index], y=bm["Billed"], marker_color="#1B5E20"),
                     go.Bar(name="Collected", x=[x.strftime("%b %y") for x in bm.index], y=bm["Collected"], marker_color=GREEN_L)])
    fig.update_layout(barmode="group", title="Billed vs Collected by Month"); fig.update_yaxes(tickprefix="$", tickformat=",.0f")
    c5.plotly_chart(style(fig), width="stretch")
    st.markdown(f"**Open AR past due:** {d['past_due_share']:.1%}")
    st.progress(min(1.0, d["past_due_share"]))

    st.subheader("Invoices matching the filters")
    cols = ["Invoice ID", "Customer", "Deal ID", "Due Date", "Amount", "Open Balance", "Status", "Aging Bucket", "Invoice Check"]
    st.dataframe(d["rows"][cols], hide_index=True, width="stretch",
                 column_config={"Due Date": st.column_config.DateColumn(format="YYYY-MM-DD"), "Amount": st.column_config.NumberColumn(format="$%.0f"),
                                "Open Balance": st.column_config.NumberColumn(format="$%.0f")})

# ------------------------------------------------------------------ Revenue schedule
with tab_rev:
    st.subheader("Revenue schedule (Closed Won deals, prorated by day)")
    st.caption("Each cell = active days in the month / contract days x contract value. One-time fees are one-day contracts, so they land in the month they start.")
    if rr.empty:
        st.info("No Closed Won deals with start and end dates.")
    else:
        tot = rr.drop(columns=["Deal ID", "Customer", "Start Date", "End Date", "Contract Days"]).sum(numeric_only=True)
        total_row = {c: (tot.get(c, "") if c not in ("Deal ID", "Customer", "Start Date", "End Date", "Contract Days") else "") for c in rr.columns}
        total_row["Deal ID"] = "TOTAL"
        full = pd.concat([pd.DataFrame([total_row]), rr], ignore_index=True)
        for c in ("Start Date", "End Date"): full[c] = pd.to_datetime(full[c], errors="coerce")
        full["Contract Days"] = pd.to_numeric(full["Contract Days"], errors="coerce")
        mcols = [c.strftime("%b %y") for c in months]
        full = month_label(full)
        cfg = {c: st.column_config.NumberColumn(format="$%.0f") for c in mcols + ["Total Contract Value", "Revenue Before Window", "Revenue After Window"]}
        cfg["Start Date"] = st.column_config.DateColumn(format="YYYY-MM-DD"); cfg["End Date"] = st.column_config.DateColumn(format="YYYY-MM-DD")
        st.dataframe(full, hide_index=True, width="stretch", column_config=cfg)
        st.markdown(f"Every row adds up to its contract value (Check column is 0): {tie((rr['Check'].abs() < 0.01).all())}", unsafe_allow_html=True)
        fig = go.Figure(go.Bar(x=mcols, y=[rr[mm].sum() for mm in months], marker_color=GREEN, text=[money(rr[mm].sum()) for mm in months], textposition="outside"))
        fig.update_yaxes(tickprefix="$", tickformat=",.0f")
        st.plotly_chart(style(fig.update_layout(title="Recognized revenue by month"), 340), width="stretch")

# ------------------------------------------------------------------ Deferred revenue
with tab_def:
    st.subheader("Deferred revenue rollforward")
    st.caption("Beginning + new valid billings - recognized revenue = ending. Positive is deferred revenue (a liability). Negative is unbilled AR (an asset).")
    show = dfd.assign(Month=dfd["Month"].dt.strftime("%b %y"))
    cfg = {c: st.column_config.NumberColumn(format="$%.0f") for c in ["Beginning Deferred Revenue", "New Billings", "Revenue Recognized", "Ending Deferred Revenue"]}
    st.dataframe(show, hide_index=True, width="stretch", column_config=cfg)
    st.markdown(f"Ending equals billed to date minus revenue to date in every month: {tie((dfd['Check'] == 'Ties').all())}", unsafe_allow_html=True)
    x = show["Month"]
    fig = go.Figure([go.Bar(name="New Billings", x=x, y=dfd["New Billings"], marker_color=GREY), go.Bar(name="Revenue Recognized", x=x, y=dfd["Revenue Recognized"], marker_color=GREEN),
                     go.Scatter(name="Ending Deferred Revenue", x=x, y=dfd["Ending Deferred Revenue"], mode="lines+markers", line=dict(color=NAVY, width=3))])
    fig.update_layout(barmode="group", title="Billings, revenue and deferred balance"); fig.update_yaxes(tickprefix="$", tickformat=",.0f")
    st.plotly_chart(style(fig, 380), width="stretch")

# ------------------------------------------------------------------ Working sheets
HS_HELP = ["Stripe Billed", "Unbilled Variance", "Billing Status", "Contract Days", "Revenue Recognized To Date", "Billed Less Recognized", "Won #"]
ST_HELP = ["Clean Customer", "Duplicate Check", "Valid Flag", "Open Balance", "Days Past Due", "Aging Bucket", "Status", "HubSpot Contract Amount", "HubSpot Stage", "Invoice Check"]


def working(df, help_cols, caption):
    st.caption(caption + " Helper columns are shaded light blue.")
    st.dataframe(df.style.set_properties(subset=help_cols, **{"background-color": "#EEF3F9"}), hide_index=True, width="stretch")


with tab_hw:
    working(hw, HS_HELP, "Raw columns first (Deal ID to Product), then Stripe Billed to Won #.")
with tab_sw:
    working(sw, ST_HELP, "Raw columns first (Invoice ID to Deal ID), then Clean Customer to Invoice Check.")

# ------------------------------------------------------------------ Support
with tab_sup:
    st.subheader("Support: lists the dashboard filters read from")
    c = st.columns(4)
    c[0].dataframe(pd.DataFrame({"Month": ["All"] + [d.strftime("%b %y") for d in pd.period_range(sw["Invoice Date"].min(), sw["Invoice Date"].max(), freq="M").to_timestamp()]}), hide_index=True)
    c[1].dataframe(pd.DataFrame({"Status": ["All"] + L.STATUS_ORDER}), hide_index=True)
    c[2].dataframe(pd.DataFrame({"Aging Bucket": ["All", "Paid"] + L.AGING_ORDER + ["Excluded"]}), hide_index=True)
    c[3].dataframe(pd.DataFrame({"Deal Check": ["All"] + L.CHECK_ORDER + ["Excluded"]}), hide_index=True)
    st.subheader("Billed and collected by month")
    sup = sw[sw["Valid Flag"] == "Valid"].groupby(sw["Invoice Date"].dt.to_period("M").dt.to_timestamp()).agg(Billed=("Amount", "sum"), Collected=("Amount Paid", "sum"), Invoices=("Invoice ID", "count"))
    sup.index = sup.index.strftime("%b %y"); sup.index.name = "Month"
    st.dataframe(sup.reset_index(), hide_index=True, width="stretch", column_config={"Billed": st.column_config.NumberColumn(format="$%.0f"), "Collected": st.column_config.NumberColumn(format="$%.0f")})
    st.subheader("Rules used")
    st.markdown("""
- **Valid Flag:** Exclude if Stripe status is Void or Draft, or the invoice is a duplicate. Otherwise Valid.
- **Duplicate Check:** second or later invoice with the same customer (trimmed, lower case), amount, Deal ID and invoice month. Void rows are ignored.
- **Aging Bucket:** Excluded, Paid, Current, 1 to 30, 31 to 60, 61 to 90, Over 90 (days past due at the as-of date).
- **Invoice Check:** Excluded, Missing Deal ID, No matching deal, Deal not won, OK.
- **Revenue:** contract value spread by day from Start Date to End Date, Closed Won deals only.""")

# ------------------------------------------------------------------ Raw data
with tab_raw:
    st.subheader("Raw_HubSpot"); st.dataframe(hub, hide_index=True, width="stretch")
    st.subheader("Raw_Stripe"); st.dataframe(stp, hide_index=True, width="stretch")

# ------------------------------------------------------------------ Download
with tab_dl:
    st.subheader("Download")
    st.markdown("**Excel workbook with live formulas.** The same tabs and formulas as Case Study 3. Open it in Excel or Google Sheets and it calculates on open.")
    st.download_button("Download workbook (formulas)", excel_with_formulas(hub, stp, pd.Timestamp(asof), first_month), file_name="SaaS_Billing_Revenue_Hub.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.markdown("**Calculated values.** The tables from this app as plain numbers, no formulas.")
    st.download_button("Download calculated values", values_workbook(r), file_name="SaaS_Billing_Revenue_Hub_Values.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
