# rebuild trigger
import os
os.environ["MPLBACKEND"] = "Agg"  # Headless-Backend

import io
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import streamlit as st
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image
from reportlab.lib import colors

st.set_page_config(page_title="KMU Report Generator", page_icon="📄", layout="centered")

# --------- Benchmarks ---------
BENCH = {
    "Bäckerei": {"ebit_margin":0.08, "equity_ratio":0.40, "liquidity_ratio_2":1.00},
    "Gastro":   {"ebit_margin":0.06, "equity_ratio":0.35, "liquidity_ratio_2":0.90},
    "Bau":      {"ebit_margin":0.07, "equity_ratio":0.40, "liquidity_ratio_2":0.95}
}

# --------- Helpers ---------
def _get(df, key):
    row = df.loc[df['account'].str.lower()==key.lower()]
    return float(row['amount'].values[0]) if not row.empty else 0.0

def compute_kpis(balance_df: pd.DataFrame, pl_df: pd.DataFrame):
    b = balance_df.rename(columns={balance_df.columns[0]:'account', balance_df.columns[1]:'amount'})
    p = pl_df.rename(columns={pl_df.columns[0]:'account', pl_df.columns[1]:'amount'})
    for df in (b,p):
        df['account'] = df['account'].astype(str).str.strip()

    cash=_get(b,'cash'); rec=_get(b,'receivables'); inv=_get(b,'inventory')
    cl=_get(b,'current_liabilities'); fd=_get(b,'financial_debt')
    eq=_get(b,'equity'); ta=b['amount'].sum()

    rev=_get(p,'revenue'); cogs=_get(p,'cogs'); pers=_get(p,'personnel')
    depr=_get(p,'depr'); it=_get(p,'interest')
    other = p['amount'].sum() - (rev+cogs+pers+depr+it)
    opex_other = -other
    ebit = rev - cogs - pers - opex_other

    return dict(
        revenue=rev, ebit=ebit,
        ebit_margin=(ebit/rev) if rev else np.nan,
        liquidity_ratio_2=((cash+rec)/cl) if cl else np.nan,
        equity_ratio=(eq/ta) if ta else np.nan,
        working_capital=(cash+rec+inv)-cl,
        interest_coverage=(ebit/it) if it else np.nan,
        inventory_turnover=(cogs/inv) if inv else np.nan,
        total_assets=ta,
        cash=cash, receivables=rec, inventory=inv,
        current_liabilities=cl, equity=eq, financial_debt=fd
    )

def detect_industry(k, pl_df):
    df = pl_df.copy()
    df = df.rename(columns={df.columns[0]:'account', df.columns[1]:'amount'})
    rev = _get(df, "revenue")
    cogs_ratio = (_get(df, "cogs")/rev) if rev else np.nan
    inv_turn = k.get("inventory_turnover", np.nan)

    if not np.isnan(cogs_ratio):
        if cogs_ratio > 0.5: return "Bau"
        elif cogs_ratio < 0.35: return "Gastro"
        else: return "Bäckerei"
    if not np.isnan(inv_turn):
        if inv_turn < 8: return "Gastro"
        elif inv_turn < 12: return "Bäckerei"
        else: return "Bau"
    return "Bäckerei"

def narrative(k, bm):
    n=[]
    if not np.isnan(k["liquidity_ratio_2"]) and k["liquidity_ratio_2"]<1.0: n.append("Liquidität knapp. Forderungen und Lager prüfen.")
    if not np.isnan(k["equity_ratio"]) and k["equity_ratio"]<bm["equity_ratio"]: n.append("Eigenkapitalquote unter Branchenniveau. Kapitalstruktur stärken.")
    if not np.isnan(k["ebit_margin"]) and k["ebit_margin"]>=bm["ebit_margin"]: n.append("Profitabilität ≥ Branchenmedian. Expansion/Capex prüfen.")
    if not np.isnan(k.get("inventory_turnover", np.nan)) and "inventory_turnover" in bm:
        if k["inventory_turnover"] < {"Bäckerei":8,"Gastro":18,"Bau":6}[list(BENCH.keys())[list(BENCH.values()).index(bm)]]:
            n.append("Lagerumschlag tief. Sortiment/Disposition optimieren.")
    if not np.isnan(k["interest_coverage"]) and k["interest_coverage"]<2: n.append("Zinsdeckungsgrad < 2×. Verschuldung/Zinskosten adressieren.")
    if not n: n.append("Finanzlage stabil. Effizienz und Wachstum fokussieren.")
    return n

def fmt_pct(x): return "-" if (x is None or np.isnan(x)) else f"{x*100:.1f} %"
def fmt_num(x): return "-" if (x is None or np.isnan(x)) else f"{x:,.0f}".replace(",", "'")

def bar_png(val, bench, title, unit="%"):
    fig, ax = plt.subplots(figsize=(3.8,2.0))
    v = val*100 if not np.isnan(val) else 0
    b = bench*100 if not np.isnan(bench) else 0
    ax.bar(["Ist", "Branche p50"], [v, b])
    ax.set_ylabel(unit); ax.set_title(title); ax.grid(axis="y", alpha=0.3)
    buf = io.BytesIO(); fig.tight_layout(); fig.savefig(buf, format="png", dpi=200); plt.close(fig); buf.seek(0)
    return buf

def build_pdf(k, notes, org, ind):
    bm = BENCH[ind]
    s = getSampleStyleSheet()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, title=f"{org} – Jahresreport ({ind})")

    story = []
    story.append(Paragraph(f"<b>{org}</b>", s["Title"]))
    story.append(Paragraph(f"Jahresreport – Branche: <b>{ind}</b>", s["Heading2"]))
    story.append(Spacer(1, 8))
    story.append(Paragraph("<b>Management Summary</b>", s["Heading2"]))
    story.append(Paragraph(" ".join(notes), s["Normal"]))
    story.append(Spacer(1, 10))

    story.append(Paragraph("<b>Kennzahlen</b>", s["Heading2"]))
    data = [
        ["Umsatz", fmt_num(k["revenue"])+" CHF", "EBIT-Marge", fmt_pct(k["ebit_margin"])],
        ["EK-Quote", fmt_pct(k["equity_ratio"]), "Liquidität II", fmt_pct(k["liquidity_ratio_2"])],
        ["Working Capital", fmt_num(k["working_capital"])+" CHF", "Zinsdeckungsgrad", "-" if np.isnan(k["interest_coverage"]) else f"{k['interest_coverage']:.1f}×"],
    ]
    t = Table(data, colWidths=[130,130,130,130])
    t.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.25,colors.grey)]))
    story.append(t)
    story.append(Spacer(1,8))

    # Charts
    charts = [
        ("EBIT-Marge", k["ebit_margin"], bm["ebit_margin"], "%"),
        ("EK-Quote",   k["equity_ratio"], bm["equity_ratio"], "%"),
        ("Liquidität II", k["liquidity_ratio_2"], bm["liquidity_ratio_2"], "%"),
    ]
    for title, val, bench, unit in charts:
        img = bar_png(val, bench, title, unit)
        story.append(Image(img, width=260, height=140))
        story.append(Spacer(1,6))

    story.append(PageBreak())
    story.append(Paragraph("<b>Empfehlungen</b>", s["Heading2"]))
    for r in [
        "30-Tage Cash-Plan. DSO senken.",
        "Kostenstruktur vs. Branche prüfen, Margenziele setzen.",
        "Bei ≥ Branchenmedian Profitabilität Capex/Expansion priorisieren.",
        "Zinsrisiko prüfen. Refinanzierung/Absicherung evaluieren."
    ]:
        story.append(Paragraph("• "+r, s["Normal"]))

    doc.build(story); buf.seek(0)
    return buf

# --------- UI ---------
st.title("📄 KMU Report Generator")

with st.form("frm"):
    org = st.text_input("Firmenname", "Bäckerei Santschi GmbH")
    industry_choice = st.selectbox("Branche", ["Automatisch erkennen", "Bäckerei", "Gastro", "Bau"], index=0)
    file = st.file_uploader("Excel (.xlsx) mit Sheets: balance_sheet, profit_loss", type=["xlsx"])
    ok = st.form_submit_button("Report generieren")

if ok:
    if file is None:
        st.warning("Bitte eine Excel-Datei hochladen.")
        st.stop()
    try:
        xls = pd.ExcelFile(file)
        b = pd.read_excel(xls, "balance_sheet", header=None)
        p = pd.read_excel(xls, "profit_loss", header=None)

        k = compute_kpis(b, p)
        if industry_choice == "Automatisch erkennen":
            industry = detect_industry(k, p)
            st.info(f"Erkannte Branche: {industry}")
        else:
            industry = industry_choice

        notes = narrative(k, BENCH[industry])
        # KPIs im UI
        c1, c2, c3 = st.columns(3)
        c1.metric("Umsatz (CHF)", fmt_num(k["revenue"]))
        c2.metric("EBIT-Marge", fmt_pct(k["ebit_margin"]))
        c3.metric("EK-Quote", fmt_pct(k["equity_ratio"]))
        c1.metric("Liquidität II", fmt_pct(k["liquidity_ratio_2"]))
        c2.metric("Working Capital (CHF)", fmt_num(k["working_capital"]))
        c3.metric("Zinsdeckungsgrad", "-" if np.isnan(k["interest_coverage"]) else f"{k['interest_coverage']:.1f}×")

        st.subheader("Management Summary")
        for n in notes: st.write("• "+n)

        # PDF
        pdf = build_pdf(k, notes, org, industry)
        st.download_button("PDF herunterladen", data=pdf, file_name="report.pdf", mime="application/pdf")

    except Exception as e:
        st.error(f"Fehler: {e}")
        st.stop()
streamlit
pandas
openpyxl
numpy
reportlab
matplotlib
