import streamlit as st
import pandas as pd
from openpyxl import load_workbook, Workbook
import io
import os
import zipfile
from pathlib import Path
from datetime import datetime

st.set_page_config(
    page_title="WV BBL Tax Reporter",
    page_icon="🍺",
    layout="centered"
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Libre+Baskerville:wght@400;700&family=Source+Code+Pro:wght@400;600&display=swap');

html, body, [class*="css"] {
    font-family: 'Libre Baskerville', serif;
}
.stApp {
    background-color: #1a1208;
    color: #f0e6c8;
}
h1 {
    font-family: 'Libre Baskerville', serif;
    color: #d4a843;
    font-size: 2rem;
    letter-spacing: 0.04em;
    border-bottom: 1px solid #4a3510;
    padding-bottom: 0.5rem;
}
h2, h3 {
    font-family: 'Libre Baskerville', serif;
    color: #d4a843;
}
.stButton > button {
    background-color: #d4a843;
    color: #1a1208;
    font-family: 'Libre Baskerville', serif;
    font-weight: 700;
    border: none;
    border-radius: 2px;
    padding: 0.5rem 1.5rem;
    letter-spacing: 0.05em;
}
.stButton > button:hover {
    background-color: #e8c060;
    color: #1a1208;
}
.stDownloadButton > button {
    background-color: #2a4a1a;
    color: #a0d080;
    font-family: 'Source Code Pro', monospace;
    border: 1px solid #4a7a2a;
    border-radius: 2px;
    width: 100%;
}
.stDownloadButton > button:hover {
    background-color: #3a6a2a;
}
.result-box {
    background: #241a08;
    border-left: 3px solid #d4a843;
    padding: 0.75rem 1rem;
    margin: 0.5rem 0;
    font-family: 'Source Code Pro', monospace;
    font-size: 0.85rem;
    color: #c8b888;
}
.error-box {
    background: #2a0808;
    border-left: 3px solid #d44343;
    padding: 0.75rem 1rem;
    margin: 0.5rem 0;
    font-family: 'Source Code Pro', monospace;
    font-size: 0.85rem;
    color: #d44343;
}
.template-saved {
    background: #0a2a1a;
    border-left: 3px solid #4a8a5a;
    padding: 0.5rem 1rem;
    margin: 0.5rem 0;
    font-family: 'Source Code Pro', monospace;
    font-size: 0.85rem;
    color: #80c880;
}
.stExpander {
    border: 1px solid #4a3510 !important;
    background: #241a08 !important;
}
label, .stFileUploader label {
    color: #c8b888 !important;
}
.stDataFrame {
    font-family: 'Source Code Pro', monospace;
    font-size: 0.8rem;
}
hr {
    border-color: #4a3510;
}
</style>
""", unsafe_allow_html=True)

# ── Constants ─────────────────────────────────────────────────────────────────
GAL_PER_BBL = 31                       # WV ABCA: 1 barrel = 31 gallons
FACILITY_CAPACITY_BBL = 5000           # Q3 — production capacity (constant)

# Template now ships WITH the repo so it survives Streamlit Cloud reboots.
# Commit __WV_Upload_Template.xlsx to the repo root (same folder as app.py).
TEMPLATE_PATH = Path(__file__).parent / "__WV_Upload_Template.xlsx"

# ── Default ABCA map ──────────────────────────────────────────────────────────
DEFAULT_ABCA = {
    "jefferson distributing co., inc.": ("02-T-001-000021", "Jefferson Distributing Co., Inc."),
    "northern eagle distributing":       ("14-T-001-000081", "Northern Eagle Distributing"),
    "beverage distributors":             ("17-T-001-000004", "Beverage Distributors"),
    "mountain eagle, inc.":              ("41-T-001-000028", "Mountain Eagle, Inc."),
    "carenbauer distributing":           ("35-T-001-000010", "Carenbauer Distributing"),
    "mona distributing":                 ("31-T-001-000027", "Mona Distributing"),
    "spriggs":                           ("54-T-001-017330", "Spriggs"),
}

if "abca_map" not in st.session_state:
    st.session_state.abca_map = DEFAULT_ABCA.copy()
if "results" not in st.session_state:
    st.session_state.results = []
if "eop" not in st.session_state:
    st.session_state.eop = None

# ── Helpers ───────────────────────────────────────────────────────────────────
def get_multiplier(product):
    p = str(product).lower().replace(" ", "")
    if "1/2bbl" in p: return 0.5
    elif "1/6bbl" in p: return 0.166667
    elif "can" in str(product).lower(): return 0.07258
    return None

def norm(name):
    return str(name).lower().split(":")[0].strip()

def detect_header_row(uploaded_file):
    raw = pd.read_excel(uploaded_file, header=None)
    for i, row in raw.iterrows():
        vals = [str(v).lower() for v in row.values]
        if any("transaction date" in v or "quantity" in v for v in vals):
            return i
    return 4

def load_template_bytes(override_file=None):
    """Session override wins; otherwise load the committed repo template."""
    if override_file is not None:
        return override_file.getvalue()
    if TEMPLATE_PATH.exists():
        return TEMPLATE_PATH.read_bytes()
    return None

def parse_sales_file(uploaded_file):
    """
    Shared parser used by BOTH the monthly tax report and the yearly EOP report.
    Returns a cleaned DataFrame with BBL (31-gal barrels) and Gallons per line item.
    """
    uploaded_file.seek(0)
    header_row = detect_header_row(uploaded_file)
    uploaded_file.seek(0)

    raw_data = pd.read_excel(uploaded_file, header=None, skiprows=header_row + 1)
    num_cols = raw_data.shape[1]

    if num_cols >= 10:
        # New 10-column QuickBooks format (Product/Service full name column present)
        raw_data.columns = ["Customer", "Trans_Date", "Trans_Type", "Num", "Product",
                            "Memo", "Quantity", "Sales_Price", "Amount", "Balance"]
        product_col = "Product"

        # Sub-customers (e.g. Northern Eagle Elkins) roll up to their ABCA-mapped
        # parent (Northern Eagle Distributing).
        customers = []
        last_top = None
        last_customer = None
        for _, row in raw_data.iterrows():
            val = row["Customer"]
            if pd.notna(val) and pd.isna(row["Trans_Date"]):
                s = str(val)
                if not s.lower().startswith("total"):
                    n = norm(s)
                    if n in st.session_state.abca_map:
                        last_top = s
                        last_customer = s
                    else:
                        last_customer = last_top if last_top else s
            customers.append(last_customer)
        raw_data["Customer_filled"] = customers
    else:
        # Old 7-column format (keg type in Memo column, customer in col 4)
        raw_data.columns = ["Customer", "Trans_Date", "Num", "ABCA", "Customer_orig", "Memo", "Quantity"]
        product_col = "Memo"
        raw_data["Customer_filled"] = raw_data["Customer_orig"].where(raw_data["Customer_orig"].notna()).ffill()

    df = raw_data.copy()
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce")
    df["Trans_Date"] = pd.to_datetime(df["Trans_Date"], errors="coerce")
    df = df.dropna(subset=["Trans_Date", "Quantity"])
    df = df[df["Quantity"] != 0]
    df["Quantity"] = df["Quantity"].astype(int)
    df["Num"] = df["Num"].astype(str).str.strip()

    df["Multiplier"] = df[product_col].apply(get_multiplier)
    df["BBL"] = (df["Quantity"] * df["Multiplier"]).round(4)
    df["Gallons"] = (df["BBL"] * GAL_PER_BBL).round(2)
    df["norm"] = df["Customer_filled"].apply(norm)
    df["ABCA_Final"] = df["norm"].map(lambda x: st.session_state.abca_map.get(x, (None, None))[0])
    df["Licensee"] = df["norm"].map(lambda x: st.session_state.abca_map.get(x, (None, None))[1])
    return df, product_col

def process_file(uploaded_file, template_bytes):
    df, product_col = parse_sales_file(uploaded_file)

    unmapped = [c for c in df[df["ABCA_Final"].isna()]["Customer_filled"].unique().tolist()
                if pd.notna(c) and str(c).strip().lower() != "nan"]

    wb = load_workbook(io.BytesIO(template_bytes))
    ws = wb["Section_2"]

    for i, row in enumerate(df.itertuples(index=False), 6):
        ws.cell(row=i, column=1, value=row.Trans_Date.date())
        ws.cell(row=i, column=1).number_format = "MM/DD/YYYY"
        ws.cell(row=i, column=2, value=row.Num)
        ws.cell(row=i, column=3, value=row.ABCA_Final)
        ws.cell(row=i, column=4, value=row.Licensee)
        ws.cell(row=i, column=5, value=round(row.BBL, 4))
        ws.cell(row=i, column=5).number_format = "0.0000"

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.read(), len(df), round(df["BBL"].sum(), 4), unmapped

def make_zip(results):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for r in results:
            if not r.get("error") and not r.get("unmapped"):
                zf.writestr(r["output_name"], r["result_bytes"])
    buf.seek(0)
    return buf.read()

# ── EOP (Production Estimate/Report) helpers ────────────────────────────────────
def summarize_year(df, brewpub_keywords):
    """
    Split the year's sales into channels:
      - Distributor: rows mapped to an ABCA license (your distributor map)
      - Brewpub:     non-distributor rows whose customer name matches a keyword
      - Self-dist:   all other non-distributor rows
    Returns gallon + barrel totals for each channel plus the 'other' breakdown.
    """
    def gb(d):
        return round(d["Gallons"].sum(), 2), round(d["BBL"].sum(), 4)

    dist = df[df["ABCA_Final"].notna()]
    other = df[df["ABCA_Final"].isna()].copy()

    kws = [k.strip().lower() for k in brewpub_keywords if k.strip()]
    if kws:
        mask = other["Customer_filled"].astype(str).str.lower().apply(
            lambda s: any(k in s for k in kws))
        brewpub = other[mask]
        selfdist = other[~mask]
    else:
        brewpub = other.iloc[0:0]
        selfdist = other

    total_g, total_b = gb(df)
    dist_g, dist_b = gb(dist)
    self_g, self_b = gb(selfdist)
    bp_g, bp_b = gb(brewpub)

    # Per-customer breakdown of the non-distributor pool (so you can spot the brewpub)
    other_breakdown = (
        other.assign(Customer=other["Customer_filled"].astype(str))
             .groupby("Customer", as_index=False)
             .agg(Gallons=("Gallons", "sum"), BBL=("BBL", "sum"))
             .sort_values("Gallons", ascending=False)
             .round({"Gallons": 2, "BBL": 4})
    )

    period_start = df["Trans_Date"].min()
    period_end = df["Trans_Date"].max()

    return {
        "total_g": total_g, "total_b": total_b,
        "dist_g": dist_g, "dist_b": dist_b,
        "self_g": self_g, "self_b": self_b,
        "bp_g": bp_g, "bp_b": bp_b,
        "other_breakdown": other_breakdown,
        "period_start": period_start, "period_end": period_end,
        "row_count": len(df),
    }

def build_eop_summary(v):
    """Build a clearly-labeled summary workbook of all 11 answers + channel split."""
    wb = Workbook()
    ws = wb.active
    ws.title = "EOP_Summary"
    rows = [
        ["WV Estimate/Report of Production", ""],
        ["Generated", datetime.now().strftime("%Y-%m-%d %H:%M")],
        ["Fiscal Year", v["fiscal_year"]],
        ["Data period (from file)", f"{v['period_start']:%m/%d/%Y} – {v['period_end']:%m/%d/%Y}"],
        ["", ""],
        ["Form line", "Answer"],
        ["1. Est. gallons produced this year", v["q1_gal"]],
        ["2. Est. barrels produced this year (gal / 31)", v["q2_bbl"]],
        ["3. Production capacity (barrels)", v["q3_bbl"]],
        ["   Production capacity (gallons)", v["q3_bbl"] * GAL_PER_BBL],
        ["4. Prior year total production (gallons)", v["q4_gal"]],
        ["5. Prior year total SALES volume (gallons)", v["total_g"]],
        ["   Prior year total sales volume (barrels)", v["total_b"]],
        ["6. Self-distributed (gallons)", v["self_g"]],
        ["7. Self-distributed (barrels)", v["self_b"]],
        ["8. Sold to WV distributors (gallons)", v["dist_g"]],
        ["9. Sold to WV distributors (barrels)", v["dist_b"]],
        ["10. Sold through brewpub (gallons)", v["bp_g"]],
        ["11. Sold through brewpub (barrels)", v["bp_b"]],
    ]
    for r in rows:
        ws.append(r)
    ws.column_dimensions["A"].width = 46
    ws.column_dimensions["B"].width = 28
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.read()

def build_history(prior_bytes, new_row):
    """Append this year's figures to a running history (dedup by Fiscal Year)."""
    cols = ["Fiscal Year", "Total Sales (gallons)", "Total Sales (barrels)",
            "Distributor (gallons)", "Self-distributed (gallons)",
            "Brewpub (gallons)", "Generated"]
    if prior_bytes:
        try:
            hist = pd.read_excel(io.BytesIO(prior_bytes))
        except Exception:
            hist = pd.DataFrame(columns=cols)
    else:
        hist = pd.DataFrame(columns=cols)

    hist = hist[hist["Fiscal Year"].astype(str) != str(new_row["Fiscal Year"])]
    hist = pd.concat([hist, pd.DataFrame([new_row])], ignore_index=True)
    hist = hist.sort_values("Fiscal Year").reset_index(drop=True)

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        hist.to_excel(w, index=False, sheet_name="Production_History")
    out.seek(0)
    return out.read(), hist

# ── UI ──────────────────────────────────────────────────────────────────────────
st.title("🍺 WV BBL Tax Reporter")
st.markdown("*Mountain State Brewing Co.*")
st.markdown("---")

tab_tax, tab_eop = st.tabs(["Monthly BBL Tax Report", "Production Estimate/Report"])

# ════════════════════════════════════════════════════════════════════════════════
#  TAB 1 — Monthly BBL Tax Report (existing functionality)
# ════════════════════════════════════════════════════════════════════════════════
with tab_tax:
    # ── Step 1: Template ─────────────────────────────────────────────────────────
    st.subheader("1. WV Upload Template")

    if TEMPLATE_PATH.exists():
        st.markdown('<div class="template-saved">✓ Template loaded from repo — '
                    '<em>__WV_Upload_Template.xlsx</em></div>', unsafe_allow_html=True)
        override = st.file_uploader(
            "Override template for this session only (optional)",
            type=["xlsx"], key="template_override"
        )
    else:
        st.markdown('<div class="error-box">⚠️ No template found in the repo. '
                    'Commit <em>__WV_Upload_Template.xlsx</em> to the repo root, or upload one '
                    'below for this session.</div>', unsafe_allow_html=True)
        override = st.file_uploader("Upload __WV_Upload_Template.xlsx", type=["xlsx"], key="template_override")

    template_bytes = load_template_bytes(override)

    st.markdown("---")

    # ── Step 2: Sales files ──────────────────────────────────────────────────────
    st.subheader("2. Upload Monthly Sales File(s)")
    sales_files = st.file_uploader(
        "Upload one or more monthly sales exports (e.g. Nov_25.xlsx)",
        type=["xlsx"],
        accept_multiple_files=True,
        key="sales"
    )

    st.markdown("---")

    # ── ABCA manager ─────────────────────────────────────────────────────────────
    with st.expander("⚙️ Manage Distributor ABCA Numbers"):
        st.markdown("**Current mappings:**")
        abca_df = pd.DataFrame([
            {"Distributor (lowercase key)": k, "ABCA License #": v[0], "Display Name": v[1]}
            for k, v in st.session_state.abca_map.items()
        ])
        st.dataframe(abca_df, use_container_width=True, hide_index=True)

        st.markdown("**Add a new distributor:**")
        col1, col2, col3 = st.columns(3)
        with col1:
            new_key = st.text_input("Customer name (as it appears in QuickBooks)", key="new_key")
        with col2:
            new_abca = st.text_input("ABCA License #", key="new_abca")
        with col3:
            new_display = st.text_input("Display Name", key="new_display")

        if st.button("Add Distributor"):
            if new_key and new_abca and new_display:
                st.session_state.abca_map[norm(new_key)] = (new_abca, new_display)
                st.success(f"Added: {new_display}")
                st.rerun()
            else:
                st.warning("Please fill in all three fields.")

    st.markdown("---")

    # ── Step 3: Process ──────────────────────────────────────────────────────────
    st.subheader("3. Generate Reports")

    if st.button("▶ Process Files", disabled=(not template_bytes or not sales_files)):
        st.session_state.results = []
        for sf in sales_files:
            filename = sf.name.replace(".xlsx", "")
            output_name = f"{filename}_Final.xlsx"
            try:
                sf.seek(0)
                result_bytes, row_count, total_bbl, unmapped = process_file(sf, template_bytes)
                st.session_state.results.append({
                    "filename": filename,
                    "output_name": output_name,
                    "result_bytes": result_bytes,
                    "row_count": row_count,
                    "total_bbl": total_bbl,
                    "unmapped": unmapped,
                    "error": None,
                })
            except Exception as e:
                st.session_state.results.append({
                    "filename": filename,
                    "output_name": output_name,
                    "unmapped": [],
                    "error": str(e),
                })

    if st.session_state.results:
        successful = [r for r in st.session_state.results if not r.get("error") and not r.get("unmapped")]

        if len(successful) > 1:
            st.download_button(
                label=f"⬇ Download All {len(successful)} Files as ZIP",
                data=make_zip(successful),
                file_name="BBL_Reports.zip",
                mime="application/zip",
                key="dl_zip"
            )
            st.markdown("&nbsp;")

        for r in st.session_state.results:
            if r.get("error"):
                st.markdown(f'<div class="error-box">✗ <b>{r["filename"]}</b> — Error: {r["error"]}</div>', unsafe_allow_html=True)
            elif r.get("unmapped"):
                unmapped = r["unmapped"]
                st.markdown(f'<div class="error-box">⚠️ <b>{r["filename"]}</b> — {len(unmapped)} unmapped distributor(s): {", ".join(str(u) for u in unmapped)}<br>Add them in the ABCA manager above and re-run.</div>', unsafe_allow_html=True)
            else:
                st.markdown(f'<div class="result-box">✓ <b>{r["output_name"]}</b> — {r["row_count"]} rows &nbsp;|&nbsp; {r["total_bbl"]} BBL total</div>', unsafe_allow_html=True)
                st.download_button(
                    label=f"⬇ Download {r['output_name']}",
                    data=r["result_bytes"],
                    file_name=r["output_name"],
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_{r['filename']}"
                )
    elif not template_bytes or not sales_files:
        st.caption("Upload a template and at least one sales file to enable processing.")

# ════════════════════════════════════════════════════════════════════════════════
#  TAB 2 — WV Estimate/Report of Production (ABCA-NONRETAIL-EOP)
# ════════════════════════════════════════════════════════════════════════════════
with tab_eop:
    st.subheader("WV Estimate/Report of Production")
    st.caption("Aggregates a full fiscal year (Jul 1 – Jun 30) of sales and splits it by channel. "
               "Barrels are 31 gallons (gallons ÷ 31), matching the ABCA form.")

    fy = st.text_input("Fiscal year label", value="", placeholder="e.g. 2025-2026", key="fy")

    year_file = st.file_uploader(
        "Upload the year's sales export (.xlsx)",
        type=["xlsx"], key="year_file"
    )

    brewpub_kw = st.text_input(
        "Brewpub / taproom name keyword(s) — comma-separated",
        value="", placeholder="e.g. taproom, brewpub, mountain state taproom",
        help="Non-distributor sales whose customer name contains any of these words are "
             "counted as brewpub (Q10/11). Everything else non-distributor counts as "
             "self-distribution (Q6/7). Leave blank to put all non-distributor sales in self-distribution."
    )

    prior_history = st.file_uploader(
        "Optional: prior production_history.xlsx (to keep a running year-to-year record)",
        type=["xlsx"], key="hist"
    )

    if st.button("▶ Analyze Year", disabled=(not year_file)):
        try:
            year_file.seek(0)
            df, _ = parse_sales_file(year_file)
            kws = [k for k in brewpub_kw.split(",")] if brewpub_kw else []
            summary = summarize_year(df, kws)
            summary["prior_history_bytes"] = prior_history.getvalue() if prior_history else None
            st.session_state.eop = summary
        except Exception as e:
            st.session_state.eop = None
            st.markdown(f'<div class="error-box">✗ Error: {e}</div>', unsafe_allow_html=True)

    v = st.session_state.eop
    if v:
        st.markdown(
            f'<div class="result-box">Period in file: '
            f'<b>{v["period_start"]:%m/%d/%Y}</b> – <b>{v["period_end"]:%m/%d/%Y}</b> '
            f'&nbsp;|&nbsp; {v["row_count"]} line items</div>',
            unsafe_allow_html=True
        )

        st.markdown("**Channel breakdown (from the file):**")
        breakdown = pd.DataFrame([
            {"Channel": "Sold to WV distributors", "Gallons": v["dist_g"], "Barrels": v["dist_b"]},
            {"Channel": "Self-distributed",        "Gallons": v["self_g"], "Barrels": v["self_b"]},
            {"Channel": "Brewpub",                 "Gallons": v["bp_g"],   "Barrels": v["bp_b"]},
            {"Channel": "TOTAL",                   "Gallons": v["total_g"],"Barrels": v["total_b"]},
        ])
        st.dataframe(breakdown, use_container_width=True, hide_index=True)

        with st.expander("Non-distributor customers (who's in self-dist vs brewpub?)"):
            st.caption("Use this to confirm which names should be tagged as the brewpub above.")
            st.dataframe(v["other_breakdown"], use_container_width=True, hide_index=True)

        st.markdown("---")
        st.markdown("**Production-side figures (you confirm these):**")
        cga, cgb = st.columns(2)
        with cga:
            q1_gal = st.number_input(
                "Q1 — Est. gallons produced this year",
                min_value=0.0, value=float(v["total_g"]), step=100.0, key="q1",
                help="Prefilled with total sales gallons from the file as a starting point. "
                     "Adjust to your production estimate."
            )
        with cgb:
            q4_gal = st.number_input(
                "Q4 — Prior year total production (gallons)",
                min_value=0.0, value=0.0, step=100.0, key="q4"
            )
        q2_bbl = round(q1_gal / GAL_PER_BBL, 2)
        st.markdown(
            f'<div class="result-box">'
            f'Q2 — Est. barrels produced (Q1 ÷ 31): <b>{q2_bbl}</b> bbl<br>'
            f'Q3 — Production capacity (constant): <b>{FACILITY_CAPACITY_BBL:,}</b> bbl '
            f'({FACILITY_CAPACITY_BBL * GAL_PER_BBL:,} gal)'
            f'</div>', unsafe_allow_html=True
        )

        values = {
            "fiscal_year": fy or f"{v['period_start']:%Y}-{v['period_end']:%Y}",
            "period_start": v["period_start"], "period_end": v["period_end"],
            "q1_gal": q1_gal, "q2_bbl": q2_bbl, "q3_bbl": FACILITY_CAPACITY_BBL, "q4_gal": q4_gal,
            "total_g": v["total_g"], "total_b": v["total_b"],
            "self_g": v["self_g"], "self_b": v["self_b"],
            "dist_g": v["dist_g"], "dist_b": v["dist_b"],
            "bp_g": v["bp_g"], "bp_b": v["bp_b"],
        }

        st.markdown("---")
        st.download_button(
            "⬇ Download EOP Summary (all 11 lines, labeled)",
            data=build_eop_summary(values),
            file_name=f"WV_Production_{values['fiscal_year']}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_eop"
        )

        new_row = {
            "Fiscal Year": values["fiscal_year"],
            "Total Sales (gallons)": v["total_g"],
            "Total Sales (barrels)": v["total_b"],
            "Distributor (gallons)": v["dist_g"],
            "Self-distributed (gallons)": v["self_g"],
            "Brewpub (gallons)": v["bp_g"],
            "Generated": datetime.now().strftime("%Y-%m-%d"),
        }
        hist_bytes, hist_df = build_history(v.get("prior_history_bytes"), new_row)
        st.markdown("**Year-to-year history (gallons):**")
        st.dataframe(hist_df, use_container_width=True, hide_index=True)
        st.download_button(
            "⬇ Download updated production_history.xlsx",
            data=hist_bytes,
            file_name="production_history.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_hist"
        )
        st.caption("Keep this file. Re-upload it next year (the box above) and it appends the new "
                   "year automatically — re-running the same fiscal year just overwrites that row.")
    elif not year_file:
        st.caption("Upload the year's sales export to enable analysis.")
