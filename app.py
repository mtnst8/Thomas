import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
import io
import json
import os

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
.uploadbox {
    background: #241a08;
    border: 1px dashed #4a3510;
    border-radius: 4px;
    padding: 1rem;
    margin: 1rem 0;
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

# ── Helpers ───────────────────────────────────────────────────────────────────
def get_multiplier(product):
    p = str(product).lower()
    if "1/2 bbl" in p: return 0.5
    elif "1/6 bbl" in p: return 0.166667
    elif "can" in p: return 0.07258
    return None

def norm(name):
    return str(name).lower().split(":")[0].strip()

def detect_header_row(uploaded_file):
    """Find the row containing 'Product/Service' to use as header."""
    raw = pd.read_excel(uploaded_file, header=None)
    for i, row in raw.iterrows():
        if any("product" in str(v).lower() for v in row.values):
            return i
    return 3  # fallback

def process_file(uploaded_file, template_bytes):
    header_row = detect_header_row(uploaded_file)
    uploaded_file.seek(0)
    df = pd.read_excel(uploaded_file, header=header_row)
    df.columns = ["Product", "Trans_Date", "Num", "ABCA", "Customer", "Memo", "Quantity"]
    df = df[df["Product"] != "TOTAL"].dropna(subset=["Product", "Quantity"])
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce")
    df = df.dropna(subset=["Quantity"])
    df["Quantity"] = df["Quantity"].astype(int)
    df["Trans_Date"] = pd.to_datetime(df["Trans_Date"])
    df["Num"] = df["Num"].astype(str).str.strip()
    df["Multiplier"] = df["Product"].apply(get_multiplier)
    df["BBL"] = (df["Quantity"] * df["Multiplier"]).round(4)
    df["norm"] = df["Customer"].apply(norm)
    df["ABCA_Final"] = df["norm"].map(lambda x: st.session_state.abca_map.get(x, (None, None))[0])
    df["Licensee"] = df["norm"].map(lambda x: st.session_state.abca_map.get(x, (None, None))[1])

    unmapped = df[df["ABCA_Final"].isna()]["Customer"].unique().tolist()

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

    total_bbl = df["BBL"].sum()
    return out.read(), len(df), round(total_bbl, 4), unmapped

# ── UI ────────────────────────────────────────────────────────────────────────
st.title("🍺 WV BBL Tax Reporter")
st.markdown("*Mountain State Brewing Co. — Monthly BBL Tax Report Generator*")
st.markdown("---")

# Template upload
st.subheader("1. Upload WV Template")
template_file = st.file_uploader(
    "Upload __WV_Upload_Template.xlsx",
    type=["xlsx"],
    key="template"
)

st.markdown("---")

# Sales file upload
st.subheader("2. Upload Monthly Sales File(s)")
sales_files = st.file_uploader(
    "Upload one or more monthly sales exports (e.g. Nov_25.xlsx)",
    type=["xlsx"],
    accept_multiple_files=True,
    key="sales"
)

st.markdown("---")

# ABCA manager
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

# Process
st.subheader("3. Generate Reports")

if st.button("▶ Process Files", disabled=(not template_file or not sales_files)):
    template_bytes = template_file.read()
    for sf in sales_files:
        filename = sf.name.replace(".xlsx", "")
        output_name = f"{filename}_Final.xlsx"
        try:
            sf.seek(0)
            result_bytes, row_count, total_bbl, unmapped = process_file(sf, template_bytes)

            if unmapped:
                st.markdown(f'<div class="error-box">⚠️ <b>{filename}</b> — {len(unmapped)} unmapped distributor(s): {", ".join(unmapped)}<br>Add them above and re-run.</div>', unsafe_allow_html=True)
            else:
                st.markdown(f'<div class="result-box">✓ <b>{output_name}</b> — {row_count} rows &nbsp;|&nbsp; {total_bbl} BBL total</div>', unsafe_allow_html=True)
                st.download_button(
                    label=f"⬇ Download {output_name}",
                    data=result_bytes,
                    file_name=output_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_{filename}"
                )
        except Exception as e:
            st.markdown(f'<div class="error-box">✗ <b>{filename}</b> — Error: {str(e)}</div>', unsafe_allow_html=True)
elif not template_file or not sales_files:
    st.caption("Upload a template and at least one sales file to enable processing.")
