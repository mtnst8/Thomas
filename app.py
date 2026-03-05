import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import io
import os
import zipfile

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

TEMPLATE_PATH = "saved_template.xlsx"

if "abca_map" not in st.session_state:
    st.session_state.abca_map = DEFAULT_ABCA.copy()
if "results" not in st.session_state:
    st.session_state.results = []

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
    raw = pd.read_excel(uploaded_file, header=None)
    for i, row in raw.iterrows():
        if any("product" in str(v).lower() for v in row.values):
            return i
    return 3

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
    return out.read(), len(df), round(df["BBL"].sum(), 4), unmapped

def make_zip(results):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for r in results:
            if not r.get("error") and not r.get("unmapped"):
                zf.writestr(r["output_name"], r["result_bytes"])
    buf.seek(0)
    return buf.read()

# ── UI ────────────────────────────────────────────────────────────────────────
st.title("🍺 WV BBL Tax Reporter")
st.markdown("*Mountain State Brewing Co. — Monthly BBL Tax Report Generator*")
st.markdown("---")

# ── Step 1: Template ──────────────────────────────────────────────────────────
st.subheader("1. WV Upload Template")

template_bytes = None

if os.path.exists(TEMPLATE_PATH):
    with open(TEMPLATE_PATH, "rb") as f:
        template_bytes = f.read()
    st.markdown('<div class="template-saved">✓ Saved template loaded — <em>__WV_Upload_Template.xlsx</em></div>', unsafe_allow_html=True)
    if st.button("Replace template"):
        os.remove(TEMPLATE_PATH)
        st.rerun()
else:
    new_template = st.file_uploader(
        "Upload __WV_Upload_Template.xlsx (only needed once)",
        type=["xlsx"],
        key="template"
    )
    if new_template:
        template_bytes = new_template.read()
        with open(TEMPLATE_PATH, "wb") as f:
            f.write(template_bytes)
        st.markdown('<div class="template-saved">✓ Template saved — won\'t need to upload again!</div>', unsafe_allow_html=True)

st.markdown("---")

# ── Step 2: Sales files ───────────────────────────────────────────────────────
st.subheader("2. Upload Monthly Sales File(s)")
sales_files = st.file_uploader(
    "Upload one or more monthly sales exports (e.g. Nov_25.xlsx)",
    type=["xlsx"],
    accept_multiple_files=True,
    key="sales"
)

st.markdown("---")

# ── ABCA manager ──────────────────────────────────────────────────────────────
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

# ── Step 3: Process ───────────────────────────────────────────────────────────
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
                "error": str(e),
            })

if st.session_state.results:
    successful = [r for r in st.session_state.results if not r.get("error") and not r.get("unmapped")]

    # Download All as ZIP (only shown when multiple successful files)
    if len(successful) > 1:
        st.download_button(
            label=f"⬇ Download All {len(successful)} Files as ZIP",
            data=make_zip(successful),
            file_name="BBL_Reports.zip",
            mime="application/zip",
            key="dl_zip"
        )
        st.markdown("&nbsp;")

    # Individual results
    for r in st.session_state.results:
        if r["error"]:
            st.markdown(f'<div class="error-box">✗ <b>{r["filename"]}</b> — Error: {r["error"]}</div>', unsafe_allow_html=True)
        elif r["unmapped"]:
            st.markdown(f'<div class="error-box">⚠️ <b>{r["filename"]}</b> — {len(r["unmapped"])} unmapped distributor(s): {", ".join(r["unmapped"])}<br>Add them in the ABCA manager above and re-run.</div>', unsafe_allow_html=True)
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
