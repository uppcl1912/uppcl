# dashboard.py
# ============================================================
# UPPCL TOOLKIT — Single App
#   Page 1: 🔗 Billed + Unbilled File Merger
#   Page 2: ⚡ UPPCL Payment Dashboard
# ============================================================
import streamlit as st

st.set_page_config(
    page_title="UPPCL Toolkit",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ═══════════════════════════════════════════════════════════
#  🧭 NAVIGATION
# ═══════════════════════════════════════════════════════════
PAGES = {
    "🔗 Billed + Unbilled Merger": "merger",
    "⚡ UPPCL Payment Dashboard":   "dashboard",
}

with st.sidebar:
    st.markdown("## 🧭 Navigation")
    page_label = st.radio(
        "Page chunein",
        options=list(PAGES.keys()),
        index=0,
        label_visibility="collapsed",
        key="nav_page",
    )
    st.markdown("---")

page = PAGES[page_label]


# ═══════════════════════════════════════════════════════════
#  📦 PAGE 1 — MERGER
# ═══════════════════════════════════════════════════════════
def render_merger():
    import pandas as pd
    import io

    st.title("🔗 Billed + Unbilled File Merger")
    st.markdown("**Dono files upload karo — matching ke baad CSV ya Excel download karo.**")

    with st.sidebar:
        st.header("⚙️ Merger Settings")
        st.subheader("📝 Column Rename Rules")
        st.caption("Format: `unbilled_col:billed_col`")
        default_rules = """LAST_PAY_AMT:LAST_PAYMENT_AMOUNT
MTR_SRL_NO:METER_SERIAL_NBR
POLE:POLE_NO"""
        rename_rules_text = st.text_area(
            "Rename rules", value=default_rules, height=120, key="mg_rules"
        )

        rename_map = {}
        for line in rename_rules_text.strip().split("\n"):
            if ":" in line:
                old, new = line.split(":", 1)
                rename_map[old.strip()] = new.strip()

        st.markdown("---")
        add_source   = st.checkbox("Add 'source_file' column", value=True, key="mg_src")
        only_common  = st.checkbox("Sirf common columns rakho", value=False, key="mg_common")
        preview_rows = st.slider("Preview rows", 5, 200, 20, key="mg_prev")

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("📁 File 1 — Billed")
        billed_file = st.file_uploader(
            "Billed CSV/Excel upload karo",
            type=["csv", "xlsx", "xls"], key="mg_billed",
        )
    with col2:
        st.subheader("📁 File 2 — Unbilled")
        unbilled_file = st.file_uploader(
            "Unbilled CSV/Excel upload karo",
            type=["csv", "xlsx", "xls"], key="mg_unbilled",
        )

    @st.cache_data(show_spinner=False, max_entries=4)
    def read_file(file_bytes, file_name):
        name = file_name.lower()
        buf  = io.BytesIO(file_bytes)
        if name.endswith((".xlsx", ".xls")):
            df = pd.read_excel(buf)
        else:
            for enc in ["utf-8", "cp1252", "latin-1"]:
                try:
                    buf.seek(0)
                    df = pd.read_csv(buf, encoding=enc, low_memory=False)
                    break
                except UnicodeDecodeError:
                    continue
            else:
                raise Exception("Encoding detect nahi ho payi")
        df.columns = df.columns.str.strip()
        return df

    @st.cache_data(show_spinner=False)
    def build_merged(billed_bytes, billed_name, unbilled_bytes, unbilled_name,
                     rename_map_tuple, add_source, only_common):
        rename_map       = dict(rename_map_tuple)
        billed_df        = read_file(billed_bytes, billed_name)
        unbilled_df      = read_file(unbilled_bytes, unbilled_name)
        unbilled_renamed = unbilled_df.rename(columns=rename_map)

        if only_common:
            final_columns = [c for c in billed_df.columns if c in unbilled_renamed.columns]
        else:
            final_columns = list(billed_df.columns)
            for c in unbilled_renamed.columns:
                if c not in final_columns:
                    final_columns.append(c)

        billed_final   = billed_df.reindex(columns=final_columns)
        unbilled_final = unbilled_renamed.reindex(columns=final_columns)

        if add_source:
            billed_final.insert(0,   "source_file", "BILLED")
            unbilled_final.insert(0, "source_file", "UNBILLED")

        merged = pd.concat([billed_final, unbilled_final], ignore_index=True)
        return merged, billed_df, unbilled_renamed

    @st.cache_data(show_spinner=False)
    def to_csv_bytes(df):
        return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")

    @st.cache_data(show_spinner=False)
    def to_excel_bytes(df):
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            df.to_excel(w, index=False, sheet_name="Merged")
        return buf.getvalue()

    if billed_file and unbilled_file:
        try:
            billed_bytes   = billed_file.getvalue()
            unbilled_bytes = unbilled_file.getvalue()

            with st.spinner("Files process ho rahi hain..."):
                merged_df, billed_df, unbilled_renamed = build_merged(
                    billed_bytes, billed_file.name,
                    unbilled_bytes, unbilled_file.name,
                    tuple(sorted(rename_map.items())),
                    add_source, only_common,
                )

            st.success("✅ Dono files successfully read ho gayi!")

            info_col1, info_col2 = st.columns(2)
            with info_col1:
                st.metric("Billed Rows",    f"{len(billed_df):,}")
                st.metric("Billed Columns", len(billed_df.columns))
            with info_col2:
                st.metric("Unbilled Rows",    f"{len(unbilled_renamed):,}")
                st.metric("Unbilled Columns", len(unbilled_renamed.columns))

            with st.expander("🔍 Column Matching Details", expanded=False):
                b_cols = list(billed_df.columns)
                u_cols = list(unbilled_renamed.columns)
                matched       = [c for c in b_cols if c in u_cols]
                only_billed   = [c for c in b_cols if c not in u_cols]
                only_unbilled = [c for c in u_cols if c not in b_cols]
                c1, c2, c3 = st.columns(3)
                with c1:
                    st.markdown(f"**✅ Matched ({len(matched)})**")
                    st.write("\n".join(f"• {c}" for c in matched))
                with c2:
                    st.markdown(f"**📘 Only Billed ({len(only_billed)})**")
                    st.write("\n".join(f"• {c}" for c in only_billed))
                with c3:
                    st.markdown(f"**📗 Only Unbilled ({len(only_unbilled)})**")
                    st.write("\n".join(f"• {c}" for c in only_unbilled))

            st.markdown("---")
            st.subheader(f"👀 Preview (First {preview_rows} rows)")
            st.dataframe(merged_df.head(preview_rows), width='stretch')

            st.markdown("---")
            st.subheader("📊 Merge Summary")
            s1, s2, s3 = st.columns(3)
            s1.metric("Total Rows",    f"{len(merged_df):,}")
            s2.metric("Total Columns", len(merged_df.columns))
            s3.metric("Billed Rows",   f"{len(billed_df):,}")

            st.markdown("---")
            st.subheader("⬇️ Download Merged File")
            st.caption("Button click karne par file generate + download hogi")

            dl_col1, dl_col2 = st.columns(2)
            with dl_col1:
                if st.button("📥 Generate CSV", width='stretch', key="mg_gen_csv"):
                    st.session_state["mg_csv_bytes"] = to_csv_bytes(merged_df)
                if "mg_csv_bytes" in st.session_state:
                    st.download_button(
                        "⬇️ Download CSV",
                        data=st.session_state["mg_csv_bytes"],
                        file_name="merged_billed_unbilled.csv",
                        mime="text/csv",
                        width='stretch',
                        type="primary",
                        key="mg_dl_csv",
                    )
            with dl_col2:
                if st.button("📥 Generate Excel", width='stretch', key="mg_gen_xlsx"):
                    with st.spinner("Excel ban raha hai..."):
                        st.session_state["mg_excel_bytes"] = to_excel_bytes(merged_df)
                if "mg_excel_bytes" in st.session_state:
                    st.download_button(
                        "⬇️ Download Excel",
                        data=st.session_state["mg_excel_bytes"],
                        file_name="merged_billed_unbilled.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        width='stretch',
                        key="mg_dl_xlsx",
                    )
        except Exception as e:
            st.error(f"❌ Error: {e}")
    else:
        st.info("👆 Upar dono files upload karo (Billed + Unbilled)")

    st.markdown("---")
    st.caption("💡 Tip: Sidebar mein rename rules customize kar sakte ho.")


# ═══════════════════════════════════════════════════════════
#  📦 PAGE 2 — UPPCL PAYMENT DASHBOARD
# ═══════════════════════════════════════════════════════════
def render_dashboard():
    # ---------- IMPORTS ----------
    import pandas as pd
    import plotly.express as px
    import plotly.graph_objects as go
    from datetime import datetime
    import io

    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                    Table, TableStyle, PageBreak, KeepTogether)
    import xlsxwriter

    # ---------- GLOBAL SETTINGS ----------
    pd.set_option('display.float_format', '{:.2f}'.format)
    pd.set_option('display.precision', 2)
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', None)

    # ---------- CSS ----------
    st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700;900&family=JetBrains+Mono:wght@400;700&display=swap');

        html, body, [class*="css"] {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
        }

        .stApp {
            background:
                radial-gradient(ellipse at top left, rgba(0,82,155,0.15) 0%, transparent 50%),
                radial-gradient(ellipse at bottom right, rgba(245,124,0,0.1) 0%, transparent 50%),
                linear-gradient(135deg, #05080f 0%, #0a0e27 50%, #05080f 100%);
            background-attachment: fixed;
            min-height: 100vh;
        }

        .main .block-container {
            padding: 1rem 2rem 2rem 2rem;
            max-width: 100%;
        }

        .stApp::before {
            content: '';
            position: fixed;
            top: 0; left: 0; right: 0; bottom: 0;
            background-image:
                linear-gradient(rgba(0,112,192,0.03) 1px, transparent 1px),
                linear-gradient(90deg, rgba(0,112,192,0.03) 1px, transparent 1px);
            background-size: 50px 50px;
            pointer-events: none;
            z-index: 0;
            animation: gridMove 20s linear infinite;
        }

        @keyframes gridMove {
            0% { transform: translate(0, 0); }
            100% { transform: translate(50px, 50px); }
        }

        .stApp, .stApp p, .stApp span, .stApp label, .stApp div {
            color: #E8F0F8;
        }

        .hero-header {
            background:
                linear-gradient(135deg, rgba(0,51,102,0.85) 0%, rgba(0,82,155,0.75) 50%, rgba(245,124,0,0.65) 100%);
            backdrop-filter: blur(30px) saturate(180%);
            -webkit-backdrop-filter: blur(30px) saturate(180%);
            border: 1px solid rgba(255,255,255,0.18);
            border-radius: 24px;
            padding: 35px 45px;
            margin-bottom: 30px;
            box-shadow:
                0 25px 70px rgba(0,0,0,0.6),
                0 0 120px rgba(0,112,192,0.4),
                inset 0 1px 0 rgba(255,255,255,0.25);
            position: relative;
            overflow: hidden;
            animation: heroSlide 1s cubic-bezier(0.4, 0, 0.2, 1);
        }

        @keyframes heroSlide {
            from { opacity: 0; transform: translateY(-40px) scale(0.98); }
            to { opacity: 1; transform: translateY(0) scale(1); }
        }

        .hero-header::before {
            content: '';
            position: absolute;
            top: -50%; left: -50%;
            width: 200%; height: 200%;
            background: conic-gradient(from 0deg,
                transparent 0%,
                rgba(245,124,0,0.4) 15%,
                transparent 30%,
                rgba(0,112,192,0.3) 50%,
                transparent 65%,
                rgba(245,124,0,0.3) 85%,
                transparent 100%);
            animation: rotate 12s linear infinite;
            z-index: -1;
            filter: blur(40px);
        }

        @keyframes rotate {
            from { transform: rotate(0deg); }
            to { transform: rotate(360deg); }
        }

        .hero-title {
            font-size: 46px;
            font-weight: 900;
            background: linear-gradient(90deg, #FFFFFF, #F57C00, #4FC3F7, #FFFFFF);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
            text-shadow: 0 0 60px rgba(245,124,0,0.6);
            letter-spacing: 3px;
            margin: 0;
            animation: shimmer 4s linear infinite;
            background-size: 300% auto;
        }

        @keyframes shimmer {
            to { background-position: 300% center; }
        }

        .hero-subtitle {
            color: #B0C4DE;
            font-size: 15px;
            margin-top: 10px;
            letter-spacing: 1.5px;
            font-weight: 500;
        }

        .hero-badge {
            display: inline-block;
            background: linear-gradient(135deg, #F57C00, #FF9800);
            color: white;
            padding: 7px 18px;
            border-radius: 24px;
            font-size: 12px;
            font-weight: 800;
            margin-left: 15px;
            box-shadow: 0 4px 20px rgba(245,124,0,0.6);
            animation: pulseGlow 2s ease-in-out infinite;
            letter-spacing: 1px;
        }

        @keyframes pulseGlow {
            0%, 100% { box-shadow: 0 4px 20px rgba(245,124,0,0.6), 0 0 0 0 rgba(245,124,0,0.7); }
            50% { box-shadow: 0 4px 40px rgba(245,124,0,1), 0 0 0 12px rgba(245,124,0,0); }
        }

        .kpi-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
            gap: 18px;
            margin: 30px 0;
        }

        .kpi-card {
            background:
                linear-gradient(135deg, rgba(0,51,102,0.85) 0%, rgba(0,82,155,0.65) 100%);
            backdrop-filter: blur(20px) saturate(180%);
            border: 1px solid rgba(255,255,255,0.15);
            border-radius: 20px;
            padding: 22px 16px;
            text-align: center;
            position: relative;
            overflow: hidden;
            transition: all 0.5s cubic-bezier(0.4, 0, 0.2, 1);
            box-shadow: 0 10px 30px rgba(0,0,0,0.4);
            animation: fadeInUp 0.7s ease-out backwards;
        }

        .kpi-card:nth-child(1) { animation-delay: 0.05s; }
        .kpi-card:nth-child(2) { animation-delay: 0.1s; }
        .kpi-card:nth-child(3) { animation-delay: 0.15s; }
        .kpi-card:nth-child(4) { animation-delay: 0.2s; }
        .kpi-card:nth-child(5) { animation-delay: 0.25s; }
        .kpi-card:nth-child(6) { animation-delay: 0.3s; }
        .kpi-card:nth-child(7) { animation-delay: 0.35s; }
        .kpi-card:nth-child(8) { animation-delay: 0.4s; }

        @keyframes fadeInUp {
            from { opacity: 0; transform: translateY(40px) scale(0.95); }
            to { opacity: 1; transform: translateY(0) scale(1); }
        }

        .kpi-card::before {
            content: '';
            position: absolute;
            top: 0; left: 0; right: 0;
            height: 3px;
            background: linear-gradient(90deg, #F57C00, #FFC107, #F57C00, #FF9800);
            background-size: 300% auto;
            animation: shimmer 2.5s linear infinite;
        }

        .kpi-card:hover {
            transform: translateY(-10px) scale(1.04);
            box-shadow: 0 25px 60px rgba(0,0,0,0.6), 0 0 50px rgba(0,112,192,0.7);
            border-color: rgba(245,124,0,0.6);
        }

        .kpi-icon {
            font-size: 30px;
            margin-bottom: 10px;
            display: block;
            filter: drop-shadow(0 0 15px rgba(245,124,0,0.7));
        }

        .kpi-label {
            color: #B0C4DE;
            font-size: 10.5px;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 1.6px;
            margin-bottom: 10px;
        }

        .kpi-value {
            font-size: 26px;
            font-weight: 900;
            color: #FFFFFF;
            text-shadow: 0 0 25px rgba(0,112,192,0.9);
            line-height: 1.1;
            word-break: break-word;
            font-family: 'JetBrains Mono', 'Inter', monospace;
            letter-spacing: -0.5px;
        }

        .kpi-delta {
            font-size: 11px;
            margin-top: 8px;
            color: #4CAF50;
            font-weight: 700;
        }

        [data-testid="stSidebar"] {
            background: linear-gradient(180deg, #05080f 0%, #0a0e27 100%);
            border-right: 1px solid rgba(0,112,192,0.3);
            box-shadow: 4px 0 30px rgba(0,0,0,0.5);
        }

        [data-testid="stSidebar"] * {
            color: #E8F0F8 !important;
        }

        [data-testid="stSidebar"] .stMarkdown h1,
        [data-testid="stSidebar"] .stMarkdown h2,
        [data-testid="stSidebar"] .stMarkdown h3,
        [data-testid="stSidebar"] .stMarkdown h4 {
            background: linear-gradient(90deg, #F57C00, #FFC107);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            font-weight: 900;
            letter-spacing: 1.5px;
            text-transform: uppercase;
        }

        [data-testid="stSidebar"] hr {
            border-color: rgba(0,112,192,0.3);
        }

        [data-testid="stSidebar"] .stSelectbox > div > div,
        [data-testid="stSidebar"] .stMultiSelect > div > div {
            background: rgba(0,51,102,0.5) !important;
            border: 1px solid rgba(0,112,192,0.5) !important;
            color: white !important;
            border-radius: 10px !important;
        }

        .stTabs [data-baseweb="tab-list"] {
            gap: 10px;
            background: rgba(0,51,102,0.35);
            backdrop-filter: blur(15px);
            padding: 10px;
            border-radius: 18px;
            border: 1px solid rgba(0,112,192,0.35);
            overflow-x: auto;
            flex-wrap: nowrap;
            margin-bottom: 25px;
            box-shadow: inset 0 2px 10px rgba(0,0,0,0.3);
        }

        .stTabs [data-baseweb="tab"] {
            background: rgba(255,255,255,0.04);
            border: 1px solid rgba(0,112,192,0.25);
            border-radius: 12px;
            padding: 12px 20px;
            font-weight: 700;
            color: #B0C4DE;
            font-size: 13px;
            white-space: nowrap;
            transition: all 0.35s cubic-bezier(0.4, 0, 0.2, 1);
            letter-spacing: 0.5px;
        }

        .stTabs [data-baseweb="tab"]:hover {
            background: rgba(0,112,192,0.35);
            color: white;
            transform: translateY(-3px);
            box-shadow: 0 6px 20px rgba(0,112,192,0.4);
        }

        .stTabs [aria-selected="true"] {
            background: linear-gradient(135deg, #00529B 0%, #F57C00 100%) !important;
            color: white !important;
            box-shadow: 0 6px 25px rgba(245,124,0,0.6);
            border-color: rgba(245,124,0,0.8) !important;
            transform: translateY(-2px);
        }

        [data-testid="stMetric"] {
            background: linear-gradient(135deg, rgba(0,51,102,0.75), rgba(0,82,155,0.55));
            backdrop-filter: blur(15px);
            border: 1px solid rgba(0,112,192,0.4);
            border-left: 5px solid #F57C00;
            padding: 18px;
            border-radius: 14px;
            box-shadow: 0 6px 20px rgba(0,0,0,0.4);
            transition: all 0.35s ease;
        }

        [data-testid="stMetric"]:hover {
            transform: translateY(-4px);
            box-shadow: 0 12px 35px rgba(0,0,0,0.5), 0 0 25px rgba(0,112,192,0.4);
            border-left-color: #FFC107;
        }

        [data-testid="stMetricLabel"] {
            color: #B0C4DE !important;
            font-weight: 700;
            font-size: 11px;
            text-transform: uppercase;
            letter-spacing: 1.5px;
        }

        [data-testid="stMetricValue"] {
            color: white !important;
            font-size: 26px;
            font-weight: 900;
            text-shadow: 0 0 20px rgba(0,112,192,0.7);
            font-family: 'JetBrains Mono', monospace;
        }

        .stButton > button, .stDownloadButton > button {
            background: linear-gradient(135deg, #00529B 0%, #003366 100%);
            color: white;
            border: 1px solid rgba(245,124,0,0.5);
            font-weight: 800;
            border-radius: 12px;
            padding: 14px 24px;
            width: 100%;
            transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
            box-shadow: 0 6px 20px rgba(0,0,0,0.4);
            text-transform: uppercase;
            letter-spacing: 1.5px;
            font-size: 12px;
        }

        .stButton > button:hover, .stDownloadButton > button:hover {
            background: linear-gradient(135deg, #F57C00 0%, #FF9800 100%);
            transform: translateY(-4px) scale(1.02);
            box-shadow: 0 12px 40px rgba(245,124,0,0.7);
            border-color: #FFC107;
        }

        .stDataFrame {
            border: 1px solid rgba(0,112,192,0.4);
            border-radius: 14px;
            overflow: hidden;
            box-shadow: 0 10px 30px rgba(0,0,0,0.4);
        }

        .streamlit-expanderHeader {
            background: rgba(0,51,102,0.6) !important;
            border-radius: 12px !important;
            border: 1px solid rgba(0,112,192,0.4) !important;
            color: white !important;
            font-weight: 700 !important;
        }

        .stProgress > div > div > div > div {
            background: linear-gradient(90deg, #F57C00, #FFC107, #F57C00);
            background-size: 300% auto;
            animation: shimmer 2s linear infinite;
            border-radius: 10px;
        }

        .stAlert {
            background: rgba(0,51,102,0.7) !important;
            backdrop-filter: blur(15px);
            border-radius: 12px;
            border-left: 5px solid #F57C00 !important;
            color: white !important;
        }

        .stApp h1, .stApp h2, .stApp h3 {
            color: white !important;
            text-shadow: 0 0 25px rgba(0,112,192,0.6);
            letter-spacing: 0.5px;
            font-weight: 800;
        }

        .stApp h3 {
            border-bottom: 2px solid rgba(245,124,0,0.6);
            padding-bottom: 12px;
            display: inline-block;
        }

        [data-testid="stFileUploader"] {
            background: rgba(0,51,102,0.4);
            border: 2px dashed rgba(0,112,192,0.6);
            border-radius: 16px;
            padding: 20px;
            transition: all 0.3s;
        }

        [data-testid="stFileUploader"]:hover {
            border-color: #F57C00;
            box-shadow: 0 0 30px rgba(245,124,0,0.3);
        }

        ::-webkit-scrollbar { width: 12px; height: 12px; }
        ::-webkit-scrollbar-track { background: rgba(0,51,102,0.3); border-radius: 6px; }
        ::-webkit-scrollbar-thumb {
            background: linear-gradient(135deg, #00529B, #F57C00);
            border-radius: 6px;
        }

        #MainMenu { visibility: hidden; }
        footer { visibility: hidden; }
        header { visibility: hidden; }

        .uppcl-footer {
            background: linear-gradient(135deg, rgba(0,51,102,0.95), rgba(0,82,155,0.75));
            backdrop-filter: blur(20px);
            border: 1px solid rgba(0,112,192,0.4);
            border-radius: 20px;
            padding: 30px;
            margin: 40px 0 15px 0;
            text-align: center;
            box-shadow: 0 20px 50px rgba(0,0,0,0.4);
        }

        .uppcl-footer h4 {
            color: #F57C00 !important;
            margin: 0 0 12px 0;
            font-size: 20px;
            letter-spacing: 3px;
            font-weight: 900;
        }

        .uppcl-footer p {
            color: #B0C4DE;
            margin: 6px 0;
            font-size: 13px;
        }

        .live-dot {
            display: inline-block;
            width: 10px;
            height: 10px;
            background: #4CAF50;
            border-radius: 50%;
            margin-right: 8px;
            box-shadow: 0 0 15px #4CAF50;
            animation: livePulse 1.5s ease-in-out infinite;
        }

        @keyframes livePulse {
            0%, 100% { opacity: 1; transform: scale(1); }
            50% { opacity: 0.5; transform: scale(1.3); }
        }

        @media screen and (max-width: 768px) {
            .hero-title { font-size: 26px; letter-spacing: 1.5px; }
            .hero-subtitle { font-size: 12px; }
            .hero-badge { font-size: 10px; padding: 5px 12px; margin-left: 8px; }
            .hero-header { padding: 20px 22px; border-radius: 16px; }
            .kpi-value { font-size: 20px; }
            .kpi-icon { font-size: 22px; }
            .kpi-label { font-size: 9.5px; }
            .kpi-card { padding: 16px 10px; border-radius: 14px; }
            .kpi-grid { gap: 10px; }
            .stTabs [data-baseweb="tab"] { font-size: 11px; padding: 8px 12px; }
            [data-testid="stMetricValue"] { font-size: 20px; }
            .main .block-container { padding: 0.5rem 0.8rem; }
        }
    </style>
    """, unsafe_allow_html=True)

    # ---------- CONFIG ----------
    BILL_DATE_COL   = "BILL_DATE"
    DATE_COL        = "LAST_PAY_DATE"
    CATEGORY_COL    = "TARIFF_TYPE"
    AMOUNT_COL      = "LAST_PAYMENT_AMOUNT"
    LOAD_COL        = "SANCTION_LOAD"
    LOAD_UOM_COL    = "SANCTION_LOAD_UOM"
    CA_COL          = "CA"
    ARREAR_COL      = "AMOUNT_PAYABLE"
    SDO_CODE_COL    = "SDO_CODE"
    SDO_NAME_COL    = "SDO_NAME"
    SUPPLY_TYPE_COL = "SUPPLY_TYPE"
    ACCT_COL        = "ACCT_ID"
    NAME_COL        = "NAME"
    MOBILE_COL      = "MOBILE_NO"
    SCNO_COL        = "SCNO"
    DUE_DATE_COL    = "DUE_DATE"

    USECOLS = [
        BILL_DATE_COL, DATE_COL, CATEGORY_COL, AMOUNT_COL, CA_COL,
        ARREAR_COL,
        LOAD_COL, LOAD_UOM_COL, SUPPLY_TYPE_COL,
        SDO_CODE_COL, SDO_NAME_COL,
        ACCT_COL, NAME_COL, MOBILE_COL, SCNO_COL,
        "PAYMENT_MODE", "DIV_NAME",
    ]

    TARIFF_MERGE = {"HV1": "HV", "HV2": "HV"}
    HV_TARIFFS = {"HV1", "HV2"}
    OTHERS_TARIFFS = {"LMV3", "LMV4", "LMV5", "LMV7", "LMV8", "LMV10"}
    LMV4_EXCEPTION_TARIFF       = "LMV4"
    LMV4_EXCEPTION_SUPPLY_TYPES = {"46", "47"}

    LOAD_ORDER = [
        "HV CONNECTION",
        ">= 10 KW/KVA/BHP",
        "5-9 KW/KVA/BHP",
        "< 5 KW/KVA/BHP",
        "Others",
    ]

    CRORE = 1_00_00_000
    EFFICIENCY_CMAP = "RdYlGn"
    TURNUP_CMAP     = "RdYlGn"

    SDO_NAME_MAP = {
        "SDO3419221": "SDO Lal Fatak",
        "SDO3419222": "SDO Faridpur",
        "SDO3419229": "SDO Faridpur",
        "SDO3419223": "SDO Izzat Nagar",
        "SDO3419224": "SDO Bhuta",
    }

    def sdo_display(code):
        if pd.isna(code):
            return code
        c = str(code).strip()
        name = SDO_NAME_MAP.get(c)
        return f"{name} ({c})" if name else c

    def sdo_short(code):
        if pd.isna(code):
            return code
        c = str(code).strip()
        return SDO_NAME_MAP.get(c, c)

    TARGET_GROUPS = [
        {"label": "SDO Lal Fatak",
         "codes": ["SDO3419221"], "target_cr": 11.36},
        {"label": "SDO Faridpur",
         "codes": ["SDO3419222", "SDO3419229"], "target_cr": 5.88},
        {"label": "SDO Izzat Nagar",
         "codes": ["SDO3419223"], "target_cr": 8.02},
        {"label": "SDO Bhuta",
         "codes": ["SDO3419224"], "target_cr": 2.40},
    ]

    INDEPENDENT_FEEDERS_TARGET_CR = 7.65
    TOTAL_TARGET_CR = round(sum(g["target_cr"] for g in TARGET_GROUPS) + INDEPENDENT_FEEDERS_TARGET_CR, 2)

    # ---------- THEME ----------
    CHART_LAYOUT = dict(
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,51,102,0.15)',
        font=dict(color='#E8F0F8', family='Inter, Segoe UI, sans-serif', size=12),
        xaxis=dict(gridcolor='rgba(0,112,192,0.2)', linecolor='rgba(0,112,192,0.5)'),
        yaxis=dict(gridcolor='rgba(0,112,192,0.2)', linecolor='rgba(0,112,192,0.5)'),
        margin=dict(l=40, r=20, t=50, b=40),
        hoverlabel=dict(bgcolor='#00529B', font_color='white', font_size=12),
    )

    def apply_theme(fig, height=400, title=None):
        fig.update_layout(**CHART_LAYOUT, height=height)
        if title:
            fig.update_layout(title=dict(text=title, font=dict(size=16, color='#F57C00')))
        return fig

    # ---------- ARROW SAFE HELPER ----------
    def _arrow_safe(df):
        """
        Streamlit 1.64 + pyarrow 24 mein mixed-type columns (jaise int + "TOTAL")
        Arrow serialization error dete hain. Yeh helper har object-dtype column
        ko string mein force kar deta hai, taaki TOTAL row ke saath bhi
        dataframe safely render ho jaye.
        """
        df = df.copy()
        for col in df.columns:
            # Object dtype columns ko string mein convert karo
            if df[col].dtype == object:
                df[col] = df[col].astype(str)
            # Mixed int/str wale numeric columns ko bhi string karo
            # (jab TOTAL row ki wajah se dtype object ho gayi ho)
            elif df[col].dtype == "int64" and "TOTAL" in df[col].astype(str).values:
                df[col] = df[col].astype(str)
        return df

    # ---------- LOAD FILE ----------
    @st.cache_data(show_spinner=False)
    def load_file(file_bytes: bytes, filename: str) -> pd.DataFrame:
        name = filename.lower()
        if name.endswith(".csv"):
            head = pd.read_csv(io.BytesIO(file_bytes), nrows=0)
            cols = [c for c in head.columns if c in USECOLS]
            return pd.read_csv(io.BytesIO(file_bytes),
                               usecols=cols if cols else None,
                               engine="c", low_memory=False)
        if name.endswith((".xlsx", ".xls")):
            xl = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
            sheet = xl.sheet_names[0]
            head = pd.read_excel(xl, sheet_name=sheet, nrows=0)
            cols = [c for c in head.columns if c in USECOLS]
            return pd.read_excel(xl, sheet_name=sheet,
                                 usecols=cols if cols else None,
                                 engine="openpyxl")
        st.error("Sirf .xlsx / .xls / .csv supported hain.")
        return pd.DataFrame()

    # ---------- PREPARE ----------
    @st.cache_data(show_spinner="Data process ho raha hai...")
    def prepare(df: pd.DataFrame, merge_tuple: tuple):
        df = df.copy()

        df[DATE_COL]      = pd.to_datetime(df[DATE_COL], errors="coerce", format="mixed")
        df[BILL_DATE_COL] = pd.to_datetime(df[BILL_DATE_COL], errors="coerce", format="mixed")

        if ARREAR_COL in df.columns:
            df[ARREAR_COL] = pd.to_numeric(df[ARREAR_COL], errors="coerce").fillna(0)
        else:
            df[ARREAR_COL] = 0.0

        if CA_COL in df.columns:
            df[CA_COL] = pd.to_numeric(df[CA_COL], errors="coerce").fillna(0)

        tariff = df[CATEGORY_COL].astype(str).str.strip().str.upper()
        supply = (df[SUPPLY_TYPE_COL].astype(str).str.strip().str.upper()
                  if SUPPLY_TYPE_COL in df.columns
                  else pd.Series("", index=df.index))
        load   = df[LOAD_COL] if LOAD_COL in df.columns else pd.Series(pd.NA, index=df.index)

        is_hv = tariff.isin(HV_TARIFFS)
        lmv4_supply = {s.upper() for s in LMV4_EXCEPTION_SUPPLY_TYPES}
        is_lmv4_exception = ((tariff == LMV4_EXCEPTION_TARIFF) &
                             (supply.isin(lmv4_supply)))
        is_others   = tariff.isin(OTHERS_TARIFFS) & ~is_lmv4_exception
        is_eligible = ~is_hv & ~is_others

        load_cat = pd.Series([None] * len(df), index=df.index, dtype="object")
        load_cat[is_hv]                                   = "HV CONNECTION"
        load_cat[is_others]                               = "Others"
        load_cat[is_eligible & (load >= 10)]              = ">= 10 KW/KVA/BHP"
        load_cat[is_eligible & (load >= 5) & (load < 10)] = "5-9 KW/KVA/BHP"
        load_cat[is_eligible & (load < 5) & load.notna()] = "< 5 KW/KVA/BHP"

        df["LOAD_CATEGORY"] = load_cat

        if SDO_CODE_COL in df.columns:
            df["SDO_NAME"] = df[SDO_CODE_COL].apply(sdo_short)
        else:
            df["SDO_NAME"] = "Unknown"

        if merge_tuple:
            df[CATEGORY_COL] = df[CATEGORY_COL].replace(dict(merge_tuple))

        df_paid = df.dropna(subset=[DATE_COL]).copy()
        df_paid = df_paid[df_paid[AMOUNT_COL].fillna(0) > 0]
        df_paid["_YM"] = df_paid[DATE_COL].dt.to_period("M").astype(str)

        df_bill = df.dropna(subset=[BILL_DATE_COL]).copy()
        df_bill["_YM"] = df_bill[BILL_DATE_COL].dt.to_period("M").astype(str)

        return df_paid, df_bill

    # ---------- SUMMARY HELPERS ----------
    def make_summary(df_paid, df_bill, group_col, order=None):
        if not df_bill.empty and CA_COL in df_bill.columns:
            agg_dict = {
                "CA_Total":      (CA_COL, "sum"),
                "Bills_Total":   (CA_COL, "size"),
            }
            if ARREAR_COL in df_bill.columns:
                agg_dict["Arrear_Total"] = (ARREAR_COL, "sum")
            else:
                agg_dict["Arrear_Total"] = (CA_COL, lambda x: 0)
            ca = (df_bill.groupby(group_col, as_index=False, observed=True)
                         .agg(**agg_dict))
        else:
            ca = pd.DataFrame({group_col: [], "CA_Total": [], "Bills_Total": [],
                               "Arrear_Total": []})

        if not df_paid.empty:
            paid = (df_paid.groupby(group_col, as_index=False, observed=True)
                           .agg(Total_Paid=(AMOUNT_COL, "sum"),
                                Txns=(AMOUNT_COL, "size")))
        else:
            paid = pd.DataFrame({group_col: [], "Total_Paid": [], "Txns": []})

        s = pd.merge(ca, paid, on=group_col, how="outer").fillna(0)
        s["CA_Cr"]        = (s["CA_Total"] / CRORE).round(2)
        s["Arrear_Cr"]    = (s["Arrear_Total"] / CRORE).round(2)
        s["Total_Cr"]     = (s["Total_Paid"] / CRORE).round(2)
        s["Txns"]         = s["Txns"].astype(int)
        s["Bills_Total"]  = s["Bills_Total"].astype(int)
        s["Efficiency_%"] = (s["Total_Paid"] /
                             s["CA_Total"].replace(0, pd.NA) * 100).round(2).fillna(0)
        s["Turnup_%"]     = (s["Txns"] /
                             s["Bills_Total"].replace(0, pd.NA) * 100).round(2).fillna(0)

        if order:
            s["__o"] = s[group_col].map({c: i for i, c in enumerate(order)})
            s = s.sort_values("__o", na_position="last").drop(columns="__o")
        else:
            s = s.sort_values("Total_Paid", ascending=False)
        return s

    def _total_row(group_col, s):
        tot_ca    = round(s["CA_Cr"].sum(), 2)
        tot_arr   = round(s["Arrear_Cr"].sum(), 2)
        tot_paid  = round(s["Total_Cr"].sum(), 2)
        tot_txns  = int(s["Txns"].sum())
        tot_bills = int(s["Bills_Total"].sum())
        tot_eff   = round((tot_paid / tot_ca * 100) if tot_ca else 0, 2)
        tot_turn  = round((tot_txns / tot_bills * 100) if tot_bills else 0, 2)
        return {
            group_col:       "TOTAL",
            "CA_Cr":         tot_ca,
            "Arrear_Cr":     tot_arr,
            "Total_Cr":      tot_paid,
            "Txns":          tot_txns,
            "Bills_Total":   tot_bills,
            "Efficiency_%":  tot_eff,
            "Turnup_%":      tot_turn,
        }

    def show_4col_table(df_paid, df_bill, group_col, title, key,
                        order=None, chart=True):
        st.subheader(title)
        if df_paid.empty and df_bill.empty:
            st.info("Koi data nahi.")
            return

        s = make_summary(df_paid, df_bill, group_col, order=order)
        total = _total_row(group_col, s)

        display = s[[group_col, "CA_Cr", "Total_Cr", "Arrear_Cr", "Txns",
                     "Bills_Total", "Efficiency_%", "Turnup_%"]].rename(columns={
            "CA_Cr":         "CA (Cr.)",
            "Total_Cr":      "Paid (Cr.)",
            "Arrear_Cr":     "Total Outstanding (Cr.)",
            "Txns":          "Paid Count",
            "Bills_Total":   "Bills",
            "Efficiency_%":  "Eff %",
            "Turnup_%":      "Turn-up %",
        })

        totals = pd.DataFrame([{
            group_col:      "TOTAL",
            "CA (Cr.)":     total["CA_Cr"],
            "Paid (Cr.)":   total["Total_Cr"],
            "Total Outstanding (Cr.)": total["Arrear_Cr"],
            "Paid Count":   total["Txns"],
            "Bills":        total["Bills_Total"],
            "Eff %":        total["Efficiency_%"],
            "Turn-up %":    total["Turnup_%"],
        }])
        display_full = pd.concat([display, totals], ignore_index=True)
        # 🛠️ ARROW FIX — mixed-type columns ko string banao
        display_full = _arrow_safe(display_full)

        col1, col2 = st.columns([1.7, 1])
        with col1:
            st.dataframe(
                display_full.style
                  .format({
                      "CA (Cr.)":     "₹ {:,.2f}",
                      "Paid (Cr.)":   "₹ {:,.2f}",
                      "Total Outstanding (Cr.)": "₹ {:,.2f}",
                      "Paid Count":   "{:,}",
                      "Bills":        "{:,}",
                      "Eff %":        "{:.2f}%",
                      "Turn-up %":    "{:.2f}%",
                  })
                  .apply(lambda r: ["font-weight: bold; background-color:#1a3a1a; color:white"] * len(r)
                         if r[group_col] == "TOTAL" else [""] * len(r), axis=1)
                  .background_gradient(subset=["Eff %"],
                                       cmap=EFFICIENCY_CMAP, vmin=0, vmax=100)
                  .background_gradient(subset=["Turn-up %"],
                                       cmap=TURNUP_CMAP, vmin=0, vmax=100),
                width='stretch', hide_index=True, height=400
            )
        with col2:
            if chart and not s.empty:
                fig = px.bar(s, x=group_col, y="Turnup_%",
                             text_auto=".2f", color="Turnup_%",
                             color_continuous_scale="RdYlGn",
                             range_color=(0, 100), height=400)
                fig.update_layout(xaxis_title="",
                                  yaxis_title="Turn-up %",
                                  coloraxis_showscale=False)
                apply_theme(fig, height=400)
                st.plotly_chart(fig, width='stretch', key=f"bar_{key}")
        return s

    def show_sdo_name_table(df_paid, df_bill, key="sdo"):
        st.subheader("🏢 SDO Name wise")
        if df_paid.empty and df_bill.empty:
            st.info("Koi data nahi.")
            return

        if SDO_NAME_COL not in df_paid.columns and SDO_NAME_COL not in df_bill.columns:
            st.info("SDO_NAME nahi mila.")
            return

        if not df_bill.empty and CA_COL in df_bill.columns:
            agg_dict = {
                "CA_Total":    (CA_COL, "sum"),
                "Bills_Total": (CA_COL, "size"),
            }
            if ARREAR_COL in df_bill.columns:
                agg_dict["Arrear_Total"] = (ARREAR_COL, "sum")
            else:
                agg_dict["Arrear_Total"] = (CA_COL, lambda x: 0)
            ca = (df_bill.groupby([SDO_NAME_COL, SDO_CODE_COL], as_index=False, observed=True)
                         .agg(**agg_dict))
        else:
            ca = pd.DataFrame({SDO_NAME_COL: [], SDO_CODE_COL: [],
                               "CA_Total": [], "Bills_Total": [], "Arrear_Total": []})

        if not df_paid.empty:
            paid = (df_paid.groupby([SDO_NAME_COL, SDO_CODE_COL], as_index=False, observed=True)
                           .agg(Total_Paid=(AMOUNT_COL, "sum"),
                                Txns=(AMOUNT_COL, "size")))
        else:
            paid = pd.DataFrame({SDO_NAME_COL: [], SDO_CODE_COL: [],
                                 "Total_Paid": [], "Txns": []})

        s = pd.merge(ca, paid, on=[SDO_NAME_COL, SDO_CODE_COL],
                     how="outer").fillna(0)
        s["CA (Cr.)"]     = (s["CA_Total"] / CRORE).round(2)
        s["Paid (Cr.)"]   = (s["Total_Paid"] / CRORE).round(2)
        s["Total Outstanding (Cr.)"] = (s["Arrear_Total"] / CRORE).round(2)
        s["Paid Count"]   = s["Txns"].astype(int)
        s["Bills"]        = s["Bills_Total"].astype(int)
        s["Eff %"]        = (s["Total_Paid"] /
                             s["CA_Total"].replace(0, pd.NA) * 100).round(2).fillna(0)
        s["Turn-up %"]    = (s["Txns"] /
                             s["Bills_Total"].replace(0, pd.NA) * 100).round(2).fillna(0)

        s = s.sort_values("Paid (Cr.)", ascending=False)
        display = s[[SDO_NAME_COL, SDO_CODE_COL, "CA (Cr.)", "Paid (Cr.)",
                     "Total Outstanding (Cr.)", "Paid Count", "Bills", "Eff %", "Turn-up %"]].copy()
        display.columns = ["SDO Name", "SDO Code", "CA (Cr.)", "Paid (Cr.)",
                           "Total Outstanding (Cr.)", "Paid Count", "Bills", "Eff %", "Turn-up %"]

        tot_ca    = round(s["CA (Cr.)"].sum(), 2)
        tot_paid  = round(s["Paid (Cr.)"].sum(), 2)
        tot_arr   = round(s["Total Outstanding (Cr.)"].sum(), 2)
        tot_txns  = int(s["Paid Count"].sum())
        tot_bills = int(s["Bills"].sum())
        tot_eff   = round((tot_paid / tot_ca * 100) if tot_ca else 0, 2)
        tot_turn  = round((tot_txns / tot_bills * 100) if tot_bills else 0, 2)

        totals = pd.DataFrame([{
            "SDO Name":     "TOTAL",
            "SDO Code":     "",
            "CA (Cr.)":     tot_ca,
            "Paid (Cr.)":   tot_paid,
            "Total Outstanding (Cr.)": tot_arr,
            "Paid Count":   tot_txns,
            "Bills":        tot_bills,
            "Eff %":        tot_eff,
            "Turn-up %":    tot_turn,
        }])
        display_full = pd.concat([display, totals], ignore_index=True)
        # 🛠️ ARROW FIX
        display_full = _arrow_safe(display_full)

        col1, col2 = st.columns([1.7, 1])
        with col1:
            st.dataframe(
                display_full.style
                  .format({
                      "CA (Cr.)":     "₹ {:,.2f}",
                      "Paid (Cr.)":   "₹ {:,.2f}",
                      "Total Outstanding (Cr.)": "₹ {:,.2f}",
                      "Paid Count":   "{:,}",
                      "Bills":        "{:,}",
                      "Eff %":        "{:.2f}%",
                      "Turn-up %":    "{:.2f}%",
                  })
                  .apply(lambda r: ["font-weight: bold; background-color:#1a3a1a; color:white"] * len(r)
                         if r["SDO Name"] == "TOTAL" else [""] * len(r), axis=1)
                  .background_gradient(subset=["Eff %"],
                                       cmap=EFFICIENCY_CMAP, vmin=0, vmax=100)
                  .background_gradient(subset=["Turn-up %"],
                                       cmap=TURNUP_CMAP, vmin=0, vmax=100),
                width='stretch', hide_index=True, height=400
            )
        with col2:
            if not s.empty:
                fig = px.bar(s, x=SDO_NAME_COL, y="Turn-up %",
                             text_auto=".2f", color="Turn-up %",
                             color_continuous_scale="RdYlGn",
                             range_color=(0, 100), height=400)
                fig.update_layout(xaxis_title="",
                                  yaxis_title="Turn-up %",
                                  coloraxis_showscale=False)
                apply_theme(fig, height=400)
                st.plotly_chart(fig, width='stretch', key=f"bar_{key}")
        return s

    def show_efficiency_table(df_paid, df_bill, group_col, title, key,
                              order=None):
        st.subheader(title)
        if df_paid.empty and df_bill.empty:
            st.info("Koi data nahi.")
            return

        s = make_summary(df_paid, df_bill, group_col, order=order)

        display = s[[group_col, "CA_Cr", "Total_Cr", "Arrear_Cr", "Efficiency_%",
                     "Txns", "Bills_Total", "Turnup_%"]].rename(columns={
            "CA_Cr":         "CA (Cr.)",
            "Total_Cr":      "Paid (Cr.)",
            "Arrear_Cr":     "Total Outstanding (Cr.)",
            "Efficiency_%":  "Eff %",
            "Txns":          "Paid",
            "Bills_Total":   "Bills",
            "Turnup_%":      "Turn-up %",
        })

        tot_ca    = round(s["CA_Cr"].sum(), 2)
        tot_paid  = round(s["Total_Cr"].sum(), 2)
        tot_arr   = round(s["Arrear_Cr"].sum(), 2)
        tot_eff   = round((tot_paid / tot_ca * 100) if tot_ca else 0, 2)
        tot_txns  = int(s["Txns"].sum())
        tot_bills = int(s["Bills_Total"].sum())
        tot_turn  = round((tot_txns / tot_bills * 100) if tot_bills else 0, 2)

        totals = pd.DataFrame([{
            group_col:     "TOTAL",
            "CA (Cr.)":    tot_ca,
            "Paid (Cr.)":  tot_paid,
            "Total Outstanding (Cr.)": tot_arr,
            "Eff %":       tot_eff,
            "Paid":        tot_txns,
            "Bills":       tot_bills,
            "Turn-up %":   tot_turn,
        }])
        display_full = pd.concat([display, totals], ignore_index=True)
        # 🛠️ ARROW FIX
        display_full = _arrow_safe(display_full)

        col1, col2 = st.columns([1.7, 1])
        with col1:
            st.dataframe(
                display_full.style
                  .format({
                      "CA (Cr.)":     "₹ {:,.2f}",
                      "Paid (Cr.)":   "₹ {:,.2f}",
                      "Total Outstanding (Cr.)": "₹ {:,.2f}",
                      "Eff %":        "{:.2f}%",
                      "Paid":         "{:,}",
                      "Bills":        "{:,}",
                      "Turn-up %":    "{:.2f}%",
                  })
                  .apply(lambda r: ["font-weight: bold; background-color:#1a3a1a; color:white"] * len(r)
                         if r[group_col] == "TOTAL" else [""] * len(r), axis=1)
                  .background_gradient(subset=["Eff %"],
                                       cmap=EFFICIENCY_CMAP, vmin=0, vmax=100)
                  .background_gradient(subset=["Turn-up %"],
                                       cmap=TURNUP_CMAP, vmin=0, vmax=100),
                width='stretch', hide_index=True, height=400
            )
        with col2:
            if not s.empty:
                fig = px.bar(s, x=group_col, y="Efficiency_%",
                             text_auto=".2f", color="Efficiency_%",
                             color_continuous_scale="RdYlGn",
                             range_color=(0, 100), height=400)
                fig.update_layout(xaxis_title="",
                                  yaxis_title="Efficiency %",
                                  coloraxis_showscale=False)
                apply_theme(fig, height=400)
                st.plotly_chart(fig, width='stretch', key=f"eff_{key}")
        return s

    def show_sdo_efficiency_table(df_paid, df_bill, key="eff_sdo"):
        st.subheader("SDO Name-wise")
        if df_paid.empty and df_bill.empty:
            st.info("Koi data nahi.")
            return

        if not df_bill.empty and CA_COL in df_bill.columns:
            agg_dict = {
                "CA_Total":    (CA_COL, "sum"),
                "Bills_Total": (CA_COL, "size"),
            }
            if ARREAR_COL in df_bill.columns:
                agg_dict["Arrear_Total"] = (ARREAR_COL, "sum")
            else:
                agg_dict["Arrear_Total"] = (CA_COL, lambda x: 0)
            ca = (df_bill.groupby(SDO_NAME_COL, as_index=False, observed=True)
                         .agg(**agg_dict))
        else:
            ca = pd.DataFrame({SDO_NAME_COL: [], "CA_Total": [], "Bills_Total": [],
                               "Arrear_Total": []})

        if not df_paid.empty:
            paid = (df_paid.groupby(SDO_NAME_COL, as_index=False, observed=True)
                           .agg(Total_Paid=(AMOUNT_COL, "sum"),
                                Txns=(AMOUNT_COL, "size")))
        else:
            paid = pd.DataFrame({SDO_NAME_COL: [], "Total_Paid": [], "Txns": []})

        s = pd.merge(ca, paid, on=SDO_NAME_COL, how="outer").fillna(0)
        s["CA (Cr.)"]     = (s["CA_Total"] / CRORE).round(2)
        s["Paid (Cr.)"]   = (s["Total_Paid"] / CRORE).round(2)
        s["Total Outstanding (Cr.)"] = (s["Arrear_Total"] / CRORE).round(2)
        s["Paid"]         = s["Txns"].astype(int)
        s["Bills"]        = s["Bills_Total"].astype(int)
        s["Eff %"]        = (s["Total_Paid"] /
                             s["CA_Total"].replace(0, pd.NA) * 100).round(2).fillna(0)
        s["Turn-up %"]    = (s["Txns"] /
                             s["Bills_Total"].replace(0, pd.NA) * 100).round(2).fillna(0)
        s = s.sort_values("Paid (Cr.)", ascending=False)

        display = s[[SDO_NAME_COL, "CA (Cr.)", "Paid (Cr.)", "Total Outstanding (Cr.)",
                     "Eff %", "Paid", "Bills", "Turn-up %"]].copy()
        display.columns = ["SDO Name", "CA (Cr.)", "Paid (Cr.)", "Total Outstanding (Cr.)",
                           "Eff %", "Paid", "Bills", "Turn-up %"]

        tot_ca    = round(s["CA (Cr.)"].sum(), 2)
        tot_paid  = round(s["Paid (Cr.)"].sum(), 2)
        tot_arr   = round(s["Total Outstanding (Cr.)"].sum(), 2)
        tot_eff   = round((tot_paid / tot_ca * 100) if tot_ca else 0, 2)
        tot_txns  = int(s["Paid"].sum())
        tot_bills = int(s["Bills"].sum())
        tot_turn  = round((tot_txns / tot_bills * 100) if tot_bills else 0, 2)

        totals = pd.DataFrame([{
            "SDO Name":     "TOTAL",
            "CA (Cr.)":     tot_ca,
            "Paid (Cr.)":   tot_paid,
            "Total Outstanding (Cr.)": tot_arr,
            "Eff %":        tot_eff,
            "Paid":         tot_txns,
            "Bills":        tot_bills,
            "Turn-up %":    tot_turn,
        }])
        display_full = pd.concat([display, totals], ignore_index=True)
        # 🛠️ ARROW FIX
        display_full = _arrow_safe(display_full)

        col1, col2 = st.columns([1.7, 1])
        with col1:
            st.dataframe(
                display_full.style
                  .format({
                      "CA (Cr.)":     "₹ {:,.2f}",
                      "Paid (Cr.)":   "₹ {:,.2f}",
                      "Total Outstanding (Cr.)": "₹ {:,.2f}",
                      "Eff %":        "{:.2f}%",
                      "Paid":         "{:,}",
                      "Bills":        "{:,}",
                      "Turn-up %":    "{:.2f}%",
                  })
                  .apply(lambda r: ["font-weight: bold; background-color:#1a3a1a; color:white"] * len(r)
                         if r["SDO Name"] == "TOTAL" else [""] * len(r), axis=1)
                  .background_gradient(subset=["Eff %"],
                                       cmap=EFFICIENCY_CMAP, vmin=0, vmax=100)
                  .background_gradient(subset=["Turn-up %"],
                                       cmap=TURNUP_CMAP, vmin=0, vmax=100),
                width='stretch', hide_index=True, height=400
            )
        with col2:
            if not s.empty:
                fig = px.bar(s, x=SDO_NAME_COL, y="Eff %",
                             text_auto=".2f", color="Eff %",
                             color_continuous_scale="RdYlGn",
                             range_color=(0, 100), height=400)
                fig.update_layout(xaxis_title="",
                                  yaxis_title="Efficiency %",
                                  coloraxis_showscale=False)
                apply_theme(fig, height=400)
                st.plotly_chart(fig, width='stretch', key=f"eff_{key}")
        return s

    # ---------- EXCEL ----------
    def build_excel(df_paid, df_bill, year, month):
        output = io.BytesIO()
        wb = xlsxwriter.Workbook(output, {"in_memory": True,
                                          "nan_inf_to_errors": True})

        fmt_header = wb.add_format({"bold": True, "bg_color": "#1F4E78",
                                    "font_color": "white", "border": 1,
                                    "align": "center"})
        fmt_cr  = wb.add_format({"num_format": "₹ #,##0.00", "border": 1})
        fmt_cr_total = wb.add_format({"num_format": "₹ #,##0.00", "border": 1,
                                      "bold": True, "bg_color": "#E8F5E9"})
        fmt_int = wb.add_format({"num_format": "#,##0", "border": 1})
        fmt_int_total = wb.add_format({"num_format": "#,##0", "border": 1,
                                       "bold": True, "bg_color": "#E8F5E9"})
        fmt_pct = wb.add_format({"num_format": "0.00\"%\"", "border": 1})
        fmt_pct_total = wb.add_format({"num_format": "0.00\"%\"", "border": 1,
                                       "bold": True, "bg_color": "#E8F5E9"})
        fmt_txt = wb.add_format({"border": 1})
        fmt_txt_total = wb.add_format({"border": 1, "bold": True,
                                       "bg_color": "#E8F5E9"})

        def add_sheet(name, df, cols):
            ws = wb.add_worksheet(name[:31])
            for i, c in enumerate(cols):
                ws.write(0, i, str(c), fmt_header)
            for r, (_, row) in enumerate(df.iterrows(), start=1):
                is_total = str(row.get(cols[0], "")).strip().upper() == "TOTAL"
                for c, col in enumerate(cols):
                    val = row.get(col, "")
                    if isinstance(val, float) and (pd.isna(val) or
                                                   val in (float("inf"), float("-inf"))):
                        val = ""
                    if isinstance(val, (int, float)) and val != "" and not pd.isna(val):
                        colname = str(col)
                        if "%" in colname or "Eff" in colname or "Turn-up" in colname or "Ach" in colname:
                            ws.write_number(r, c, round(float(val), 2),
                                            fmt_pct_total if is_total else fmt_pct)
                        elif "Count" in colname or "Txns" in colname or "Bills" in colname:
                            ws.write_number(r, c, int(val),
                                            fmt_int_total if is_total else fmt_int)
                        else:
                            ws.write_number(r, c, round(float(val), 2),
                                            fmt_cr_total if is_total else fmt_cr)
                    else:
                        ws.write(r, c, str(val),
                                 fmt_txt_total if is_total else fmt_txt)
            for i, c in enumerate(cols):
                ws.set_column(i, i, max(14, len(str(c)) + 4))

        ws = wb.add_worksheet("Summary")
        ws.write(0, 0, f"Payment Dashboard — {year}-{month:02d}", fmt_header)
        total_paid = round(float(df_paid[AMOUNT_COL].sum()), 2) if not df_paid.empty else 0
        total_ca   = round(float(df_bill[CA_COL].sum()), 2) if not df_bill.empty else 0
        total_outstanding = (round(float(df_bill[ARREAR_COL].sum()), 2)
                             if (not df_bill.empty and ARREAR_COL in df_bill.columns) else 0)
        eff  = round((total_paid / total_ca * 100) if total_ca else 0, 2)
        turn = round((len(df_paid) / len(df_bill) * 100) if len(df_bill) else 0, 2)
        ws.write(1, 0, "Current Assessment (Cr.)", fmt_header); ws.write(1, 1, round(total_ca / CRORE, 2), fmt_cr)
        ws.write(2, 0, "Total Paid (Cr.)", fmt_header);          ws.write(2, 1, round(total_paid / CRORE, 2), fmt_cr)
        ws.write(3, 0, "Total Outstanding (Cr.)", fmt_header);   ws.write(3, 1, round(total_outstanding / CRORE, 2), fmt_cr)
        ws.write(4, 0, "Transactions", fmt_header);              ws.write(4, 1, len(df_paid), fmt_int)
        ws.write(5, 0, "Total Bills", fmt_header);               ws.write(5, 1, len(df_bill), fmt_int)
        ws.write(6, 0, "Efficiency %", fmt_header);              ws.write(6, 1, eff, fmt_pct)
        ws.write(7, 0, "Turn-up %", fmt_header);                 ws.write(7, 1, turn, fmt_pct)
        ws.set_column(0, 0, 30); ws.set_column(1, 1, 18)

        s = make_summary(df_paid, df_bill, CATEGORY_COL)
        s2 = s[[CATEGORY_COL, "CA_Cr", "Total_Cr", "Arrear_Cr", "Txns",
                "Bills_Total", "Efficiency_%", "Turnup_%"]].rename(columns={
            "CA_Cr": "CA (Cr.)",
            "Total_Cr": "Total Paid (Cr.)",
            "Arrear_Cr": "Total Outstanding (Cr.)",
            "Txns": "Paid Count", "Bills_Total": "Total Bills",
            "Efficiency_%": "Eff %", "Turnup_%": "Turn-up %"})
        tot = {
            CATEGORY_COL:   "TOTAL",
            "CA (Cr.)":     round(s["CA_Cr"].sum(), 2),
            "Total Paid (Cr.)": round(s["Total_Cr"].sum(), 2),
            "Total Outstanding (Cr.)": round(s["Arrear_Cr"].sum(), 2),
            "Paid Count":   int(s["Txns"].sum()),
            "Total Bills":  int(s["Bills_Total"].sum()),
            "Eff %":        round((s["Total_Cr"].sum() / s["CA_Cr"].sum() * 100) if s["CA_Cr"].sum() else 0, 2),
            "Turn-up %":    round((s["Txns"].sum() / s["Bills_Total"].sum() * 100) if s["Bills_Total"].sum() else 0, 2),
        }
        s2 = pd.concat([s2, pd.DataFrame([tot])], ignore_index=True)
        add_sheet("Tariff wise", s2, list(s2.columns))

        if not df_paid.empty and "LOAD_CATEGORY" in df_paid.columns:
            s = make_summary(df_paid.dropna(subset=["LOAD_CATEGORY"]),
                             df_bill.dropna(subset=["LOAD_CATEGORY"]) if not df_bill.empty else df_bill,
                             "LOAD_CATEGORY", order=LOAD_ORDER)
            s3 = s[["LOAD_CATEGORY", "CA_Cr", "Total_Cr", "Arrear_Cr", "Txns",
                    "Bills_Total", "Efficiency_%", "Turnup_%"]].rename(columns={
                "CA_Cr": "CA (Cr.)",
                "Total_Cr": "Total Paid (Cr.)",
                "Arrear_Cr": "Total Outstanding (Cr.)",
                "Txns": "Paid Count", "Bills_Total": "Total Bills",
                "Efficiency_%": "Eff %", "Turnup_%": "Turn-up %"})
            tot = {
                "LOAD_CATEGORY": "TOTAL",
                "CA (Cr.)":     round(s["CA_Cr"].sum(), 2),
                "Total Paid (Cr.)": round(s["Total_Cr"].sum(), 2),
                "Total Outstanding (Cr.)": round(s["Arrear_Cr"].sum(), 2),
                "Paid Count":   int(s["Txns"].sum()),
                "Total Bills":  int(s["Bills_Total"].sum()),
                "Eff %":        round((s["Total_Cr"].sum() / s["CA_Cr"].sum() * 100) if s["CA_Cr"].sum() else 0, 2),
                "Turn-up %":    round((s["Txns"].sum() / s["Bills_Total"].sum() * 100) if s["Bills_Total"].sum() else 0, 2),
            }
            s3 = pd.concat([s3, pd.DataFrame([tot])], ignore_index=True)
            add_sheet("Load wise", s3, list(s3.columns))

        if SDO_CODE_COL in df_paid.columns or SDO_CODE_COL in df_bill.columns:
            s = make_summary(df_paid, df_bill, SDO_CODE_COL)
            s4 = s[[SDO_CODE_COL, "CA_Cr", "Total_Cr", "Arrear_Cr", "Txns",
                    "Bills_Total", "Efficiency_%", "Turnup_%"]].rename(columns={
                "CA_Cr": "CA (Cr.)",
                "Total_Cr": "Total Paid (Cr.)",
                "Arrear_Cr": "Total Outstanding (Cr.)",
                "Txns": "Paid Count", "Bills_Total": "Total Bills",
                "Efficiency_%": "Eff %", "Turnup_%": "Turn-up %"})
            s4.insert(0, "SDO Name", s4[SDO_CODE_COL].apply(sdo_short))
            tot = {
                "SDO Name":     "TOTAL",
                SDO_CODE_COL:   "",
                "CA (Cr.)":     round(s["CA_Cr"].sum(), 2),
                "Total Paid (Cr.)": round(s["Total_Cr"].sum(), 2),
                "Total Outstanding (Cr.)": round(s["Arrear_Cr"].sum(), 2),
                "Paid Count":   int(s["Txns"].sum()),
                "Total Bills":  int(s["Bills_Total"].sum()),
                "Eff %":        round((s["Total_Cr"].sum() / s["CA_Cr"].sum() * 100) if s["CA_Cr"].sum() else 0, 2),
                "Turn-up %":    round((s["Txns"].sum() / s["Bills_Total"].sum() * 100) if s["Bills_Total"].sum() else 0, 2),
            }
            tot_df = pd.DataFrame([tot])[list(s4.columns)]
            s4 = pd.concat([s4, tot_df], ignore_index=True)
            add_sheet("SDO wise", s4, list(s4.columns))

        rows = []
        for grp in TARGET_GROUPS:
            sub_p = df_paid[df_paid[SDO_CODE_COL].isin(grp["codes"])] if not df_paid.empty else pd.DataFrame()
            sub_b = df_bill[df_bill[SDO_CODE_COL].isin(grp["codes"])] if not df_bill.empty else pd.DataFrame()
            actual = round(float(sub_p[AMOUNT_COL].sum()) / CRORE, 2) if not sub_p.empty else 0.0
            txns   = len(sub_p)
            bills  = len(sub_b)
            target = grp["target_cr"]
            ach    = round((actual / target * 100) if target else 0, 2)
            turn   = round((txns / bills * 100) if bills else 0, 2)
            rows.append({
                "SDO":           grp["label"],
                "Actual (Cr.)":  round(actual, 2),
                "Target (Cr.)":  target,
                "Ach %":         round(ach, 2),
                "Txns":          txns,
                "Bills":         bills,
                "Turn-up %":     round(turn, 2),
            })
        tgt_df = pd.DataFrame(rows)
        total_row = {
            "SDO": "TOTAL",
            "Actual (Cr.)": round(tgt_df["Actual (Cr.)"].sum(), 2),
            "Target (Cr.)": TOTAL_TARGET_CR,
            "Ach %": round(tgt_df["Actual (Cr.)"].sum() / TOTAL_TARGET_CR * 100, 2),
            "Txns": int(tgt_df["Txns"].sum()),
            "Bills": int(tgt_df["Bills"].sum()),
            "Turn-up %": round(tgt_df["Txns"].sum() /
                               max(tgt_df["Bills"].sum(), 1) * 100, 2),
        }
        tgt_df = pd.concat([tgt_df, pd.DataFrame([total_row])], ignore_index=True)
        add_sheet("Target vs Achievement", tgt_df, list(tgt_df.columns))

        if not df_paid.empty:
            d = df_paid.copy()
            d["Day"] = d[DATE_COL].dt.day
            daily = (d.groupby("Day", as_index=False)
                       .agg(Amount=(AMOUNT_COL, "sum"), Txns=(AMOUNT_COL, "size")))
            daily["Amount (Cr.)"] = (daily["Amount"] / CRORE).round(2)
            daily = daily[["Day", "Amount (Cr.)", "Txns"]]
            tot = {
                "Day": "TOTAL",
                "Amount (Cr.)": round(daily["Amount (Cr.)"].sum(), 2),
                "Txns": int(daily["Txns"].sum()),
            }
            daily = pd.concat([daily, pd.DataFrame([tot])], ignore_index=True)
            add_sheet("Daily Collection", daily, list(daily.columns))

        if not df_paid.empty and ACCT_COL in df_paid.columns:
            grp = [ACCT_COL, NAME_COL] if NAME_COL in df_paid.columns else [ACCT_COL]
            cons = (df_paid.groupby(grp, as_index=False)
                            .agg(Total_Paid=(AMOUNT_COL, "sum"),
                                 Txns=(AMOUNT_COL, "size"))
                            .sort_values("Total_Paid", ascending=False).head(10))
            cons["Total_Cr"] = (cons["Total_Paid"] / CRORE).round(2)
            cons = cons[grp + ["Total_Cr", "Txns"]].rename(columns={
                "Total_Cr": "Total Paid (Cr.)", "Txns": "Paid Count"})
            tot_dict = {grp[0]: "TOTAL"}
            if len(grp) > 1:
                tot_dict[grp[1]] = ""
            tot_dict["Total Paid (Cr.)"] = round(cons["Total Paid (Cr.)"].sum(), 2)
            tot_dict["Paid Count"] = int(cons["Paid Count"].sum())
            cons = pd.concat([cons, pd.DataFrame([tot_dict])[list(cons.columns)]],
                             ignore_index=True)
            add_sheet("Top 10 Consumers", cons, list(cons.columns))

        wb.close()
        output.seek(0)
        return output.read()

    # ---------- PDF ----------
    def build_pdf(df_paid, df_bill, year, month):
        buffer = io.BytesIO()
        page_size = A4
        doc = SimpleDocTemplate(
            buffer,
            pagesize=page_size,
            leftMargin=28, rightMargin=28,
            topMargin=26, bottomMargin=38,
            title=f"UPPCL Payment Dashboard — {year}-{month:02d}",
            author="U.P. Power Corporation Limited",
        )
        styles = getSampleStyleSheet()

        title_style = ParagraphStyle(
            "T", parent=styles["Title"],
            fontSize=18,
            textColor=colors.HexColor("#1F4E78"),
            spaceAfter=3,
            alignment=1,
            leading=22,
        )
        sub_style = ParagraphStyle(
            "S", parent=styles["Normal"],
            fontSize=9,
            textColor=colors.HexColor("#555555"),
            alignment=1,
            spaceAfter=8,
        )
        h3 = ParagraphStyle(
            "H3", parent=styles["Heading3"],
            fontSize=11.5,
            textColor=colors.HexColor("#1F4E78"),
            spaceBefore=8,
            spaceAfter=4,
            alignment=0,
        )
        footer_style = ParagraphStyle(
            "F", parent=styles["Normal"],
            fontSize=8,
            textColor=colors.HexColor("#666666"),
            alignment=1,
            spaceBefore=8,
        )

        story = []

        story.append(Paragraph(
            "<b>⚡ U.P. POWER CORPORATION LIMITED</b>",
            title_style
        ))
        story.append(Paragraph(
            f"<b>Payment Dashboard</b> &nbsp;|&nbsp; Month: <b>{year}-{month:02d}</b> "
            f"&nbsp;|&nbsp; Generated: {datetime.now().strftime('%d-%b-%Y %H:%M')}",
            sub_style
        ))

        total_paid = round(float(df_paid[AMOUNT_COL].sum()) / CRORE, 2) if not df_paid.empty else 0
        total_ca   = round(float(df_bill[CA_COL].sum()) / CRORE, 2) if not df_bill.empty else 0
        total_outstanding = (round(float(df_bill[ARREAR_COL].sum()) / CRORE, 2)
                             if (not df_bill.empty and ARREAR_COL in df_bill.columns) else 0)
        txns       = len(df_paid)
        bills      = len(df_bill)
        eff        = round((total_paid / total_ca * 100) if total_ca else 0, 2)
        turn       = round((txns / bills * 100) if bills else 0, 2)

        kpi_data = [
            ["Current Assessment", "Total Paid", "Total Outstanding",
             "Efficiency", "Turn-up %", "Txns", "Bills"],
            [f"₹ {total_ca:,.2f} Cr.",
             f"₹ {total_paid:,.2f} Cr.",
             f"₹ {total_outstanding:,.2f} Cr.",
             f"{eff:.2f}%",
             f"{turn:.2f}%",
             f"{txns:,}",
             f"{bills:,}"],
        ]

        kpi_table = Table(kpi_data, colWidths=[78, 78, 88, 65, 65, 55, 55])
        kpi_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("BACKGROUND", (0, 1), (-1, 1), colors.HexColor("#E8F0F8")),
            ("TEXTCOLOR", (0, 1), (-1, 1), colors.HexColor("#1F4E78")),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTNAME", (0, 1), (-1, 1), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#1F4E78")),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]))
        story.append(kpi_table)
        story.append(Spacer(1, 10))

        def make_table(title, df, cols, col_widths=None):
            story.append(Paragraph(f"<b>{title}</b>", h3))
            data = [list(cols)] + df[cols].astype(str).values.tolist()
            if col_widths is None:
                n = len(cols)
                col_widths = [538 / n] * n
            t = Table(data, colWidths=col_widths, repeatRows=1)
            style_list = [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#00529B")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#999999")),
                ("ALIGN", (0, 0), (0, -1), "LEFT"),
                ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
            ]
            last = len(data) - 1
            if last > 0 and str(data[last][0]).strip().upper().startswith("TOTAL"):
                style_list.append(("BACKGROUND", (0, last), (-1, last),
                                   colors.HexColor("#E8F5E9")))
                style_list.append(("FONTNAME", (0, last), (-1, last),
                                   "Helvetica-Bold"))
            for i in range(1, len(data) - 1):
                if i % 2 == 0:
                    style_list.append(("BACKGROUND", (0, i), (-1, i),
                                       colors.HexColor("#F5F9FF")))
            t.setStyle(TableStyle(style_list))
            story.append(t)
            story.append(Spacer(1, 9))

        rows = []
        for grp in TARGET_GROUPS:
            sub_p = df_paid[df_paid[SDO_CODE_COL].isin(grp["codes"])] if not df_paid.empty else pd.DataFrame()
            sub_b = df_bill[df_bill[SDO_CODE_COL].isin(grp["codes"])] if not df_bill.empty else pd.DataFrame()
            actual = round(float(sub_p[AMOUNT_COL].sum()) / CRORE, 2) if not sub_p.empty else 0.0
            txns_g = len(sub_p)
            bills_g = len(sub_b)
            target = grp["target_cr"]
            ach    = round((actual / target * 100) if target else 0, 2)
            turn_g = round((txns_g / bills_g * 100) if bills_g else 0, 2)
            rows.append({
                "SDO":           grp["label"],
                "Actual (Cr.)":  f"{actual:.2f}",
                "Target (Cr.)":  f"{target:.2f}",
                "Ach %":         f"{ach:.2f}%",
                "Turn-up %":     f"{turn_g:.2f}%",
            })
        tgt_df = pd.DataFrame(rows)
        ta = round(sum(float(r["Actual (Cr.)"]) for r in rows), 2)
        tgt_df = pd.concat([tgt_df, pd.DataFrame([{
            "SDO": "TOTAL (incl. Independent Feeders)",
            "Actual (Cr.)": f"{ta:.2f}",
            "Target (Cr.)": f"{TOTAL_TARGET_CR:.2f}",
            "Ach %": f"{(ta/TOTAL_TARGET_CR*100):.2f}%",
            "Turn-up %": "—",
        }])], ignore_index=True)

        make_table(
            "🎯 Target vs Achievement",
            tgt_df,
            list(tgt_df.columns),
            col_widths=[200, 85, 85, 84, 84]
        )

        s_tar = make_summary(df_paid, df_bill, CATEGORY_COL)
        s_tar["CA (Cr.)"]                = (s_tar["CA_Total"] / CRORE).round(2)
        s_tar["Paid (Cr.)"]              = (s_tar["Total_Paid"] / CRORE).round(2)
        s_tar["Total Outstanding (Cr.)"] = (s_tar["Arrear_Total"] / CRORE).round(2)
        s_tar["Eff %"]                   = s_tar["Efficiency_%"].round(2)
        s_tar["Turn-up %"]               = s_tar["Turnup_%"].round(2)
        tdf = s_tar[[CATEGORY_COL, "CA (Cr.)", "Paid (Cr.)",
                     "Total Outstanding (Cr.)", "Eff %", "Turn-up %"]].copy()
        tdf = pd.concat([tdf, pd.DataFrame([{
            CATEGORY_COL: "TOTAL",
            "CA (Cr.)":                round(tdf["CA (Cr.)"].sum(), 2),
            "Paid (Cr.)":              round(tdf["Paid (Cr.)"].sum(), 2),
            "Total Outstanding (Cr.)": round(tdf["Total Outstanding (Cr.)"].sum(), 2),
            "Eff %": round((tdf["Paid (Cr.)"].sum() / tdf["CA (Cr.)"].sum() * 100)
                           if tdf["CA (Cr.)"].sum() else 0, 2),
            "Turn-up %": round((s_tar["Txns"].sum() / s_tar["Bills_Total"].sum() * 100)
                               if s_tar["Bills_Total"].sum() else 0, 2),
        }])], ignore_index=True)

        make_table(
            "📊 Tariff Type wise",
            tdf,
            list(tdf.columns),
            col_widths=[105, 78, 78, 118, 78, 81]
        )

        if not df_paid.empty and "LOAD_CATEGORY" in df_paid.columns:
            s_load = make_summary(
                df_paid.dropna(subset=["LOAD_CATEGORY"]),
                df_bill.dropna(subset=["LOAD_CATEGORY"]) if not df_bill.empty else df_bill,
                "LOAD_CATEGORY", order=LOAD_ORDER
            )
            s_load["CA (Cr.)"]                = (s_load["CA_Total"] / CRORE).round(2)
            s_load["Paid (Cr.)"]              = (s_load["Total_Paid"] / CRORE).round(2)
            s_load["Total Outstanding (Cr.)"] = (s_load["Arrear_Total"] / CRORE).round(2)
            s_load["Eff %"]                   = s_load["Efficiency_%"].round(2)
            s_load["Turn-up %"]               = s_load["Turnup_%"].round(2)
            ldf = s_load[["LOAD_CATEGORY", "CA (Cr.)", "Paid (Cr.)",
                          "Total Outstanding (Cr.)", "Eff %", "Turn-up %"]].copy()
            ldf = pd.concat([ldf, pd.DataFrame([{
                "LOAD_CATEGORY": "TOTAL",
                "CA (Cr.)":                round(ldf["CA (Cr.)"].sum(), 2),
                "Paid (Cr.)":              round(ldf["Paid (Cr.)"].sum(), 2),
                "Total Outstanding (Cr.)": round(ldf["Total Outstanding (Cr.)"].sum(), 2),
                "Eff %": round((ldf["Paid (Cr.)"].sum() / ldf["CA (Cr.)"].sum() * 100)
                               if ldf["CA (Cr.)"].sum() else 0, 2),
                "Turn-up %": round((s_load["Txns"].sum() / s_load["Bills_Total"].sum() * 100)
                                   if s_load["Bills_Total"].sum() else 0, 2),
            }])], ignore_index=True)

            make_table(
                "⚡ Load Category wise",
                ldf,
                list(ldf.columns),
                col_widths=[105, 78, 78, 118, 78, 81]
            )

        story.append(PageBreak())

        story.append(Paragraph(
            f"<b>⚡ UPPCL Payment Dashboard — {year}-{month:02d}</b>",
            ParagraphStyle("sub", parent=styles["Normal"], fontSize=10,
                           textColor=colors.HexColor("#1F4E78"),
                           alignment=1, spaceAfter=8)
        ))

        if SDO_CODE_COL in df_paid.columns or SDO_CODE_COL in df_bill.columns:
            s_sdo = make_summary(df_paid, df_bill, SDO_CODE_COL)
            s_sdo["SDO Name"] = s_sdo[SDO_CODE_COL].apply(sdo_short)
            s_sdo["CA (Cr.)"]                = (s_sdo["CA_Total"] / CRORE).round(2)
            s_sdo["Paid (Cr.)"]              = (s_sdo["Total_Paid"] / CRORE).round(2)
            s_sdo["Total Outstanding (Cr.)"] = (s_sdo["Arrear_Total"] / CRORE).round(2)
            s_sdo["Eff %"]                   = s_sdo["Efficiency_%"].round(2)
            s_sdo["Turn-up %"]               = s_sdo["Turnup_%"].round(2)

            sdf = s_sdo[["SDO Name", SDO_CODE_COL, "CA (Cr.)", "Paid (Cr.)",
                         "Total Outstanding (Cr.)", "Eff %", "Turn-up %"]].copy()
            sdf.columns = ["SDO Name", "Code", "CA (Cr.)", "Paid (Cr.)",
                           "Total Outstanding (Cr.)", "Eff %", "Turn-up %"]
            sdf = pd.concat([sdf, pd.DataFrame([{
                "SDO Name": "TOTAL",
                "Code": "",
                "CA (Cr.)":                round(sdf["CA (Cr.)"].sum(), 2),
                "Paid (Cr.)":              round(sdf["Paid (Cr.)"].sum(), 2),
                "Total Outstanding (Cr.)": round(sdf["Total Outstanding (Cr.)"].sum(), 2),
                "Eff %": round((sdf["Paid (Cr.)"].sum() / sdf["CA (Cr.)"].sum() * 100)
                               if sdf["CA (Cr.)"].sum() else 0, 2),
                "Turn-up %": round((s_sdo["Txns"].sum() / s_sdo["Bills_Total"].sum() * 100)
                                   if s_sdo["Bills_Total"].sum() else 0, 2),
            }])], ignore_index=True)

            make_table(
                "🏢 SDO Name wise",
                sdf,
                list(sdf.columns),
                col_widths=[110, 95, 72, 72, 100, 45, 44]
            )

        if not df_paid.empty:
            story.append(PageBreak())

            story.append(Paragraph(
                f"<b>⚡ UPPCL Payment Dashboard — {year}-{month:02d}</b>",
                ParagraphStyle("sub2", parent=styles["Normal"], fontSize=10,
                               textColor=colors.HexColor("#1F4E78"),
                               alignment=1, spaceAfter=8)
            ))

            d = df_paid.copy()
            d["Day"] = d[DATE_COL].dt.day
            daily = (d.groupby("Day", as_index=False)
                       .agg(Amount=(AMOUNT_COL, "sum"), Txns=(AMOUNT_COL, "size")))
            daily["Amount (Cr.)"] = (daily["Amount"] / CRORE).round(2)
            daily["Amount (Cr.)"] = daily["Amount (Cr.)"].apply(lambda x: f"{x:.2f}")
            daily = daily[["Day", "Amount (Cr.)", "Txns"]]

            total_amount = round(sum(float(x) for x in daily["Amount (Cr.)"]), 2)
            total_row = {
                "Day": "TOTAL",
                "Amount (Cr.)": f"{total_amount:.2f}",
                "Txns": int(daily["Txns"].sum()),
            }
            daily = pd.concat([daily, pd.DataFrame([total_row])], ignore_index=True)

            make_table(
                "📅 Daily Collection (Day-wise)",
                daily,
                list(daily.columns),
                col_widths=[100, 200, 150]
            )

        story.append(Spacer(1, 14))
        story.append(Paragraph(
            f"<b>⚡ U.P. Power Corporation Limited</b> — Shakti Bhawan, Lucknow | "
            f"Toll Free: 1912 | www.uppcl.org | "
            f"<i>Generated: {datetime.now().strftime('%d-%b-%Y %H:%M:%S')}</i>",
            footer_style
        ))

        def add_page_number(canvas, doc_):
            canvas.saveState()
            canvas.setFont('Helvetica', 8)
            canvas.setFillColor(colors.HexColor("#666666"))
            canvas.drawRightString(page_size[0] - 28, 20,
                                   f"Page {doc_.page}")
            canvas.drawString(28, 20, "UPPCL Payment Dashboard")
            canvas.restoreState()

        doc.build(story, onFirstPage=add_page_number, onLaterPages=add_page_number)
        buffer.seek(0)
        return buffer.read()

    # ---------- TABS ----------
    def tab_target(df_paid, df_bill):
        st.subheader("🎯 SDO Name — Target vs Achievement")
        st.caption("Targets pre-defined hain (SDO Name wise)")

        if df_paid.empty and df_bill.empty:
            st.info("Data nahi mila.")
            return

        rows = []
        for grp in TARGET_GROUPS:
            sub_p = df_paid[df_paid[SDO_CODE_COL].isin(grp["codes"])] if not df_paid.empty else pd.DataFrame()
            sub_b = df_bill[df_bill[SDO_CODE_COL].isin(grp["codes"])] if not df_bill.empty else pd.DataFrame()
            actual = round(float(sub_p[AMOUNT_COL].sum()) / CRORE, 2) if not sub_p.empty else 0.0
            txns   = len(sub_p)
            bills  = len(sub_b)
            target = grp["target_cr"]
            ach    = round((actual / target * 100) if target else 0, 2)
            turn   = round((txns / bills * 100) if bills else 0, 2)
            rows.append({
                "SDO":           grp["label"],
                "Actual_Cr":     actual,
                "Target_Cr":     target,
                "Achievement_%": ach,
                "Txns":          txns,
                "Bills":         bills,
                "Turnup_%":      turn,
            })
        sdo_actual = pd.DataFrame(rows)

        total_actual = round(sdo_actual["Actual_Cr"].sum(), 2)
        total_target = TOTAL_TARGET_CR
        total_ach    = round((total_actual / total_target * 100) if total_target else 0, 2)
        total_txns   = int(sdo_actual["Txns"].sum())
        total_bills  = int(sdo_actual["Bills"].sum())
        total_turn   = round((total_txns / total_bills * 100) if total_bills else 0, 2)

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Total Actual", f"₹ {total_actual:,.2f} Cr.")
        c2.metric("Total Target", f"₹ {total_target:,.2f} Cr.")
        c3.metric("Achievement", f"{total_ach:.2f}%")
        c4.metric("Turn-up %", f"{total_turn:.2f}%")
        c5.metric("Transactions", f"{total_txns:,}")

        st.divider()

        st.markdown("### 📊 Live Progress")
        for _, row in sdo_actual.iterrows():
            pct = row["Achievement_%"]
            turn = row["Turnup_%"]
            icon = "🟢" if pct >= 90 else "🟡" if pct >= 60 else "🔴"
            c1, c2 = st.columns([3, 1])
            with c1:
                st.markdown(
                    f"**{icon} {row['SDO']}** — "
                    f"₹ {row['Actual_Cr']:.2f} / ₹ {row['Target_Cr']:.2f} Cr. "
                    f"({pct:.2f}%) | Turn-up: **{turn:.2f}%**"
                )
                st.progress(min(pct / 100, 1.0))
            with c2:
                gap = round(row["Target_Cr"] - row["Actual_Cr"], 2)
                st.metric("Gap", f"₹ {gap:+.2f} Cr.")

        st.divider()

        st.markdown("### 📋 Table")
        display = sdo_actual[["SDO", "Actual_Cr", "Target_Cr",
                              "Achievement_%", "Txns", "Bills",
                              "Turnup_%"]].rename(columns={
            "Actual_Cr":     "Actual (Cr.)",
            "Target_Cr":     "Target (Cr.)",
            "Achievement_%": "Ach %",
            "Txns":          "Paid",
            "Bills":         "Bills",
            "Turnup_%":      "Turn-up %",
        })
        totals = pd.DataFrame([{
            "SDO": "TOTAL",
            "Actual (Cr.)": total_actual,
            "Target (Cr.)": total_target,
            "Ach %":        total_ach,
            "Paid":         total_txns,
            "Bills":        total_bills,
            "Turn-up %":    total_turn,
        }])
        display_full = pd.concat([display, totals], ignore_index=True)
        # 🛠️ ARROW FIX
        display_full = _arrow_safe(display_full)

        st.dataframe(
            display_full.style.format({
                "Actual (Cr.)": "₹ {:,.2f}",
                "Target (Cr.)": "₹ {:,.2f}",
                "Ach %":        "{:.2f}%",
                "Paid":         "{:,}",
                "Bills":        "{:,}",
                "Turn-up %":    "{:.2f}%",
            }).apply(lambda r: ["font-weight: bold; background-color:#1a3a1a; color:white"] * len(r)
                     if r["SDO"] == "TOTAL" else [""] * len(r), axis=1)
              .background_gradient(subset=["Ach %"], cmap="RdYlGn", vmin=0, vmax=100)
              .background_gradient(subset=["Turn-up %"], cmap="RdYlGn", vmin=0, vmax=100),
            width='stretch', hide_index=True
        )

        st.markdown("### 📊 Target vs Actual")
        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=sdo_actual["SDO"], y=sdo_actual["Actual_Cr"],
            name="Actual", marker_color="#2E7D32",
            text=sdo_actual["Actual_Cr"].round(2), textposition="auto",
        ))
        fig.add_trace(go.Bar(
            x=sdo_actual["SDO"], y=sdo_actual["Target_Cr"],
            name="Target", marker_color="#B0BEC5",
            text=sdo_actual["Target_Cr"].round(2), textposition="auto",
        ))
        fig.update_layout(barmode="group", height=420,
                          xaxis_title="", yaxis_title="Amount (Cr.)")
        apply_theme(fig, height=420)
        st.plotly_chart(fig, width='stretch', key="target_chart")

        st.markdown("### 📈 Turn-up %")
        colors_turn = ["#2E7D32" if p >= 90 else "#F9A825" if p >= 60 else "#C62828"
                       for p in sdo_actual["Turnup_%"]]
        fig3 = go.Figure(go.Bar(
            x=sdo_actual["SDO"], y=sdo_actual["Turnup_%"],
            marker_color=colors_turn,
            text=sdo_actual["Turnup_%"].round(2).astype(str) + "%",
            textposition="auto",
        ))
        fig3.add_hline(y=100, line_dash="dash", line_color="gray",
                       annotation_text="100% Turn-up")
        fig3.update_layout(height=350, xaxis_title="", yaxis_title="Turn-up %")
        apply_theme(fig3, height=350)
        st.plotly_chart(fig3, width='stretch', key="target_turn")

        csv = display_full.to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ CSV", csv, "target_vs_achievement.csv",
                           "text/csv", key="dl_target")

    def tab_daily_calendar(df_paid, selected_month):
        st.subheader("📅 Daily Collection Calendar")
        if df_paid.empty:
            st.info("Koi paid data nahi mila.")
            return

        dfx = df_paid.copy()
        dfx["_DAY"] = dfx[DATE_COL].dt.day

        daily = (dfx.groupby("_DAY", as_index=False)
                    .agg(Amount=(AMOUNT_COL, "sum"),
                         Txns=(AMOUNT_COL, "size")))
        daily["Amount_Cr"] = (daily["Amount"] / CRORE).round(2)

        best_day  = daily.loc[daily["Amount"].idxmax()] if not daily.empty else None
        worst_day = daily.loc[daily["Amount"].idxmin()] if not daily.empty else None
        avg_daily = round(daily["Amount"].mean() / CRORE, 2) if not daily.empty else 0

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Best Day",
                  f"₹ {best_day['Amount_Cr']:,.2f} Cr." if best_day is not None else "—",
                  f"Day {int(best_day['_DAY'])}" if best_day is not None else "")
        c2.metric("Worst Day",
                  f"₹ {worst_day['Amount_Cr']:,.2f} Cr." if worst_day is not None else "—",
                  f"Day {int(worst_day['_DAY'])}" if worst_day is not None else "")
        c3.metric("Avg / Day", f"₹ {avg_daily:,.2f} Cr.")
        c4.metric("Days with Payment", f"{len(daily):,}")

        st.divider()

        st.markdown("### 🔥 Calendar Heatmap")
        year, month = int(selected_month[:4]), int(selected_month[5:7])
        days_in_month = pd.Timestamp(year=year, month=month, day=1).days_in_month
        first_weekday = pd.Timestamp(year=year, month=month, day=1).weekday()

        weeks = []
        week = [None] * first_weekday
        for d in range(1, days_in_month + 1):
            week.append(d)
            if len(week) == 7:
                weeks.append(week)
                week = []
        if week:
            while len(week) < 7:
                week.append(None)
            weeks.append(week)

        daily_map = dict(zip(daily["_DAY"], daily["Amount_Cr"]))
        z_data, text_data = [], []
        for wk in weeks:
            z_row, t_row = [], []
            for d in wk:
                if d is None:
                    z_row.append(None)
                    t_row.append("")
                else:
                    v = daily_map.get(d, 0)
                    z_row.append(v)
                    t_row.append(f"{d}<br>₹ {v:.2f} Cr.")
            z_data.append(z_row)
            text_data.append(t_row)

        fig = go.Figure(data=go.Heatmap(
            z=z_data, text=text_data, texttemplate="%{text}",
            colorscale="Greens", showscale=True,
            colorbar=dict(title="₹ Cr."), hoverinfo="text",
        ))
        fig.update_layout(
            height=60 + 60 * len(weeks),
            xaxis=dict(tickmode="array", tickvals=list(range(7)),
                       ticktext=["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"],
                       side="top"),
            yaxis=dict(autorange="reversed", showticklabels=False),
            margin=dict(l=20, r=20, t=40, b=20),
        )
        apply_theme(fig, height=60 + 60 * len(weeks))
        st.plotly_chart(fig, width='stretch', key="daily_calendar_heat")

        st.markdown("### 📊 Day-wise Collection")
        fig2 = px.bar(daily, x="_DAY", y="Amount_Cr",
                      text_auto=".2f", height=350,
                      color="Amount_Cr", color_continuous_scale="Blues")
        fig2.update_layout(xaxis_title="Day of Month",
                           yaxis_title="Amount (Cr.)",
                           coloraxis_showscale=False)
        apply_theme(fig2, height=350)
        st.plotly_chart(fig2, width='stretch', key="daily_bar")

        with st.expander("📋 Daily Table + CSV download"):
            show = daily[["_DAY", "Amount_Cr", "Txns"]].copy()
            show.columns = ["Day", "Amount (Cr.)", "Txns"]
            total_row = pd.DataFrame([{
                "Day": "TOTAL",
                "Amount (Cr.)": round(show["Amount (Cr.)"].sum(), 2),
                "Txns": int(show["Txns"].sum()),
            }])
            show_full = pd.concat([show, total_row], ignore_index=True)
            # 🛠️ ARROW FIX — Day column ko string banao
            show_full = _arrow_safe(show_full)

            st.dataframe(
                show_full.style
                  .format({"Amount (Cr.)": "₹ {:,.2f}", "Txns": "{:,}"})
                  .apply(lambda r: ["font-weight: bold; background-color:#1a3a1a; color:white"] * len(r)
                         if r["Day"] == "TOTAL" else [""] * len(r), axis=1),
                width='stretch', hide_index=True
            )
            csv = show_full.to_csv(index=False).encode("utf-8")
            st.download_button("⬇️ Daily CSV", csv,
                               f"daily_{selected_month}.csv",
                               "text/csv", key="dl_daily")

    def tab_summary_card(df_paid, df_bill, selected_month):
        st.subheader("🖨️ Print-Friendly Summary Card")

        total_paid = round(float(df_paid[AMOUNT_COL].sum()), 2) if not df_paid.empty else 0
        total_ca   = round(float(df_bill[CA_COL].sum()), 2)     if not df_bill.empty else 0
        total_outstanding = (round(float(df_bill[ARREAR_COL].sum()), 2)
                             if (not df_bill.empty and ARREAR_COL in df_bill.columns) else 0)
        txns       = len(df_paid)
        bills      = len(df_bill)
        efficiency = round((total_paid / total_ca * 100) if total_ca else 0, 2)
        turnup     = round((txns / bills * 100) if bills else 0, 2)

        card_html = f"""
        <div style="
            background: linear-gradient(135deg, #1F4E78 0%, #2E7D32 100%);
            border-radius: 24px;
            padding: 40px;
            color: white;
            font-family: 'Inter', sans-serif;
            box-shadow: 0 25px 70px rgba(0,0,0,0.6), 0 0 100px rgba(0,112,192,0.5);
            max-width: 900px;
            margin: 0 auto;
            border: 1px solid rgba(255,255,255,0.2);
        ">
            <h1 style="margin: 0; font-size: 30px; text-align: center;
                       background: linear-gradient(90deg, #fff, #F57C00, #fff);
                       -webkit-background-clip: text; -webkit-text-fill-color: transparent;
                       background-size: 200% auto; font-weight: 900; letter-spacing: 2px;">
                ⚡ UPPCL Payment Dashboard ⚡
            </h1>
            <p style="text-align: center; opacity: 0.9; margin-top: 12px;
                      font-size: 17px; letter-spacing: 3px; font-weight: 600;">
                📅 {selected_month}
            </p>
            <hr style="border: 0; border-top: 1px solid rgba(255,255,255,0.3); margin: 28px 0;">
            <table style="width: 100%; color: white; font-size: 16px; border-collapse: collapse;">
                <tr style="border-bottom: 1px solid rgba(255,255,255,0.1);">
                    <td style="padding: 16px 0; font-weight: 500;">📊 Current Assessment</td>
                    <td style="text-align: right; font-weight: 800; font-size: 24px; font-family: 'JetBrains Mono', monospace;">
                        ₹ {total_ca/CRORE:,.2f} Cr.
                    </td>
                </tr>
                <tr style="border-bottom: 1px solid rgba(255,255,255,0.1);">
                    <td style="padding: 16px 0; font-weight: 500;">✅ Total Paid</td>
                    <td style="text-align: right; font-weight: 800; font-size: 24px; color: #4CAF50; font-family: 'JetBrains Mono', monospace;">
                        ₹ {total_paid/CRORE:,.2f} Cr.
                    </td>
                </tr>
                <tr style="border-bottom: 1px solid rgba(255,255,255,0.1);">
                    <td style="padding: 16px 0; font-weight: 500;">📌 Total Outstanding</td>
                    <td style="text-align: right; font-weight: 800; font-size: 24px; color: #FF9800; font-family: 'JetBrains Mono', monospace;">
                        ₹ {total_outstanding/CRORE:,.2f} Cr.
                    </td>
                </tr>
                <tr style="border-bottom: 1px solid rgba(255,255,255,0.1);">
                    <td style="padding: 16px 0; font-weight: 500;">🎯 Collection Efficiency</td>
                    <td style="text-align: right; font-weight: 800; font-size: 24px; color: #FF9800; font-family: 'JetBrains Mono', monospace;">
                        {efficiency:.2f}%
                    </td>
                </tr>
                <tr style="border-bottom: 1px solid rgba(255,255,255,0.1);">
                    <td style="padding: 16px 0; font-weight: 500;">🔄 Turn-up %</td>
                    <td style="text-align: right; font-weight: 800; font-size: 24px; font-family: 'JetBrains Mono', monospace;">
                        {turnup:.2f}%
                    </td>
                </tr>
                <tr>
                    <td style="padding: 16px 0; font-weight: 500;">🔢 Transactions</td>
                    <td style="text-align: right; font-weight: 800; font-size: 24px; font-family: 'JetBrains Mono', monospace;">
                        {txns:,} / {bills:,}
                    </td>
                </tr>
            </table>
        </div>
        """
        st.markdown(card_html, unsafe_allow_html=True)

        st.divider()

        st.markdown("### 🎯 Target Achievement (SDO Name wise)")
        rows = []
        for grp in TARGET_GROUPS:
            sub_p = df_paid[df_paid[SDO_CODE_COL].isin(grp["codes"])] if not df_paid.empty else pd.DataFrame()
            sub_b = df_bill[df_bill[SDO_CODE_COL].isin(grp["codes"])] if not df_bill.empty else pd.DataFrame()
            actual = round(float(sub_p[AMOUNT_COL].sum()) / CRORE, 2) if not sub_p.empty else 0.0
            txns_g = len(sub_p)
            bills_g = len(sub_b)
            target = grp["target_cr"]
            ach    = round((actual / target * 100) if target else 0, 2)
            turn_g = round((txns_g / bills_g * 100) if bills_g else 0, 2)
            icon   = "🟢" if ach >= 90 else "🟡" if ach >= 60 else "🔴"
            rows.append({"SDO": grp["label"], "Actual": actual,
                         "Target": target, "Ach": ach,
                         "Turnup": turn_g, "Icon": icon})

        cols = st.columns(len(rows))
        for c, r in zip(cols, rows):
            with c:
                st.markdown(f"**{r['Icon']} {r['SDO']}**")
                st.metric("Actual", f"₹ {r['Actual']:.2f} Cr.")
                st.metric("Target", f"₹ {r['Target']:.2f} Cr.")
                st.metric("Ach %", f"{r['Ach']:.2f}%")
                st.metric("Turn-up", f"{r['Turnup']:.2f}%")

        st.markdown("---")
        st.markdown("#### 💰 **OVERALL TOTAL** (Independent Feeders included)")
        total_actual_sc = round(sum(r["Actual"] for r in rows), 2)
        total_target_sc = TOTAL_TARGET_CR
        total_ach_sc = round((total_actual_sc / total_target_sc * 100) if total_target_sc else 0, 2)

        tc1, tc2, tc3 = st.columns(3)
        tc1.metric("💰 Total Actual", f"₹ {total_actual_sc:.2f} Cr.")
        tc2.metric("🎯 Total Target", f"₹ {total_target_sc:.2f} Cr.")
        tc3.metric("📊 Overall Ach %", f"{total_ach_sc:.2f}%")

        st.divider()

        st.markdown("### 📤 Share Options")
        col1, col2, col3 = st.columns(3)

        with col1:
            share_text = (
                f"💰 UPPCL Payment Dashboard — {selected_month}\n\n"
                f"📊 Current Assessment: ₹ {total_ca/CRORE:,.2f} Cr.\n"
                f"✅ Total Paid: ₹ {total_paid/CRORE:,.2f} Cr.\n"
                f"📌 Total Outstanding: ₹ {total_outstanding/CRORE:,.2f} Cr.\n"
                f"🎯 Collection Efficiency: {efficiency:.2f}%\n"
                f"🔄 Turn-up %: {turnup:.2f}%\n"
                f"🔢 Transactions: {txns:,} / {bills:,} bills\n\n"
                f"SDO-wise:\n"
                + "\n".join([f"{r['Icon']} {r['SDO']}: "
                             f"Ach {r['Ach']:.2f}%, Turn-up {r['Turnup']:.2f}% "
                             f"(₹ {r['Actual']:.2f}/{r['Target']:.2f} Cr.)"
                             for r in rows])
                + f"\n\n🎯 TOTAL: Ach {total_ach_sc:.2f}% "
                  f"(₹ {total_actual_sc:.2f}/{total_target_sc:.2f} Cr.)"
            )
            st.download_button(
                "📋 Summary Text (WhatsApp)",
                share_text.encode("utf-8"),
                file_name=f"summary_{selected_month}.txt",
                mime="text/plain",
                key="dl_share_text",
            )

        with col2:
            st.info("PDF download upar 'Downloads' me hai")

        with col3:
            st.info("💡 Screenshot: Windows me `Win+Shift+S`")

    # ---------- MAIN DASHBOARD BODY ----------
    st.markdown(f"""
    <div class="hero-header">
        <div style="display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 25px;">
            <div>
                <h1 class="hero-title">⚡ UPPCL PAYMENT DASHBOARD</h1>
                <p class="hero-subtitle">
                    उत्तर प्रदेश पावर कॉर्पोरेशन लिमिटेड •
                    U.P. Power Corporation Limited
                    <span class="hero-badge"><span class="live-dot"></span> LIVE</span>
                </p>
            </div>
            <div style="text-align: right;">
                <p style="color: #B0C4DE; margin: 0; font-size: 14px; font-weight: 500;">
                    📅 {datetime.now().strftime('%d %B %Y')}
                </p>
                <p style="color: #F57C00; margin: 6px 0 0 0; font-size: 14px; font-weight: 700; font-family: 'JetBrains Mono', monospace;">
                    🕐 {datetime.now().strftime('%H:%M:%S')}
                </p>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    with st.sidebar:
        st.header("📂 File Upload")
        uploaded = st.file_uploader(
            "File chunein (.xlsx / .xls / .csv)",
            type=["xlsx", "xls", "csv"],
            key="dash_upload",
        )

    if uploaded is None:
        st.markdown("""
        <div style="text-align: center; padding: 70px 20px;">
            <div style="font-size: 90px; margin-bottom: 25px;
                        filter: drop-shadow(0 0 30px rgba(245,124,0,0.8));">⚡</div>
            <h2 style="color: white; font-size: 36px; margin-bottom: 18px;
                       font-weight: 900; letter-spacing: 1px;
                       text-shadow: 0 0 40px rgba(0,112,192,0.6);">
                Welcome to UPPCL Payment Dashboard
            </h2>
            <p style="color: #B0C4DE; font-size: 17px; max-width: 650px; margin: 0 auto;">
                👈 <b style="color: #F57C00;">Left sidebar</b> se apni Excel/CSV file upload karein
                aur apna payment data real-time me analyze karein.
            </p>
        </div>
        """, unsafe_allow_html=True)

        cols = st.columns(4)
        features = [
            ("📊", "Real-time KPIs", "Live metrics with animations"),
            ("🎯", "Target Tracking", "SDO-wise achievement"),
            ("📅", "Daily Calendar", "Heatmap visualization"),
            ("📄", "PDF & Excel", "Professional reports"),
        ]
        for col, (icon, title, desc) in zip(cols, features):
            with col:
                st.markdown(f"""
                <div class="kpi-card" style="padding: 28px 15px;">
                    <div class="kpi-icon" style="font-size: 40px;">{icon}</div>
                    <div class="kpi-label" style="font-size: 13px;">{title}</div>
                    <div style="color: #B0C4DE; font-size: 11px; margin-top: 10px; line-height: 1.5;">{desc}</div>
                </div>
                """, unsafe_allow_html=True)
        st.stop()

    t0 = datetime.now()
    file_bytes = uploaded.getvalue()

    with st.spinner("⚡ Fast loading..."):
        df_all = load_file(file_bytes, uploaded.name)

    if df_all.empty:
        st.error("File empty ya read nahi ho payi.")
        st.stop()

    load_ms = (datetime.now() - t0).total_seconds() * 1000
    st.success(
        f"✅ **{uploaded.name}** loaded — "
        f"{len(df_all):,} rows × {len(df_all.columns)} cols "
        f"in **{load_ms:.0f} ms**"
    )

    required = [BILL_DATE_COL, DATE_COL, CATEGORY_COL, AMOUNT_COL, CA_COL]
    missing = [c for c in required if c not in df_all.columns]
    if missing:
        st.error(f"❌ Ye columns file me nahi mile: {missing}")
        st.write("**File me ye columns hain:**", list(df_all.columns))
        st.stop()

    merge_tuple = tuple(sorted(TARIFF_MERGE.items()))
    df_paid_all, df_bill_all = prepare(df_all, merge_tuple)

    if df_paid_all.empty and df_bill_all.empty:
        st.warning("Koi valid record nahi mila.")
        st.stop()

    st.sidebar.markdown("---")
    st.sidebar.markdown("## 🎛️ Filters")

    all_months = sorted(set(df_paid_all["_YM"].dropna().unique()) |
                        set(df_bill_all["_YM"].dropna().unique()), reverse=True)
    if all_months:
        selected_month = st.sidebar.selectbox("📅 Month", all_months, index=0, key="dash_month")
    else:
        selected_month = datetime.today().strftime("%Y-%m")

    filters = {}

    if SDO_CODE_COL in df_all.columns:
        sdo_codes = sorted(df_all[SDO_CODE_COL].dropna().astype(str).unique().tolist())
        sdo_options = {f"{sdo_short(c)} ({c})": c for c in sdo_codes}
        chosen_labels = st.sidebar.multiselect(
            "🏢 SDO",
            options=list(sdo_options.keys()),
            key="dash_sdo",
        )
        chosen_codes = [sdo_options[l] for l in chosen_labels]
        if chosen_codes:
            filters[SDO_CODE_COL] = chosen_codes

    if CATEGORY_COL in df_all.columns:
        tariff_list = sorted(df_all[CATEGORY_COL].dropna().unique().tolist())
        filters[CATEGORY_COL] = st.sidebar.multiselect("📋 Tariff Type", tariff_list, key="dash_tariff")

    if "LOAD_CATEGORY" in df_all.columns:
        load_list = sorted([x for x in df_all["LOAD_CATEGORY"].dropna().unique()
                            if x in LOAD_ORDER])
        filters["LOAD_CATEGORY"] = st.sidebar.multiselect("⚡ Load Category", load_list, key="dash_load")

    if "PAYMENT_MODE" in df_all.columns:
        pm_list = sorted(df_all["PAYMENT_MODE"].dropna().unique().tolist())
        filters["PAYMENT_MODE"] = st.sidebar.multiselect("💳 Payment Mode", pm_list, key="dash_pm")

    if st.sidebar.button("🔄 Reset Filters", width='stretch', key="dash_reset"):
        st.rerun()

    df_paid = df_paid_all[df_paid_all["_YM"] == selected_month].copy()
    df_bill = df_bill_all[df_bill_all["_YM"] == selected_month].copy()

    for col, vals in filters.items():
        if vals:
            if col in df_paid.columns:
                df_paid = df_paid[df_paid[col].isin(vals)]
            if col in df_bill.columns:
                df_bill = df_bill[df_bill[col].isin(vals)]

    st.markdown(f"### 📅 Showing: **{selected_month}**")

    total_paid = round(float(df_paid[AMOUNT_COL].sum()), 2) if not df_paid.empty else 0.0
    total_ca   = round(float(df_bill[CA_COL].sum()), 2)     if not df_bill.empty else 0.0
    total_outstanding = (round(float(df_bill[ARREAR_COL].sum()), 2)
                         if (not df_bill.empty and ARREAR_COL in df_bill.columns) else 0.0)
    txns       = len(df_paid)
    bills      = len(df_bill)
    efficiency = round((total_paid / total_ca * 100) if total_ca else 0, 2)
    turnup     = round((txns / bills * 100) if bills else 0, 2)

    kpi_html = f"""
    <div class="kpi-grid">
        <div class="kpi-card">
            <span class="kpi-icon">📅</span>
            <div class="kpi-label">Month</div>
            <div class="kpi-value">{selected_month}</div>
        </div>
        <div class="kpi-card">
            <span class="kpi-icon">📊</span>
            <div class="kpi-label">Current Assessment</div>
            <div class="kpi-value">₹ {total_ca / CRORE:,.2f}</div>
            <div class="kpi-delta">Cr.</div>
        </div>
        <div class="kpi-card">
            <span class="kpi-icon">✅</span>
            <div class="kpi-label">Total Paid</div>
            <div class="kpi-value" style="color: #4CAF50;">₹ {total_paid / CRORE:,.2f}</div>
            <div class="kpi-delta">Cr.</div>
        </div>
        <div class="kpi-card">
            <span class="kpi-icon">📌</span>
            <div class="kpi-label">Total Outstanding</div>
            <div class="kpi-value" style="color: #FF9800;">₹ {total_outstanding / CRORE:,.2f}</div>
            <div class="kpi-delta" style="color: #FF9800;">Cr.</div>
        </div>
        <div class="kpi-card">
            <span class="kpi-icon">🎯</span>
            <div class="kpi-label">Efficiency</div>
            <div class="kpi-value" style="color: #F57C00;">{efficiency:.2f}%</div>
        </div>
        <div class="kpi-card">
            <span class="kpi-icon">🔄</span>
            <div class="kpi-label">Turn-up %</div>
            <div class="kpi-value">{turnup:.2f}%</div>
        </div>
        <div class="kpi-card">
            <span class="kpi-icon">💰</span>
            <div class="kpi-label">Transactions</div>
            <div class="kpi-value">{txns:,}</div>
        </div>
        <div class="kpi-card">
            <span class="kpi-icon">📄</span>
            <div class="kpi-label">Bills</div>
            <div class="kpi-value">{bills:,}</div>
        </div>
    </div>
    """
    st.markdown(kpi_html, unsafe_allow_html=True)

    st.markdown("### 📥 Downloads")
    dl1, dl2, _ = st.columns([1, 1, 3])
    with dl1:
        try:
            excel_bytes = build_excel(df_paid, df_bill,
                                      int(selected_month[:4]),
                                      int(selected_month[5:7]))
            st.download_button(
                "📊 Excel Report", excel_bytes,
                file_name=f"dashboard_{selected_month}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_excel_all",
            )
        except Exception as e:
            st.error(f"Excel error: {e}")
    with dl2:
        try:
            pdf_bytes = build_pdf(df_paid, df_bill,
                                  int(selected_month[:4]),
                                  int(selected_month[5:7]))
            st.download_button(
                "📄 PDF Report", pdf_bytes,
                file_name=f"dashboard_{selected_month}.pdf",
                mime="application/pdf", key="dl_pdf_all",
            )
        except Exception as e:
            st.error(f"PDF error: {e}")

    st.divider()

    tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
        "📋 Tariff Type",
        "⚡ Load wise",
        "🏢 SDO Name wise",
        "📈 Efficiency + Turn-up",
        "📅 Daily Calendar",
        "🎯 Target vs Achievement",
        "🏆 Top/Bottom Consumers",
        "🖨️ Summary Card",
    ])

    with tab1:
        show_4col_table(df_paid, df_bill, CATEGORY_COL,
                        "Tariff Type wise", key="tariff")
        with st.expander("🥧 Share by Tariff (Paid)"):
            s = make_summary(df_paid, df_bill, CATEGORY_COL)
            if not s.empty:
                fig = px.pie(s, names=CATEGORY_COL, values="Total_Cr", hole=0.4)
                apply_theme(fig, height=400)
                st.plotly_chart(fig, width='stretch', key="pie_tariff")

    with tab2:
        if not df_paid.empty and "LOAD_CATEGORY" in df_paid.columns:
            df_p_load = df_paid.dropna(subset=["LOAD_CATEGORY"])
            df_b_load = df_bill.dropna(subset=["LOAD_CATEGORY"]) if not df_bill.empty else df_bill
            show_4col_table(df_p_load, df_b_load, "LOAD_CATEGORY",
                            "Load Category wise", key="load", order=LOAD_ORDER)
            with st.expander("🥧 Share by Load Category"):
                s = make_summary(df_p_load, df_b_load, "LOAD_CATEGORY",
                                 order=LOAD_ORDER)
                if not s.empty:
                    fig = px.pie(s, names="LOAD_CATEGORY",
                                 values="Total_Cr", hole=0.4)
                    apply_theme(fig, height=400)
                    st.plotly_chart(fig, width='stretch',
                                    key="pie_load")
        else:
            st.info("SANCTION_LOAD column nahi mila.")

    with tab3:
        if SDO_CODE_COL in df_all.columns:
            show_sdo_name_table(df_paid, df_bill, key="sdo_name")
        else:
            st.info("SDO_CODE column nahi mila.")

    with tab4:
        st.markdown("### 📈 Collection Efficiency (Paid / CA × 100)")
        st.caption("हरा = अच्छा collection, लाल = कम")
        show_efficiency_table(df_paid, df_bill, CATEGORY_COL,
                              "Tariff-wise", key="eff_tariff")

        if not df_paid.empty and "LOAD_CATEGORY" in df_paid.columns:
            df_p_load = df_paid.dropna(subset=["LOAD_CATEGORY"])
            df_b_load = df_bill.dropna(subset=["LOAD_CATEGORY"]) if not df_bill.empty else df_bill
            show_efficiency_table(df_p_load, df_b_load, "LOAD_CATEGORY",
                                  "Load-wise", key="eff_load", order=LOAD_ORDER)

        if SDO_CODE_COL in df_all.columns:
            show_sdo_efficiency_table(df_paid, df_bill, key="eff_sdo_name")

        if "PAYMENT_MODE" in df_all.columns:
            show_efficiency_table(df_paid, df_bill, "PAYMENT_MODE",
                                  "Payment Mode-wise", key="eff_pm")

    with tab5:
        tab_daily_calendar(df_paid, selected_month)

    with tab6:
        tab_target(df_paid, df_bill)

    with tab7:
        if df_paid.empty or ACCT_COL not in df_paid.columns:
            st.info("ACCT_ID / NAME column nahi mila.")
        else:
            grp = [ACCT_COL, NAME_COL] if NAME_COL in df_paid.columns else [ACCT_COL]
            cons = (df_paid.groupby(grp, as_index=False)
                            .agg(Total_Paid=(AMOUNT_COL, "sum"),
                                 Txns=(AMOUNT_COL, "size")))
            cons["Total_Cr"] = (cons["Total_Paid"] / CRORE).round(2)
            top_n = st.slider("Kitne dikhayein?", 5, 50, 10, key="topn")
            c1, c2 = st.columns(2)
            with c1:
                st.subheader(f"🏆 Top {top_n} Consumers")
                top = cons.sort_values("Total_Paid", ascending=False).head(top_n)
                top_show = top[grp + ["Total_Cr", "Txns"]].copy()
                top_show.columns = (["ACCT_ID", "NAME", "Total (Cr.)", "Txns"]
                                    if len(grp) == 2
                                    else ["ACCT_ID", "Total (Cr.)", "Txns"])
                total_top = {"ACCT_ID": "TOTAL"}
                if len(grp) == 2:
                    total_top["NAME"] = ""
                total_top["Total (Cr.)"] = round(top_show["Total (Cr.)"].sum(), 2)
                total_top["Txns"] = int(top_show["Txns"].sum())
                top_show = pd.concat([top_show, pd.DataFrame([total_top])],
                                     ignore_index=True)
                # 🛠️ ARROW FIX
                top_show = _arrow_safe(top_show)
                st.dataframe(
                    top_show.style.format({
                        "Total (Cr.)": "₹ {:,.2f}", "Txns": "{:,}"})
                      .apply(lambda r: ["font-weight: bold; background-color:#1a3a1a; color:white"] * len(r)
                             if r["ACCT_ID"] == "TOTAL" else [""] * len(r), axis=1),
                    width='stretch', hide_index=True, height=420
                )
            with c2:
                st.subheader(f"🔻 Bottom {top_n} Consumers")
                bot = cons[cons["Total_Paid"] > 0].sort_values("Total_Paid").head(top_n)
                bot_show = bot[grp + ["Total_Cr", "Txns"]].copy()
                bot_show.columns = (["ACCT_ID", "NAME", "Total (Cr.)", "Txns"]
                                    if len(grp) == 2
                                    else ["ACCT_ID", "Total (Cr.)", "Txns"])
                total_bot = {"ACCT_ID": "TOTAL"}
                if len(grp) == 2:
                    total_bot["NAME"] = ""
                total_bot["Total (Cr.)"] = round(bot_show["Total (Cr.)"].sum(), 2)
                total_bot["Txns"] = int(bot_show["Txns"].sum())
                bot_show = pd.concat([bot_show, pd.DataFrame([total_bot])],
                                     ignore_index=True)
                # 🛠️ ARROW FIX
                bot_show = _arrow_safe(bot_show)
                st.dataframe(
                    bot_show.style.format({
                        "Total (Cr.)": "₹ {:,.2f}", "Txns": "{:,}"})
                      .apply(lambda r: ["font-weight: bold; background-color:#1a3a1a; color:white"] * len(r)
                             if r["ACCT_ID"] == "TOTAL" else [""] * len(r), axis=1),
                    width='stretch', hide_index=True, height=420
                )
            with st.expander("⬇️ Download consumers list (CSV)"):
                full = cons.sort_values("Total_Paid", ascending=False).copy()
                full["Total (Cr.)"] = full["Total_Cr"]
                full_show = full[grp + ["Total (Cr.)", "Txns"]].copy()
                full_show.columns = (["ACCT_ID", "NAME", "Total (Cr.)", "Txns"]
                                     if len(grp) == 2
                                     else ["ACCT_ID", "Total (Cr.)", "Txns"])
                total_all = {"ACCT_ID": "TOTAL"}
                if len(grp) == 2:
                    total_all["NAME"] = ""
                total_all["Total (Cr.)"] = round(full_show["Total (Cr.)"].sum(), 2)
                total_all["Txns"] = int(full_show["Txns"].sum())
                full_show = pd.concat([full_show, pd.DataFrame([total_all])],
                                      ignore_index=True)
                # 🛠️ ARROW FIX
                full_show = _arrow_safe(full_show)
                csv = full_show.to_csv(index=False).encode("utf-8")
                st.download_button("Consumers CSV", csv,
                                   "consumers.csv", "text/csv", key="dl_cons")

    with tab8:
        tab_summary_card(df_paid, df_bill, selected_month)

    st.markdown(f"""
    <div class="uppcl-footer">
        <h4>⚡ U.P. POWER CORPORATION LIMITED</h4>
        <p>Shakti Bhawan, 14 Ashok Marg, Lucknow - 226001</p>
        <p>📞 0522-2286618 | 🌐 www.uppcl.org | ☎️ Toll Free: 1912</p>
        <p style="margin-top: 18px; font-size: 11px; opacity: 0.7;">
            Dashboard Generated: {datetime.now().strftime('%d %B %Y, %H:%M:%S')}
        </p>
    </div>
    """, unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════
#  🚀 ROUTER
# ═══════════════════════════════════════════════════════════
if page == "merger":
    render_merger()
elif page == "dashboard":
    render_dashboard()
