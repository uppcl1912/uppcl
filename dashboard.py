# dashboard.py — PERFORMANCE OPTIMIZED
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
            st.dataframe(merged_df.head(preview_rows), use_container_width=True)

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
                if st.button("📥 Generate CSV", use_container_width=True, key="mg_gen_csv"):
                    st.session_state["mg_csv_bytes"] = to_csv_bytes(merged_df)
                if "mg_csv_bytes" in st.session_state:
                    st.download_button(
                        "⬇️ Download CSV",
                        data=st.session_state["mg_csv_bytes"],
                        file_name="merged_billed_unbilled.csv",
                        mime="text/csv",
                        use_container_width=True,
                        type="primary",
                        key="mg_dl_csv",
                    )
            with dl_col2:
                if st.button("📥 Generate Excel", use_container_width=True, key="mg_gen_xlsx"):
                    with st.spinner("Excel ban raha hai..."):
                        st.session_state["mg_excel_bytes"] = to_excel_bytes(merged_df)
                if "mg_excel_bytes" in st.session_state:
                    st.download_button(
                        "⬇️ Download Excel",
                        data=st.session_state["mg_excel_bytes"],
                        file_name="merged_billed_unbilled.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="mg_dl_xlsx",
                    )
        except Exception as e:
            st.error(f"❌ Error: {e}")
    else:
        st.info("👆 Upar dono files upload karo (Billed + Unbilled)")

    st.markdown("---")
    st.caption("💡 Tip: Sidebar mein rename rules customize kar sakte ho.")


# ═══════════════════════════════════════════════════════════
#  📦 PAGE 2 — UPPCL PAYMENT DASHBOARD (OPTIMIZED)
# ═══════════════════════════════════════════════════════════
def render_dashboard():
    import pandas as pd
    import plotly.express as px
    import plotly.graph_objects as go
    from datetime import datetime
    import io
    import traceback

    pd.set_option('display.float_format', '{:.2f}'.format)
    pd.set_option('display.precision', 2)
    pd.set_option('display.max_columns', None)

    # ---------- CSS (HALKA) ----------
    st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700;900&display=swap');
        html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
        .stApp {
            background: linear-gradient(135deg, #05080f 0%, #0a0e27 50%, #05080f 100%);
            background-attachment: fixed;
        }
        .main .block-container { padding: 1rem 2rem 2rem 2rem; max-width: 100%; }
        .hero-header {
            background: linear-gradient(135deg, rgba(0,51,102,0.85) 0%, rgba(0,82,155,0.75) 50%, rgba(245,124,0,0.65) 100%);
            border: 1px solid rgba(255,255,255,0.18);
            border-radius: 24px;
            padding: 30px 40px;
            margin-bottom: 25px;
            box-shadow: 0 25px 70px rgba(0,0,0,0.6), 0 0 120px rgba(0,112,192,0.4);
        }
        .hero-title {
            font-size: 42px; font-weight: 900;
            background: linear-gradient(90deg, #FFFFFF, #F57C00, #4FC3F7, #FFFFFF);
            -webkit-background-clip: text; -webkit-text-fill-color: transparent;
            letter-spacing: 3px; margin: 0;
        }
        .hero-subtitle { color: #B0C4DE; font-size: 14px; margin-top: 8px; letter-spacing: 1.5px; }
        .hero-badge {
            display: inline-block; background: linear-gradient(135deg, #F57C00, #FF9800);
            color: white; padding: 6px 16px; border-radius: 24px;
            font-size: 12px; font-weight: 800; margin-left: 12px;
        }
        .kpi-grid {
            display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
            gap: 15px; margin: 25px 0;
        }
        .kpi-card {
            background: linear-gradient(135deg, rgba(0,51,102,0.85), rgba(0,82,155,0.65));
            border: 1px solid rgba(255,255,255,0.15); border-radius: 18px;
            padding: 20px 14px; text-align: center;
            box-shadow: 0 10px 30px rgba(0,0,0,0.4);
        }
        .kpi-icon { font-size: 28px; margin-bottom: 8px; display: block; }
        .kpi-label {
            color: #B0C4DE; font-size: 10px; font-weight: 700;
            text-transform: uppercase; letter-spacing: 1.5px; margin-bottom: 8px;
        }
        .kpi-value {
            font-size: 24px; font-weight: 900; color: #FFFFFF;
            text-shadow: 0 0 20px rgba(0,112,192,0.8);
        }
        [data-testid="stSidebar"] {
            background: linear-gradient(180deg, #05080f, #0a0e27);
            border-right: 1px solid rgba(0,112,192,0.3);
        }
        [data-testid="stSidebar"] * { color: #E8F0F8 !important; }
        .stTabs [data-baseweb="tab-list"] {
            gap: 8px; background: rgba(0,51,102,0.35); padding: 8px;
            border-radius: 15px; margin-bottom: 20px;
        }
        .stTabs [data-baseweb="tab"] {
            background: rgba(255,255,255,0.04); border-radius: 10px;
            padding: 10px 18px; font-weight: 700; color: #B0C4DE;
        }
        .stTabs [aria-selected="true"] {
            background: linear-gradient(135deg, #00529B, #F57C00) !important;
            color: white !important;
        }
        [data-testid="stMetric"] {
            background: linear-gradient(135deg, rgba(0,51,102,0.75), rgba(0,82,155,0.55));
            border-left: 5px solid #F57C00; padding: 15px; border-radius: 12px;
        }
        [data-testid="stMetricValue"] { color: white !important; font-weight: 900; }
        .stApp h1, .stApp h2, .stApp h3 { color: white !important; }
        .uppcl-footer {
            background: linear-gradient(135deg, rgba(0,51,102,0.95), rgba(0,82,155,0.75));
            border-radius: 18px; padding: 25px; margin: 30px 0 15px 0; text-align: center;
        }
        .uppcl-footer h4 { color: #F57C00 !important; margin: 0 0 10px 0; letter-spacing: 3px; }
        .uppcl-footer p { color: #B0C4DE; margin: 5px 0; font-size: 13px; }
        #MainMenu { visibility: hidden; } footer { visibility: hidden; } header { visibility: hidden; }
        @media screen and (max-width: 768px) {
            .hero-title { font-size: 24px; }
            .kpi-value { font-size: 18px; }
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
    CA_COL          = "CA"
    ARREAR_COL      = "AMOUNT_PAYABLE"
    SDO_CODE_COL    = "SDO_CODE"
    SDO_NAME_COL    = "SDO_NAME"
    SUPPLY_TYPE_COL = "SUPPLY_TYPE"
    ACCT_COL        = "ACCT_ID"
    NAME_COL        = "NAME"

    USECOLS = [
        BILL_DATE_COL, DATE_COL, CATEGORY_COL, AMOUNT_COL, CA_COL,
        ARREAR_COL, LOAD_COL, SUPPLY_TYPE_COL,
        SDO_CODE_COL, ACCT_COL, NAME_COL, "PAYMENT_MODE",
    ]

    TARIFF_MERGE = {"HV1": "HV", "HV2": "HV"}
    HV_TARIFFS = {"HV1", "HV2"}
    OTHERS_TARIFFS = {"LMV3", "LMV4", "LMV5", "LMV7", "LMV8", "LMV10"}
    LMV4_EXCEPTION_TARIFF       = "LMV4"
    LMV4_EXCEPTION_SUPPLY_TYPES = {"46", "47"}

    LOAD_ORDER = [
        "HV CONNECTION", ">= 10 KW/KVA/BHP",
        "5-9 KW/KVA/BHP", "< 5 KW/KVA/BHP", "Others",
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

    def sdo_short(code):
        if pd.isna(code): return code
        c = str(code).strip()
        return SDO_NAME_MAP.get(c, c)

    TARGET_GROUPS = [
        {"label": "SDO Lal Fatak", "codes": ["SDO3419221"], "target_cr": 11.36},
        {"label": "SDO Faridpur",  "codes": ["SDO3419222", "SDO3419229"], "target_cr": 5.88},
        {"label": "SDO Izzat Nagar","codes": ["SDO3419223"], "target_cr": 8.02},
        {"label": "SDO Bhuta",     "codes": ["SDO3419224"], "target_cr": 2.40},
    ]
    INDEPENDENT_FEEDERS_TARGET_CR = 7.65
    TOTAL_TARGET_CR = round(sum(g["target_cr"] for g in TARGET_GROUPS) + INDEPENDENT_FEEDERS_TARGET_CR, 2)

    # ---------- THEME ----------
    CHART_LAYOUT = dict(
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,51,102,0.15)',
        font=dict(color='#E8F0F8', family='Inter, sans-serif', size=11),
        xaxis=dict(gridcolor='rgba(0,112,192,0.2)'),
        yaxis=dict(gridcolor='rgba(0,112,192,0.2)'),
        margin=dict(l=40, r=20, t=50, b=40),
    )

    def apply_theme(fig, height=350):
        fig.update_layout(**CHART_LAYOUT, height=height)
        return fig

    # ---------- LOAD FILE ----------
    @st.cache_data(show_spinner=False, max_entries=2, ttl=3600)
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
    @st.cache_data(show_spinner="⚡ Data process ho raha hai...", max_entries=2, ttl=3600)
    def prepare(df: pd.DataFrame, merge_tuple: tuple):
        try:
            keep = [c for c in [BILL_DATE_COL, DATE_COL, CATEGORY_COL, AMOUNT_COL,
                                CA_COL, ARREAR_COL, LOAD_COL, SUPPLY_TYPE_COL,
                                SDO_CODE_COL, ACCT_COL, NAME_COL, "PAYMENT_MODE"]
                    if c in df.columns]
            df = df[keep].copy()

            df[DATE_COL]      = pd.to_datetime(df[DATE_COL], errors="coerce")
            df[BILL_DATE_COL] = pd.to_datetime(df[BILL_DATE_COL], errors="coerce")

            df[ARREAR_COL] = pd.to_numeric(df.get(ARREAR_COL, 0), errors="coerce").fillna(0)
            df[CA_COL]     = pd.to_numeric(df.get(CA_COL, 0), errors="coerce").fillna(0)
            df[AMOUNT_COL] = pd.to_numeric(df.get(AMOUNT_COL, 0), errors="coerce").fillna(0)

            tariff = df[CATEGORY_COL].astype(str).str.strip().str.upper()
            supply = (df[SUPPLY_TYPE_COL].astype(str).str.strip().str.upper()
                      if SUPPLY_TYPE_COL in df.columns
                      else pd.Series("", index=df.index))
            load = (pd.to_numeric(df[LOAD_COL], errors="coerce")
                    if LOAD_COL in df.columns
                    else pd.Series(pd.NA, index=df.index, dtype="float64"))

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
            df["SDO_NAME"] = (df[SDO_CODE_COL].apply(sdo_short)
                              if SDO_CODE_COL in df.columns else "Unknown")

            if merge_tuple:
                df[CATEGORY_COL] = df[CATEGORY_COL].replace(dict(merge_tuple))

            df_paid = df.dropna(subset=[DATE_COL]).copy()
            df_paid = df_paid[df_paid[AMOUNT_COL] > 0]
            df_paid["_YM"] = df_paid[DATE_COL].dt.to_period("M").astype(str)

            df_bill = df.dropna(subset=[BILL_DATE_COL]).copy()
            df_bill["_YM"] = df_bill[BILL_DATE_COL].dt.to_period("M").astype(str)

            return df_paid, df_bill

        except Exception as e:
            st.error(f"❌ prepare() fail: {type(e).__name__}: {e}")
            st.code(traceback.format_exc(), language="python")
            st.stop()

    # ---------- SUMMARY ----------
    def make_summary(df_paid, df_bill, group_col, order=None):
        if not df_bill.empty and CA_COL in df_bill.columns:
            agg_dict = {
                "CA_Total":    (CA_COL, "sum"),
                "Bills_Total": (CA_COL, "size"),
                "Arrear_Total": (ARREAR_COL, "sum") if ARREAR_COL in df_bill.columns
                                else (CA_COL, lambda x: 0),
            }
            ca = df_bill.groupby(group_col, as_index=False, observed=True).agg(**agg_dict)
        else:
            ca = pd.DataFrame({group_col: [], "CA_Total": [], "Bills_Total": [], "Arrear_Total": []})

        if not df_paid.empty:
            paid = df_paid.groupby(group_col, as_index=False, observed=True).agg(
                Total_Paid=(AMOUNT_COL, "sum"), Txns=(AMOUNT_COL, "size"))
        else:
            paid = pd.DataFrame({group_col: [], "Total_Paid": [], "Txns": []})

        s = pd.merge(ca, paid, on=group_col, how="outer").fillna(0)
        s["CA_Cr"]        = (s["CA_Total"] / CRORE).round(2)
        s["Arrear_Cr"]    = (s["Arrear_Total"] / CRORE).round(2)
        s["Total_Cr"]     = (s["Total_Paid"] / CRORE).round(2)
        s["Txns"]         = s["Txns"].astype(int)
        s["Bills_Total"]  = s["Bills_Total"].astype(int)
        s["Efficiency_%"] = (s["Total_Paid"] / s["CA_Total"].replace(0, pd.NA) * 100).round(2).fillna(0)
        s["Turnup_%"]     = (s["Txns"] / s["Bills_Total"].replace(0, pd.NA) * 100).round(2).fillna(0)

        if order:
            s["__o"] = s[group_col].map({c: i for i, c in enumerate(order)})
            s = s.sort_values("__o", na_position="last").drop(columns="__o")
        else:
            s = s.sort_values("Total_Paid", ascending=False)
        return s

    def _total_row(group_col, s):
        tot_ca   = round(s["CA_Cr"].sum(), 2)
        tot_arr  = round(s["Arrear_Cr"].sum(), 2)
        tot_paid = round(s["Total_Cr"].sum(), 2)
        tot_txns = int(s["Txns"].sum())
        tot_bills = int(s["Bills_Total"].sum())
        return {
            group_col: "TOTAL",
            "CA_Cr": tot_ca,
            "Arrear_Cr": tot_arr,
            "Total_Cr": tot_paid,
            "Txns": tot_txns,
            "Bills_Total": tot_bills,
            "Efficiency_%": round((tot_paid / tot_ca * 100) if tot_ca else 0, 2),
            "Turnup_%": round((tot_txns / tot_bills * 100) if tot_bills else 0, 2),
        }

    def show_4col_table(df_paid, df_bill, group_col, title, key, order=None, chart=True):
        st.subheader(title)
        if df_paid.empty and df_bill.empty:
            st.info("Koi data nahi.")
            return
        s = make_summary(df_paid, df_bill, group_col, order=order)
        total = _total_row(group_col, s)

        display = s[[group_col, "CA_Cr", "Total_Cr", "Arrear_Cr", "Txns",
                     "Bills_Total", "Efficiency_%", "Turnup_%"]].rename(columns={
            "CA_Cr": "CA (Cr.)", "Total_Cr": "Paid (Cr.)",
            "Arrear_Cr": "Total Outstanding (Cr.)", "Txns": "Paid Count",
            "Bills_Total": "Bills", "Efficiency_%": "Eff %", "Turnup_%": "Turn-up %",
        })
        display[group_col] = display[group_col].astype(str)

        totals = pd.DataFrame([{
            group_col: "TOTAL",
            "CA (Cr.)": total["CA_Cr"], "Paid (Cr.)": total["Total_Cr"],
            "Total Outstanding (Cr.)": total["Arrear_Cr"],
            "Paid Count": total["Txns"], "Bills": total["Bills_Total"],
            "Eff %": total["Efficiency_%"], "Turn-up %": total["Turnup_%"],
        }])
        totals[group_col] = totals[group_col].astype(str)
        full = pd.concat([display, totals], ignore_index=True)

        col1, col2 = st.columns([1.7, 1])
        with col1:
            st.dataframe(
                full.style.format({
                    "CA (Cr.)": "₹ {:,.2f}", "Paid (Cr.)": "₹ {:,.2f}",
                    "Total Outstanding (Cr.)": "₹ {:,.2f}",
                    "Paid Count": "{:,}", "Bills": "{:,}",
                    "Eff %": "{:.2f}%", "Turn-up %": "{:.2f}%",
                }).apply(lambda r: ["font-weight: bold; background-color:#1a3a1a; color:white"] * len(r)
                         if r[group_col] == "TOTAL" else [""] * len(r), axis=1)
                  .background_gradient(subset=["Eff %"], cmap=EFFICIENCY_CMAP, vmin=0, vmax=100)
                  .background_gradient(subset=["Turn-up %"], cmap=TURNUP_CMAP, vmin=0, vmax=100),
                use_container_width=True, hide_index=True, height=400
            )
        with col2:
            if chart and not s.empty:
                fig = px.bar(s, x=group_col, y="Turnup_%", text_auto=".2f",
                             color="Turnup_%", color_continuous_scale="RdYlGn",
                             range_color=(0, 100), height=400)
                fig.update_layout(xaxis_title="", yaxis_title="Turn-up %",
                                  coloraxis_showscale=False)
                apply_theme(fig, height=400)
                st.plotly_chart(fig, use_container_width=True, key=f"bar_{key}")
        return s

    def show_sdo_name_table(df_paid, df_bill, key="sdo"):
        st.subheader("🏢 SDO Name wise")
        if df_paid.empty and df_bill.empty:
            st.info("Koi data nahi."); return

        if not df_bill.empty and CA_COL in df_bill.columns:
            ca = df_bill.groupby([SDO_NAME_COL, SDO_CODE_COL], as_index=False, observed=True).agg(
                CA_Total=(CA_COL, "sum"), Bills_Total=(CA_COL, "size"),
                Arrear_Total=(ARREAR_COL, "sum") if ARREAR_COL in df_bill.columns
                            else (CA_COL, lambda x: 0))
        else:
            ca = pd.DataFrame({SDO_NAME_COL: [], SDO_CODE_COL: [],
                              "CA_Total": [], "Bills_Total": [], "Arrear_Total": []})

        if not df_paid.empty:
            paid = df_paid.groupby([SDO_NAME_COL, SDO_CODE_COL], as_index=False, observed=True).agg(
                Total_Paid=(AMOUNT_COL, "sum"), Txns=(AMOUNT_COL, "size"))
        else:
            paid = pd.DataFrame({SDO_NAME_COL: [], SDO_CODE_COL: [],
                                "Total_Paid": [], "Txns": []})

        s = pd.merge(ca, paid, on=[SDO_NAME_COL, SDO_CODE_COL], how="outer").fillna(0)
        s["CA (Cr.)"]   = (s["CA_Total"] / CRORE).round(2)
        s["Paid (Cr.)"] = (s["Total_Paid"] / CRORE).round(2)
        s["Total Outstanding (Cr.)"] = (s["Arrear_Total"] / CRORE).round(2)
        s["Paid Count"] = s["Txns"].astype(int)
        s["Bills"]      = s["Bills_Total"].astype(int)
        s["Eff %"]      = (s["Total_Paid"] / s["CA_Total"].replace(0, pd.NA) * 100).round(2).fillna(0)
        s["Turn-up %"]  = (s["Txns"] / s["Bills_Total"].replace(0, pd.NA) * 100).round(2).fillna(0)
        s = s.sort_values("Paid (Cr.)", ascending=False)

        display = s[[SDO_NAME_COL, SDO_CODE_COL, "CA (Cr.)", "Paid (Cr.)",
                     "Total Outstanding (Cr.)", "Paid Count", "Bills", "Eff %", "Turn-up %"]].copy()
        display.columns = ["SDO Name", "SDO Code", "CA (Cr.)", "Paid (Cr.)",
                           "Total Outstanding (Cr.)", "Paid Count", "Bills", "Eff %", "Turn-up %"]
        display["SDO Name"] = display["SDO Name"].astype(str)
        display["SDO Code"] = display["SDO Code"].astype(str)

        tot_ca = round(s["CA (Cr.)"].sum(), 2); tot_paid = round(s["Paid (Cr.)"].sum(), 2)
        tot_arr = round(s["Total Outstanding (Cr.)"].sum(), 2)
        tot_txns = int(s["Paid Count"].sum()); tot_bills = int(s["Bills"].sum())
        tot_eff = round((tot_paid / tot_ca * 100) if tot_ca else 0, 2)
        tot_turn = round((tot_txns / tot_bills * 100) if tot_bills else 0, 2)

        totals = pd.DataFrame([{
            "SDO Name": "TOTAL", "SDO Code": "",
            "CA (Cr.)": tot_ca, "Paid (Cr.)": tot_paid,
            "Total Outstanding (Cr.)": tot_arr,
            "Paid Count": tot_txns, "Bills": tot_bills,
            "Eff %": tot_eff, "Turn-up %": tot_turn,
        }])
        full = pd.concat([display, totals], ignore_index=True)

        col1, col2 = st.columns([1.7, 1])
        with col1:
            st.dataframe(
                full.style.format({
                    "CA (Cr.)": "₹ {:,.2f}", "Paid (Cr.)": "₹ {:,.2f}",
                    "Total Outstanding (Cr.)": "₹ {:,.2f}",
                    "Paid Count": "{:,}", "Bills": "{:,}",
                    "Eff %": "{:.2f}%", "Turn-up %": "{:.2f}%",
                }).apply(lambda r: ["font-weight: bold; background-color:#1a3a1a; color:white"] * len(r)
                         if r["SDO Name"] == "TOTAL" else [""] * len(r), axis=1)
                  .background_gradient(subset=["Eff %"], cmap=EFFICIENCY_CMAP, vmin=0, vmax=100)
                  .background_gradient(subset=["Turn-up %"], cmap=TURNUP_CMAP, vmin=0, vmax=100),
                use_container_width=True, hide_index=True, height=400
            )
        with col2:
            if not s.empty:
                fig = px.bar(s, x=SDO_NAME_COL, y="Turn-up %", text_auto=".2f",
                             color="Turn-up %", color_continuous_scale="RdYlGn",
                             range_color=(0, 100), height=400)
                fig.update_layout(xaxis_title="", yaxis_title="Turn-up %",
                                  coloraxis_showscale=False)
                apply_theme(fig, height=400)
                st.plotly_chart(fig, use_container_width=True, key=f"bar_{key}")
        return s

    def show_efficiency_table(df_paid, df_bill, group_col, title, key, order=None):
        st.subheader(title)
        if df_paid.empty and df_bill.empty:
            st.info("Koi data nahi."); return
        s = make_summary(df_paid, df_bill, group_col, order=order)

        display = s[[group_col, "CA_Cr", "Total_Cr", "Arrear_Cr", "Efficiency_%",
                     "Txns", "Bills_Total", "Turnup_%"]].rename(columns={
            "CA_Cr": "CA (Cr.)", "Total_Cr": "Paid (Cr.)",
            "Arrear_Cr": "Total Outstanding (Cr.)", "Efficiency_%": "Eff %",
            "Txns": "Paid", "Bills_Total": "Bills", "Turnup_%": "Turn-up %",
        })
        display[group_col] = display[group_col].astype(str)

        tot_ca = round(s["CA_Cr"].sum(), 2); tot_paid = round(s["Total_Cr"].sum(), 2)
        tot_arr = round(s["Arrear_Cr"].sum(), 2)
        tot_txns = int(s["Txns"].sum()); tot_bills = int(s["Bills_Total"].sum())
        tot_eff = round((tot_paid / tot_ca * 100) if tot_ca else 0, 2)
        tot_turn = round((tot_txns / tot_bills * 100) if tot_bills else 0, 2)

        totals = pd.DataFrame([{
            group_col: "TOTAL", "CA (Cr.)": tot_ca, "Paid (Cr.)": tot_paid,
            "Total Outstanding (Cr.)": tot_arr, "Eff %": tot_eff,
            "Paid": tot_txns, "Bills": tot_bills, "Turn-up %": tot_turn,
        }])
        totals[group_col] = totals[group_col].astype(str)
        full = pd.concat([display, totals], ignore_index=True)

        col1, col2 = st.columns([1.7, 1])
        with col1:
            st.dataframe(
                full.style.format({
                    "CA (Cr.)": "₹ {:,.2f}", "Paid (Cr.)": "₹ {:,.2f}",
                    "Total Outstanding (Cr.)": "₹ {:,.2f}",
                    "Eff %": "{:.2f}%", "Paid": "{:,}", "Bills": "{:,}",
                    "Turn-up %": "{:.2f}%",
                }).apply(lambda r: ["font-weight: bold; background-color:#1a3a1a; color:white"] * len(r)
                         if r[group_col] == "TOTAL" else [""] * len(r), axis=1)
                  .background_gradient(subset=["Eff %"], cmap=EFFICIENCY_CMAP, vmin=0, vmax=100)
                  .background_gradient(subset=["Turn-up %"], cmap=TURNUP_CMAP, vmin=0, vmax=100),
                use_container_width=True, hide_index=True, height=400
            )
        with col2:
            if not s.empty:
                fig = px.bar(s, x=group_col, y="Efficiency_%", text_auto=".2f",
                             color="Efficiency_%", color_continuous_scale="RdYlGn",
                             range_color=(0, 100), height=400)
                fig.update_layout(xaxis_title="", yaxis_title="Efficiency %",
                                  coloraxis_showscale=False)
                apply_theme(fig, height=400)
                st.plotly_chart(fig, use_container_width=True, key=f"eff_{key}")
        return s

    def show_sdo_efficiency_table(df_paid, df_bill, key="eff_sdo"):
        st.subheader("SDO Name-wise")
        if df_paid.empty and df_bill.empty:
            st.info("Koi data nahi."); return

        if not df_bill.empty and CA_COL in df_bill.columns:
            ca = df_bill.groupby(SDO_NAME_COL, as_index=False, observed=True).agg(
                CA_Total=(CA_COL, "sum"), Bills_Total=(CA_COL, "size"),
                Arrear_Total=(ARREAR_COL, "sum") if ARREAR_COL in df_bill.columns
                            else (CA_COL, lambda x: 0))
        else:
            ca = pd.DataFrame({SDO_NAME_COL: [], "CA_Total": [], "Bills_Total": [], "Arrear_Total": []})

        if not df_paid.empty:
            paid = df_paid.groupby(SDO_NAME_COL, as_index=False, observed=True).agg(
                Total_Paid=(AMOUNT_COL, "sum"), Txns=(AMOUNT_COL, "size"))
        else:
            paid = pd.DataFrame({SDO_NAME_COL: [], "Total_Paid": [], "Txns": []})

        s = pd.merge(ca, paid, on=SDO_NAME_COL, how="outer").fillna(0)
        s["CA (Cr.)"]   = (s["CA_Total"] / CRORE).round(2)
        s["Paid (Cr.)"] = (s["Total_Paid"] / CRORE).round(2)
        s["Total Outstanding (Cr.)"] = (s["Arrear_Total"] / CRORE).round(2)
        s["Paid"]       = s["Txns"].astype(int)
        s["Bills"]      = s["Bills_Total"].astype(int)
        s["Eff %"]      = (s["Total_Paid"] / s["CA_Total"].replace(0, pd.NA) * 100).round(2).fillna(0)
        s["Turn-up %"]  = (s["Txns"] / s["Bills_Total"].replace(0, pd.NA) * 100).round(2).fillna(0)
        s = s.sort_values("Paid (Cr.)", ascending=False)

        display = s[[SDO_NAME_COL, "CA (Cr.)", "Paid (Cr.)", "Total Outstanding (Cr.)",
                     "Eff %", "Paid", "Bills", "Turn-up %"]].copy()
        display.columns = ["SDO Name", "CA (Cr.)", "Paid (Cr.)", "Total Outstanding (Cr.)",
                           "Eff %", "Paid", "Bills", "Turn-up %"]
        display["SDO Name"] = display["SDO Name"].astype(str)

        tot_ca = round(s["CA (Cr.)"].sum(), 2); tot_paid = round(s["Paid (Cr.)"].sum(), 2)
        tot_arr = round(s["Total Outstanding (Cr.)"].sum(), 2)
        tot_eff = round((tot_paid / tot_ca * 100) if tot_ca else 0, 2)
        tot_txns = int(s["Paid"].sum()); tot_bills = int(s["Bills"].sum())
        tot_turn = round((tot_txns / tot_bills * 100) if tot_bills else 0, 2)

        totals = pd.DataFrame([{
            "SDO Name": "TOTAL", "CA (Cr.)": tot_ca, "Paid (Cr.)": tot_paid,
            "Total Outstanding (Cr.)": tot_arr, "Eff %": tot_eff,
            "Paid": tot_txns, "Bills": tot_bills, "Turn-up %": tot_turn,
        }])
        full = pd.concat([display, totals], ignore_index=True)

        col1, col2 = st.columns([1.7, 1])
        with col1:
            st.dataframe(
                full.style.format({
                    "CA (Cr.)": "₹ {:,.2f}", "Paid (Cr.)": "₹ {:,.2f}",
                    "Total Outstanding (Cr.)": "₹ {:,.2f}",
                    "Eff %": "{:.2f}%", "Paid": "{:,}", "Bills": "{:,}",
                    "Turn-up %": "{:.2f}%",
                }).apply(lambda r: ["font-weight: bold; background-color:#1a3a1a; color:white"] * len(r)
                         if r["SDO Name"] == "TOTAL" else [""] * len(r), axis=1)
                  .background_gradient(subset=["Eff %"], cmap=EFFICIENCY_CMAP, vmin=0, vmax=100)
                  .background_gradient(subset=["Turn-up %"], cmap=TURNUP_CMAP, vmin=0, vmax=100),
                use_container_width=True, hide_index=True, height=400
            )
        with col2:
            if not s.empty:
                fig = px.bar(s, x=SDO_NAME_COL, y="Eff %", text_auto=".2f",
                             color="Eff %", color_continuous_scale="RdYlGn",
                             range_color=(0, 100), height=400)
                fig.update_layout(xaxis_title="", yaxis_title="Efficiency %",
                                  coloraxis_showscale=False)
                apply_theme(fig, height=400)
                st.plotly_chart(fig, use_container_width=True, key=f"eff_{key}")
        return s

    # ---------- LAZY EXCEL/PDF ----------
    def build_excel_lazy(df_paid, df_bill, year, month):
        import xlsxwriter
        output = io.BytesIO()
        wb = xlsxwriter.Workbook(output, {"in_memory": True, "nan_inf_to_errors": True})
        fmt_header = wb.add_format({"bold": True, "bg_color": "#1F4E78",
                                    "font_color": "white", "border": 1, "align": "center"})
        fmt_cr  = wb.add_format({"num_format": "₹ #,##0.00", "border": 1})
        fmt_int = wb.add_format({"num_format": "#,##0", "border": 1})
        fmt_pct = wb.add_format({"num_format": "0.00\"%\"", "border": 1})
        fmt_txt = wb.add_format({"border": 1})

        def add_sheet(name, df, cols):
            ws = wb.add_worksheet(name[:31])
            for i, c in enumerate(cols):
                ws.write(0, i, str(c), fmt_header)
            for r, (_, row) in enumerate(df.iterrows(), start=1):
                for c, col in enumerate(cols):
                    val = row.get(col, "")
                    if isinstance(val, (int, float)) and not pd.isna(val):
                        colname = str(col)
                        if "%" in colname or "Eff" in colname or "Turn-up" in colname or "Ach" in colname:
                            ws.write_number(r, c, round(float(val), 2), fmt_pct)
                        elif "Count" in colname or "Txns" in colname or "Bills" in colname:
                            ws.write_number(r, c, int(val), fmt_int)
                        else:
                            ws.write_number(r, c, round(float(val), 2), fmt_cr)
                    else:
                        ws.write(r, c, str(val), fmt_txt)
            for i, c in enumerate(cols):
                ws.set_column(i, i, max(14, len(str(c)) + 4))

        ws = wb.add_worksheet("Summary")
        total_paid = float(df_paid[AMOUNT_COL].sum()) if not df_paid.empty else 0
        total_ca   = float(df_bill[CA_COL].sum()) if not df_bill.empty else 0
        ws.write(0, 0, f"Payment Dashboard — {year}-{month:02d}", fmt_header)
        ws.write(1, 0, "Total Paid (Cr.)", fmt_header); ws.write(1, 1, round(total_paid/CRORE, 2), fmt_cr)
        ws.write(2, 0, "Current Assessment (Cr.)", fmt_header); ws.write(2, 1, round(total_ca/CRORE, 2), fmt_cr)

        s = make_summary(df_paid, df_bill, CATEGORY_COL)
        add_sheet("Tariff wise", s, list(s.columns))

        if SDO_CODE_COL in df_paid.columns or SDO_CODE_COL in df_bill.columns:
            s = make_summary(df_paid, df_bill, SDO_CODE_COL)
            s.insert(0, "SDO Name", s[SDO_CODE_COL].apply(sdo_short))
            add_sheet("SDO wise", s, list(s.columns))

        wb.close()
        output.seek(0)
        return output.read()

    def build_pdf_lazy(df_paid, df_bill, year, month):
        from reportlab.lib.pagesizes import A4
        from reportlab.lib import colors
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle

        buffer = io.BytesIO()
        page_size = A4
        doc = SimpleDocTemplate(buffer, pagesize=page_size,
                                leftMargin=28, rightMargin=28, topMargin=26, bottomMargin=38)
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle("T", parent=styles["Title"], fontSize=18,
                                     textColor=colors.HexColor("#1F4E78"), alignment=1)
        story = [Paragraph("<b>⚡ U.P. POWER CORPORATION LIMITED</b>", title_style)]
        story.append(Paragraph(f"Payment Dashboard — {year}-{month:02d}",
                               ParagraphStyle("s", parent=styles["Normal"], alignment=1)))

        total_paid = round(float(df_paid[AMOUNT_COL].sum())/CRORE, 2) if not df_paid.empty else 0
        total_ca   = round(float(df_bill[CA_COL].sum())/CRORE, 2) if not df_bill.empty else 0
        data = [["Total Paid", "Current Assessment"],
                [f"₹ {total_paid:,.2f} Cr.", f"₹ {total_ca:,.2f} Cr."]]
        t = Table(data, colWidths=[250, 250])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ]))
        story.append(Spacer(1, 15))
        story.append(t)
        doc.build(story)
        buffer.seek(0)
        return buffer.read()

    # ---------- DAILY CALENDAR ----------
    def tab_daily_calendar(df_paid, selected_month):
        st.subheader("📅 Daily Collection Calendar")
        if df_paid.empty:
            st.info("Koi paid data nahi mila."); return

        dfx = df_paid.copy()
        dfx["_DAY"] = dfx[DATE_COL].dt.day
        daily = dfx.groupby("_DAY", as_index=False).agg(
            Amount=(AMOUNT_COL, "sum"), Txns=(AMOUNT_COL, "size"))
        daily["Amount_Cr"] = (daily["Amount"] / CRORE).round(2)

        best_day = daily.loc[daily["Amount"].idxmax()] if not daily.empty else None
        worst_day = daily.loc[daily["Amount"].idxmin()] if not daily.empty else None
        avg_daily = round(daily["Amount"].mean() / CRORE, 2) if not daily.empty else 0

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Best Day", f"₹ {best_day['Amount_Cr']:,.2f} Cr." if best_day is not None else "—",
                  f"Day {int(best_day['_DAY'])}" if best_day is not None else "")
        c2.metric("Worst Day", f"₹ {worst_day['Amount_Cr']:,.2f} Cr." if worst_day is not None else "—",
                  f"Day {int(worst_day['_DAY'])}" if worst_day is not None else "")
        c3.metric("Avg / Day", f"₹ {avg_daily:,.2f} Cr.")
        c4.metric("Days with Payment", f"{len(daily):,}")

        st.divider()
        st.markdown("### 📊 Day-wise Collection")
        fig2 = px.bar(daily, x="_DAY", y="Amount_Cr", text_auto=".2f",
                      height=350, color="Amount_Cr", color_continuous_scale="Blues")
        fig2.update_layout(xaxis_title="Day of Month", yaxis_title="Amount (Cr.)",
                           coloraxis_showscale=False)
        apply_theme(fig2, height=350)
        st.plotly_chart(fig2, use_container_width=True, key="daily_bar")

        with st.expander("📋 Daily Table + CSV download"):
            show = daily[["_DAY", "Amount_Cr", "Txns"]].copy()
            show.columns = ["Day", "Amount (Cr.)", "Txns"]
            show["Day"] = show["Day"].astype(str)
            total_row = pd.DataFrame([{
                "Day": "TOTAL",
                "Amount (Cr.)": round(show["Amount (Cr.)"].sum(), 2),
                "Txns": int(show["Txns"].sum()),
            }])
            show_full = pd.concat([show, total_row], ignore_index=True)
            st.dataframe(
                show_full.style.format({"Amount (Cr.)": "₹ {:,.2f}", "Txns": "{:,}"})
                  .apply(lambda r: ["font-weight: bold; background-color:#1a3a1a; color:white"] * len(r)
                         if r["Day"] == "TOTAL" else [""] * len(r), axis=1),
                use_container_width=True, hide_index=True
            )
            csv = show_full.to_csv(index=False).encode("utf-8")
            st.download_button("⬇️ Daily CSV", csv, f"daily_{selected_month}.csv",
                               "text/csv", key="dl_daily")

    # ---------- TARGET ----------
    def tab_target(df_paid, df_bill):
        st.subheader("🎯 SDO Name — Target vs Achievement")
        if df_paid.empty and df_bill.empty:
            st.info("Data nahi mila."); return

        rows = []
        for grp in TARGET_GROUPS:
            sub_p = df_paid[df_paid[SDO_CODE_COL].isin(grp["codes"])] if not df_paid.empty else pd.DataFrame()
            sub_b = df_bill[df_bill[SDO_CODE_COL].isin(grp["codes"])] if not df_bill.empty else pd.DataFrame()
            actual = round(float(sub_p[AMOUNT_COL].sum()) / CRORE, 2) if not sub_p.empty else 0.0
            txns = len(sub_p); bills = len(sub_b)
            target = grp["target_cr"]
            ach = round((actual / target * 100) if target else 0, 2)
            turn = round((txns / bills * 100) if bills else 0, 2)
            rows.append({"SDO": grp["label"], "Actual_Cr": actual, "Target_Cr": target,
                         "Achievement_%": ach, "Txns": txns, "Bills": bills, "Turnup_%": turn})
        sdo_actual = pd.DataFrame(rows)

        total_actual = round(sdo_actual["Actual_Cr"].sum(), 2)
        total_target = TOTAL_TARGET_CR
        total_ach = round((total_actual / total_target * 100) if total_target else 0, 2)
        total_txns = int(sdo_actual["Txns"].sum())
        total_bills = int(sdo_actual["Bills"].sum())
        total_turn = round((total_txns / total_bills * 100) if total_bills else 0, 2)

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Total Actual", f"₹ {total_actual:,.2f} Cr.")
        c2.metric("Total Target", f"₹ {total_target:,.2f} Cr.")
        c3.metric("Achievement", f"{total_ach:.2f}%")
        c4.metric("Turn-up %", f"{total_turn:.2f}%")
        c5.metric("Transactions", f"{total_txns:,}")

        st.divider()
        st.markdown("### 📊 Target vs Actual")
        fig = go.Figure()
        fig.add_trace(go.Bar(x=sdo_actual["SDO"], y=sdo_actual["Actual_Cr"],
                             name="Actual", marker_color="#2E7D32",
                             text=sdo_actual["Actual_Cr"].round(2), textposition="auto"))
        fig.add_trace(go.Bar(x=sdo_actual["SDO"], y=sdo_actual["Target_Cr"],
                             name="Target", marker_color="#B0BEC5",
                             text=sdo_actual["Target_Cr"].round(2), textposition="auto"))
        fig.update_layout(barmode="group", height=400, xaxis_title="", yaxis_title="Amount (Cr.)")
        apply_theme(fig, height=400)
        st.plotly_chart(fig, use_container_width=True, key="target_chart")

        st.markdown("### 📋 Table")
        display = sdo_actual.rename(columns={
            "Actual_Cr": "Actual (Cr.)", "Target_Cr": "Target (Cr.)",
            "Achievement_%": "Ach %", "Txns": "Paid", "Bills": "Bills", "Turnup_%": "Turn-up %"})
        totals = pd.DataFrame([{
            "SDO": "TOTAL", "Actual (Cr.)": total_actual, "Target (Cr.)": total_target,
            "Ach %": total_ach, "Paid": total_txns, "Bills": total_bills, "Turn-up %": total_turn}])
        display["SDO"] = display["SDO"].astype(str); totals["SDO"] = totals["SDO"].astype(str)
        full = pd.concat([display, totals], ignore_index=True)
        st.dataframe(
            full.style.format({
                "Actual (Cr.)": "₹ {:,.2f}", "Target (Cr.)": "₹ {:,.2f}",
                "Ach %": "{:.2f}%", "Paid": "{:,}", "Bills": "{:,}", "Turn-up %": "{:.2f}%"
            }).apply(lambda r: ["font-weight: bold; background-color:#1a3a1a; color:white"] * len(r)
                     if r["SDO"] == "TOTAL" else [""] * len(r), axis=1)
              .background_gradient(subset=["Ach %"], cmap="RdYlGn", vmin=0, vmax=100)
              .background_gradient(subset=["Turn-up %"], cmap="RdYlGn", vmin=0, vmax=100),
            use_container_width=True, hide_index=True
        )

    # ---------- SUMMARY CARD ----------
    def tab_summary_card(df_paid, df_bill, selected_month):
        st.subheader("🖨️ Print-Friendly Summary Card")
        total_paid = float(df_paid[AMOUNT_COL].sum()) if not df_paid.empty else 0
        total_ca = float(df_bill[CA_COL].sum()) if not df_bill.empty else 0
        total_outstanding = float(df_bill[ARREAR_COL].sum()) if (not df_bill.empty and ARREAR_COL in df_bill.columns) else 0
        txns = len(df_paid); bills = len(df_bill)
        efficiency = round((total_paid / total_ca * 100) if total_ca else 0, 2)
        turnup = round((txns / bills * 100) if bills else 0, 2)

        st.markdown(f"""
        <div style="background: linear-gradient(135deg, #1F4E78 0%, #2E7D32 100%);
                    border-radius: 20px; padding: 30px; color: white; max-width: 900px; margin: 0 auto;">
            <h1 style="text-align: center; color: #F57C00; margin: 0;">⚡ UPPCL Payment Dashboard</h1>
            <p style="text-align: center; font-size: 16px;">📅 {selected_month}</p>
            <hr style="border: 0; border-top: 1px solid rgba(255,255,255,0.3); margin: 20px 0;">
            <table style="width: 100%; color: white; font-size: 16px;">
                <tr><td>📊 Current Assessment</td><td style="text-align:right;font-weight:800;">₹ {total_ca/CRORE:,.2f} Cr.</td></tr>
                <tr><td>✅ Total Paid</td><td style="text-align:right;font-weight:800;color:#4CAF50;">₹ {total_paid/CRORE:,.2f} Cr.</td></tr>
                <tr><td>📌 Total Outstanding</td><td style="text-align:right;font-weight:800;color:#FF9800;">₹ {total_outstanding/CRORE:,.2f} Cr.</td></tr>
                <tr><td>🎯 Efficiency</td><td style="text-align:right;font-weight:800;">{efficiency:.2f}%</td></tr>
                <tr><td>🔄 Turn-up %</td><td style="text-align:right;font-weight:800;">{turnup:.2f}%</td></tr>
                <tr><td>🔢 Transactions</td><td style="text-align:right;font-weight:800;">{txns:,} / {bills:,}</td></tr>
            </table>
        </div>
        """, unsafe_allow_html=True)

        share_text = (
            f"💰 UPPCL Payment Dashboard — {selected_month}\n\n"
            f"📊 Current Assessment: ₹ {total_ca/CRORE:,.2f} Cr.\n"
            f"✅ Total Paid: ₹ {total_paid/CRORE:,.2f} Cr.\n"
            f"📌 Outstanding: ₹ {total_outstanding/CRORE:,.2f} Cr.\n"
            f"🎯 Efficiency: {efficiency:.2f}%\n"
            f"🔄 Turn-up: {turnup:.2f}%\n"
        )
        st.download_button("📋 Summary Text (WhatsApp)", share_text.encode("utf-8"),
                           file_name=f"summary_{selected_month}.txt",
                           mime="text/plain", key="dl_share_text")

    # ---------- MAIN ----------
    st.markdown(f"""
    <div class="hero-header">
        <div style="display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 20px;">
            <div>
                <h1 class="hero-title">⚡ UPPCL PAYMENT DASHBOARD</h1>
                <p class="hero-subtitle">उत्तर प्रदेश पावर कॉर्पोरेशन लिमिटेड • U.P. Power Corporation Limited
                <span class="hero-badge">⚡ LIVE</span></p>
            </div>
            <div style="text-align: right;">
                <p style="color: #B0C4DE; margin: 0;">📅 {datetime.now().strftime('%d %B %Y')}</p>
                <p style="color: #F57C00; margin: 5px 0 0 0; font-weight: 700;">🕐 {datetime.now().strftime('%H:%M:%S')}</p>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    with st.sidebar:
        st.header("📂 File Upload")
        uploaded = st.file_uploader("File chunein (.xlsx / .xls / .csv)",
                                    type=["xlsx", "xls", "csv"], key="dash_upload")

    if uploaded is None:
        st.markdown("""
        <div style="text-align: center; padding: 60px 20px;">
            <div style="font-size: 80px; margin-bottom: 20px;">⚡</div>
            <h2 style="color: white; font-size: 32px;">Welcome to UPPCL Payment Dashboard</h2>
            <p style="color: #B0C4DE; font-size: 16px;">
                👈 <b style="color: #F57C00;">Left sidebar</b> se apni Excel/CSV file upload karein
            </p>
        </div>
        """, unsafe_allow_html=True)
        st.stop()

    t0 = datetime.now()
    file_bytes = uploaded.getvalue()
    with st.spinner("⚡ Loading..."):
        df_all = load_file(file_bytes, uploaded.name)

    if df_all.empty:
        st.error("File empty ya read nahi ho payi."); st.stop()

    load_ms = (datetime.now() - t0).total_seconds() * 1000
    st.success(f"✅ **{uploaded.name}** loaded — {len(df_all):,} rows in **{load_ms:.0f} ms**")

    required = [BILL_DATE_COL, DATE_COL, CATEGORY_COL, AMOUNT_COL, CA_COL]
    missing = [c for c in required if c not in df_all.columns]
    if missing:
        st.error(f"❌ Ye columns file me nahi mile: {missing}")
        st.write("**File me ye columns hain:**", list(df_all.columns))
        st.stop()

    merge_tuple = tuple(sorted(TARIFF_MERGE.items()))
    df_paid_all, df_bill_all = prepare(df_all, merge_tuple)

    if df_paid_all.empty and df_bill_all.empty:
        st.warning("Koi valid record nahi mila."); st.stop()

    st.sidebar.markdown("---")
    st.sidebar.markdown("## 🎛️ Filters")

    all_months = sorted(set(df_paid_all["_YM"].dropna().unique()) |
                        set(df_bill_all["_YM"].dropna().unique()), reverse=True)
    selected_month = (st.sidebar.selectbox("📅 Month", all_months, index=0, key="dash_month")
                      if all_months else datetime.today().strftime("%Y-%m"))

    filters = {}
    if SDO_CODE_COL in df_all.columns:
        sdo_codes = sorted(df_all[SDO_CODE_COL].dropna().astype(str).unique().tolist())
        sdo_options = {f"{sdo_short(c)} ({c})": c for c in sdo_codes}
        chosen_labels = st.sidebar.multiselect("🏢 SDO", options=list(sdo_options.keys()), key="dash_sdo")
        chosen_codes = [sdo_options[l] for l in chosen_labels]
        if chosen_codes:
            filters[SDO_CODE_COL] = chosen_codes

    if CATEGORY_COL in df_all.columns:
        tariff_list = sorted(df_all[CATEGORY_COL].dropna().unique().tolist())
        filters[CATEGORY_COL] = st.sidebar.multiselect("📋 Tariff Type", tariff_list, key="dash_tariff")

    if "PAYMENT_MODE" in df_all.columns:
        pm_list = sorted(df_all["PAYMENT_MODE"].dropna().unique().tolist())
        filters["PAYMENT_MODE"] = st.sidebar.multiselect("💳 Payment Mode", pm_list, key="dash_pm")

    if st.sidebar.button("🔄 Reset Filters", use_container_width=True, key="dash_reset"):
        st.rerun()

    df_paid = df_paid_all[df_paid_all["_YM"] == selected_month].copy()
    df_bill = df_bill_all[df_bill_all["_YM"] == selected_month].copy()

    for col, vals in filters.items():
        if vals:
            if col in df_paid.columns: df_paid = df_paid[df_paid[col].isin(vals)]
            if col in df_bill.columns: df_bill = df_bill[df_bill[col].isin(vals)]

    st.markdown(f"### 📅 Showing: **{selected_month}**")

    total_paid = float(df_paid[AMOUNT_COL].sum()) if not df_paid.empty else 0
    total_ca   = float(df_bill[CA_COL].sum()) if not df_bill.empty else 0
    total_outstanding = float(df_bill[ARREAR_COL].sum()) if (not df_bill.empty and ARREAR_COL in df_bill.columns) else 0
    txns = len(df_paid); bills = len(df_bill)
    efficiency = round((total_paid / total_ca * 100) if total_ca else 0, 2)
    turnup = round((txns / bills * 100) if bills else 0, 2)

    st.markdown(f"""
    <div class="kpi-grid">
        <div class="kpi-card"><span class="kpi-icon">📅</span><div class="kpi-label">Month</div><div class="kpi-value">{selected_month}</div></div>
        <div class="kpi-card"><span class="kpi-icon">📊</span><div class="kpi-label">Current Assessment</div><div class="kpi-value">₹ {total_ca/CRORE:,.2f}</div></div>
        <div class="kpi-card"><span class="kpi-icon">✅</span><div class="kpi-label">Total Paid</div><div class="kpi-value" style="color:#4CAF50;">₹ {total_paid/CRORE:,.2f}</div></div>
        <div class="kpi-card"><span class="kpi-icon">📌</span><div class="kpi-label">Outstanding</div><div class="kpi-value" style="color:#FF9800;">₹ {total_outstanding/CRORE:,.2f}</div></div>
        <div class="kpi-card"><span class="kpi-icon">🎯</span><div class="kpi-label">Efficiency</div><div class="kpi-value" style="color:#F57C00;">{efficiency:.2f}%</div></div>
        <div class="kpi-card"><span class="kpi-icon">🔄</span><div class="kpi-label">Turn-up %</div><div class="kpi-value">{turnup:.2f}%</div></div>
        <div class="kpi-card"><span class="kpi-icon">💰</span><div class="kpi-label">Transactions</div><div class="kpi-value">{txns:,}</div></div>
        <div class="kpi-card"><span class="kpi-icon">📄</span><div class="kpi-label">Bills</div><div class="kpi-value">{bills:,}</div></div>
    </div>
    """, unsafe_allow_html=True)

    # ---------- LAZY Downloads ----------
    st.markdown("### 📥 Downloads")
    dl1, dl2, _ = st.columns([1, 1, 3])
    with dl1:
        if st.button("📊 Generate Excel", use_container_width=True, key="btn_excel"):
            with st.spinner("Excel ban raha hai..."):
                try:
                    st.session_state["excel_bytes"] = build_excel_lazy(
                        df_paid, df_bill, int(selected_month[:4]), int(selected_month[5:7]))
                except Exception as e:
                    st.error(f"Excel error: {e}")
        if "excel_bytes" in st.session_state:
            st.download_button("⬇️ Download Excel", st.session_state["excel_bytes"],
                               file_name=f"dashboard_{selected_month}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               key="dl_excel_all")
    with dl2:
        if st.button("📄 Generate PDF", use_container_width=True, key="btn_pdf"):
            with st.spinner("PDF ban raha hai..."):
                try:
                    st.session_state["pdf_bytes"] = build_pdf_lazy(
                        df_paid, df_bill, int(selected_month[:4]), int(selected_month[5:7]))
                except Exception as e:
                    st.error(f"PDF error: {e}")
        if "pdf_bytes" in st.session_state:
            st.download_button("⬇️ Download PDF", st.session_state["pdf_bytes"],
                               file_name=f"dashboard_{selected_month}.pdf",
                               mime="application/pdf", key="dl_pdf_all")

    st.divider()

    tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
        "📋 Tariff", "⚡ Load", "🏢 SDO", "📈 Efficiency",
        "📅 Daily", "🎯 Target", "🏆 Consumers", "🖨️ Summary",
    ])

    with tab1:
        show_4col_table(df_paid, df_bill, CATEGORY_COL, "Tariff Type wise", key="tariff")
    with tab2:
        if not df_paid.empty and "LOAD_CATEGORY" in df_paid.columns:
            df_p_load = df_paid.dropna(subset=["LOAD_CATEGORY"])
            df_b_load = df_bill.dropna(subset=["LOAD_CATEGORY"]) if not df_bill.empty else df_bill
            show_4col_table(df_p_load, df_b_load, "LOAD_CATEGORY", "Load Category wise",
                            key="load", order=LOAD_ORDER)
        else:
            st.info("SANCTION_LOAD column nahi mila.")
    with tab3:
        if SDO_CODE_COL in df_all.columns:
            show_sdo_name_table(df_paid, df_bill, key="sdo_name")
        else:
            st.info("SDO_CODE column nahi mila.")
    with tab4:
        show_efficiency_table(df_paid, df_bill, CATEGORY_COL, "Tariff-wise", key="eff_tariff")
        if not df_paid.empty and "LOAD_CATEGORY" in df_paid.columns:
            df_p_load = df_paid.dropna(subset=["LOAD_CATEGORY"])
            df_b_load = df_bill.dropna(subset=["LOAD_CATEGORY"]) if not df_bill.empty else df_bill
            show_efficiency_table(df_p_load, df_b_load, "LOAD_CATEGORY", "Load-wise",
                                  key="eff_load", order=LOAD_ORDER)
        if SDO_CODE_COL in df_all.columns:
            show_sdo_efficiency_table(df_paid, df_bill, key="eff_sdo_name")
    with tab5:
        tab_daily_calendar(df_paid, selected_month)
    with tab6:
        tab_target(df_paid, df_bill)
    with tab7:
        if df_paid.empty or ACCT_COL not in df_paid.columns:
            st.info("ACCT_ID column nahi mila.")
        else:
            grp = [ACCT_COL, NAME_COL] if NAME_COL in df_paid.columns else [ACCT_COL]
            cons = df_paid.groupby(grp, as_index=False).agg(
                Total_Paid=(AMOUNT_COL, "sum"), Txns=(AMOUNT_COL, "size"))
            cons["Total_Cr"] = (cons["Total_Paid"] / CRORE).round(2)
            top_n = st.slider("Kitne dikhayein?", 5, 50, 10, key="topn")
            c1, c2 = st.columns(2)
            with c1:
                st.subheader(f"🏆 Top {top_n} Consumers")
                top = cons.sort_values("Total_Paid", ascending=False).head(top_n)
                top_show = top[grp + ["Total_Cr", "Txns"]].copy()
                top_show.columns = (["ACCT_ID", "NAME", "Total (Cr.)", "Txns"]
                                    if len(grp) == 2 else ["ACCT_ID", "Total (Cr.)", "Txns"])
                top_show["ACCT_ID"] = top_show["ACCT_ID"].astype(str)
                st.dataframe(top_show.style.format({"Total (Cr.)": "₹ {:,.2f}", "Txns": "{:,}"}),
                             use_container_width=True, hide_index=True, height=400)
            with c2:
                st.subheader(f"🔻 Bottom {top_n} Consumers")
                bot = cons[cons["Total_Paid"] > 0].sort_values("Total_Paid").head(top_n)
                bot_show = bot[grp + ["Total_Cr", "Txns"]].copy()
                bot_show.columns = (["ACCT_ID", "NAME", "Total (Cr.)", "Txns"]
                                    if len(grp) == 2 else ["ACCT_ID", "Total (Cr.)", "Txns"])
                bot_show["ACCT_ID"] = bot_show["ACCT_ID"].astype(str)
                st.dataframe(bot_show.style.format({"Total (Cr.)": "₹ {:,.2f}", "Txns": "{:,}"}),
                             use_container_width=True, hide_index=True, height=400)
    with tab8:
        tab_summary_card(df_paid, df_bill, selected_month)

    st.markdown(f"""
    <div class="uppcl-footer">
        <h4>⚡ U.P. POWER CORPORATION LIMITED</h4>
        <p>Shakti Bhawan, 14 Ashok Marg, Lucknow - 226001</p>
        <p>📞 0522-2286618 | 🌐 www.uppcl.org | ☎️ Toll Free: 1912</p>
    </div>
    """, unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════
#  🚀 ROUTER
# ═══════════════════════════════════════════════════════════
if page == "merger":
    render_merger()
elif page == "dashboard":
    render_dashboard()
