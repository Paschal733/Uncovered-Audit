
# =============================================================================
# UNCOVERED ORDERS AUDIT — STREAMLIT WEB APP
# Amazon Freight Scheduling Team
# =============================================================================
# HOW TO DEPLOY (Streamlit Community Cloud — free):
#   1. Create a free account at streamlit.io
#   2. Push this file to a GitHub repository (public or private)
#   3. Go to share.streamlit.io -> "New app" -> select your repo and this file
#   4. Click Deploy — your app will be live at a permanent URL
#
# HOW TO RUN LOCALLY (for testing):
#   1. pip install streamlit pandas openpyxl
#   2. streamlit run uncovered_audit_app.py
# =============================================================================

import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime, timedelta

# =============================================================================
# PAGE CONFIG
# =============================================================================

st.set_page_config(
    page_title="Uncovered Orders Audit",
    page_icon="🚛",
    layout="wide",
)

# =============================================================================
# CST SHIPPER LIST
# =============================================================================

CST_SHIPPERS = [
    "Anheuser-Busch InBev Deutschland GmbH & Co KG", "ARTSANA S.P.A.",
    "Beiersdorf Customer Supply GmbH", "BDSK Handels GmbH & Co.KG",
    "Charles Kendall Freight", "EST contracts B.V.", "Coyote Logistics UK Ltd",
    "DANONE NUTRICIA SPA", "Danone UK", "DANONE IT SN", "Danone UK SN",
    "DELONGHI APPLIANCES", "Howard Tenens", "DGL- Ingram (KSP) ES",
    "Eddie Stobart (Appleton Culina House- Culina Group) Unilever",
    "Great Bear Distribution CHORLEY", "Great Bear Distribution Ltd (Spectrum)",
    "Great Bear Distribution MV1", "Great Bear Distribution OLDHAM",
    "Great Bear MV1 - Helen of Troy", "Great Bear Distribution SHEFFIELD",
    "Great Bear Port Salford", "Great Bear Port Salford Mars",
    "H.J. Heinz BV", "Hoover Ltd", "iRobot UK Ltd",
    "JACOBS DOUWE EGBERTS GB LTD", "JYSK SE",
    "Kellogg Marketing And Sales Company (UK) Limited",
    "Keter Italia S.p.A.", "Keter Iberia S.L.U", "Keter Germany Gmbh", "Keter France Sas",
    "Mars PF France", "Mars GmbH (CBT-DE)", "Mars Multisales Spain S.L.",
    "Mars Food Europe CV", "Mars GmbH FLOERSHEIM", "Mars GmbH MINDEN",
    "Messaggerie Libri Spa", "Mömax Logistik GmbH", "Mondi Logistik GmbH",
    "Nestlé Enterprises SA, Business Growth Solutions Division", "Nestlé UK",
    "Nestrade S.A. T-Hub Central", "Nestrade SA (t-hub North)",
    "Nestrade S.A. T-Hub North", "Nestrade S.A. T-Hub South",
    "Nestrade SA", "Nestrade SA (Worms)", "Nestrade T-hub West",
    "PepsiCo Deutschland GmbH", "Philips Consumer Lifestyle BV", "Pregis Ltd",
    "Procter & Gamble International Operations SA",
    "Robert Bosch Power Tools GmbH", "Samsonite GmBH", "saturn petcare gmbh",
    "Skechers EDC", "SharkNinja Europe Ltd", "SharkNinja Germany Gmbh",
    "SIG Combibloc GmbH - Linnich", "SIG Combibloc GmbH - Wittenberg",
    "S.L. Systemlogistik GmbH", "Soffass spa", "Sofidel Germany GmbH",
    "Sofidel France SAS", "Sofidel Spain", "Sofidel UK",
    "Spectrum Brands (UK) Ltd", "Tetra GmbH", "TYRE ECO CHAIN",
    "The Book Service Limited", "Unilever Europe B.V. (UK)",
    "Unilever Europe B.V. (DE)", "Unilever Europe B.V. (EU)",
    "Versuni Nederland B.V", "Walkers Snacks Distribution Ltd",
    "Wincanton (J SAINSBURY PLC)", "Yankee Candle Co (Europe) LTD",
    "Yankee Candle Co - DE", "Zalando SE", "Zeitfracht Medien GmbH",
    "Danone Deutschland GmbH", "Danone UK Waters", "Danone FR", "Tchibo GmbH",
    "Hachette UK Distribution", "Wacker Chemie AG",
    "Electrolux Hausgeräte GmbH", "Sharp Consumer Electronics Poland sp. z o.o.",
    "Brita France", "BRITA SE - Shipments Beselich",
    "Brita Italia s.r.l. Unipersonale", "home24 eLogistics GmbH & Co. KG",
    "Cargill Poland Sp. z o.o.", "Nitto Advanced Film Gronau GmbH",
    "Cargill S.L.U.", "Coca-Cola Europacific Partners Deutschland GmbH",
    "Schlaadt HighCut GmbH", "Fressnapf Logistics Management GmbH",
    "tegut... gute Lebensmittel GmbH & Co. KG", "Bio Springer",
    "La Palette Rouge Iberica s.a. succ.le in Italia",
    "COLGATE PALMOLIVE EUROPE", "Hisense UK", "ECOSCOOTING DELIVERY SL",
    "EDT BE SRL (TEMU)", "Eddie Stobart (Appleton Culina House- Culina Group)",
    "L'Oreal Italy", "Geodis D&E Normandie", "Hager Electro SAS",
    "Heineken Deutschland GmbH", "HarperCollins Publishers Ltd",
    "Euro Pool System UK Ltd", "3M EMEA GmbH", "Falken Tyre Europe GmbH",
    "SACHSENMILCH Leppersdorf GmbH", "XPO Transport Solutions UK Limited",
    "La Palette Rouge Iberica Sa", "LPR - LA PALETTE ROUGE (GB) LTD",
    "BONDUELLE RE", "HARIBO Sp. z o.o.", "Groupe SEB WMF Consumer GmbH",
    "HOYER GmbH Internationale Fachspedition", "BLACK & DECKER LIMITED BV",
    "Cycleon B.v.", "Wella International Operations Switzerland Sarl",
]

# =============================================================================
# MATCHING LOGIC
# =============================================================================

_STOPWORDS = {
    'the', 'and', 'for', 'von', 'van', 'de', 'di', 'du', 'der',
    'gmbh', 'bv', 'sa', 'ltd', 'llc', 'inc', 'co', 'kg', 'spa',
    'sas', 'sl', 'nv', 'ag', 'plc', 'ug', 'bvba', 'srl', 'spzoo',
}

def _normalise(s):
    if not isinstance(s, str): return ""
    s = re.sub(r'[^\w\s]', '', s)
    s = re.sub(r'\s+', ' ', s)
    return s.strip().lower()

def _core_tokens(s):
    words = re.sub(r'[^\w\s]', ' ', s).lower().split()
    return frozenset(w for w in words if len(w) > 2 and w not in _STOPWORDS)

_CST_EXACT  = set(s.strip().lower() for s in CST_SHIPPERS)
_CST_FUZZY  = set(_normalise(s) for s in CST_SHIPPERS)
_CST_TOKENS = [_core_tokens(s) for s in CST_SHIPPERS]

def is_cst_shipper(name):
    if not isinstance(name, str) or not name.strip(): return False
    if name.strip().lower() in _CST_EXACT: return True
    if _normalise(name) in _CST_FUZZY: return True
    input_tokens = _core_tokens(name)
    if len(input_tokens) < 2: return False
    for cst_tokens in _CST_TOKENS:
        if not cst_tokens: continue
        overlap = len(input_tokens & cst_tokens)
        if overlap >= 2 and overlap / len(input_tokens) >= 0.8: return True
    return False

AMAZON_ALIAS_PATTERN = re.compile(r'^[a-z]{5,8}$')
FC_PATTERN = re.compile(r'^[A-Z]{3}\d{1,2}$')

REQUIRED_COLUMNS = [
    "Order ID", "Source", "Shipper",
    "Destination Stop Date and Time", "Destination Stop Facility Name",
    "Creation Date and Time", "Created by",
]

def is_fc_facility(name):
    if not isinstance(name, str): return False
    return bool(FC_PATTERN.match(name.strip()))

def classify_source(created_by):
    if not isinstance(created_by, str): return "R4S"
    return "SMC" if AMAZON_ALIAS_PATTERN.match(created_by.strip()) else "R4S"

def load_smc_file(uploaded_file):
    raw = uploaded_file.read()
    if raw[:6] == b'Sheet0' or raw[:5] == b'Sheet':
        df = pd.read_csv(io.BytesIO(raw), sep='\t', dtype=str, skiprows=1)
        if 'Unnamed: 0' in df.columns:
            df = df.drop(columns=['Unnamed: 0'])
        return df
    return pd.read_excel(io.BytesIO(raw), dtype=str, engine='openpyxl')

def to_excel_bytes(sheets):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    return buf.getvalue()

# =============================================================================
# SESSION STATE INITIALISATION
# =============================================================================

defaults = {
    "step": 1,
    "df_raw": None,
    "df_clean": None,
    "df_formatted": None,
    "df_step5": None,
    "cst_ext": None,
    "non_cst_ext": None,
}
for key, val in defaults.items():
    if key not in st.session_state:
        st.session_state[key] = val

# =============================================================================
# APP HEADER
# =============================================================================

st.title("Uncovered Orders Audit")
st.caption("Amazon Freight Scheduling Team — Automated Audit Workflow")
st.divider()

step_labels = [
    "1. Load File", "2. Cleanup", "3. Classify",
    "4. External Orders", "5. Portal Check", "6. Final Results"
]
progress_val = (st.session_state.step - 1) / (len(step_labels) - 1)
st.progress(progress_val, text="Step {} of {}: {}".format(
    st.session_state.step, len(step_labels), step_labels[st.session_state.step - 1]))
st.divider()

# =============================================================================
# STEP 1 — LOAD SMC EXPORT FILE
# =============================================================================

if st.session_state.step == 1:
    st.header("Step 1 — Upload SMC Export File")
    st.info("Download the uncovered orders file from SMC TMS (untick LTL, export), then upload it here to begin the audit.")

    uploaded = st.file_uploader(
        "Upload your SMC uncovered orders export (.xlsx)",
        type=["xlsx", "xls", "csv"],
        key="smc_upload",
    )

    if uploaded is not None:
        try:
            df = load_smc_file(uploaded)
            st.session_state.df_raw = df
            st.success("File loaded: {} orders, {} columns detected.".format(len(df), len(df.columns)))
            st.dataframe(df.head(10), use_container_width=True)
            st.caption("Showing first 10 of {} rows.".format(len(df)))

            if st.button("Proceed to Step 2 — Data Cleanup", type="primary"):
                st.session_state.step = 2
                st.rerun()

        except Exception as e:
            st.error("Error reading file: {}. Please check the file and try again.".format(e))

# =============================================================================
# STEP 2 — DATA CLEANUP
# =============================================================================

elif st.session_state.step == 2:
    st.header("Step 2 — Data Cleanup")
    st.info("Removing: test orders (Shipper contains 'Test'), orders outside the current year, and orders created more than 2 months ago.")

    df = st.session_state.df_raw.copy()
    initial_count = len(df)
    log = []

    col_map      = {col.strip().lower(): col for col in df.columns}
    shipper_col  = col_map.get("shipper")
    creation_col = col_map.get("creation date and time")

    if shipper_col:
        before = len(df)
        df = df[~df[shipper_col].str.contains("test", case=False, na=False)]
        removed = before - len(df)
        log.append("Removed {} test order(s).".format(removed))

    if creation_col:
        df[creation_col] = pd.to_datetime(df[creation_col], errors='coerce', dayfirst=False)
        current_year   = datetime.now().year
        two_months_ago = datetime.now() - timedelta(days=60)
        before = len(df)
        df = df[
            (df[creation_col].dt.year == current_year) &
            (df[creation_col] >= two_months_ago)
        ]
        removed = before - len(df)
        log.append("Removed {} order(s) outside current year or older than 2 months.".format(removed))

    log.append("Cleanup complete. {} orders removed. {} orders remaining.".format(
        initial_count - len(df), len(df)))

    for msg in log:
        st.markdown(msg)

    st.dataframe(df, use_container_width=True)

    col1, col2 = st.columns(2)
    with col1:
        if st.button("Back to Step 1"):
            st.session_state.step = 1
            st.rerun()
    with col2:
        if st.button("Proceed to Step 3 — Classify Orders", type="primary"):
            st.session_state.df_clean = df
            st.session_state.step = 3
            st.rerun()

# =============================================================================
# STEP 3 — FORMAT AND CLASSIFY
# =============================================================================

elif st.session_state.step == 3:
    st.header("Step 3 — Format Report & Classify Orders")
    st.info("Renaming Column B to 'Source', keeping only the 7 required columns, and classifying each order as SMC (Amazon alias) or R4S (all others).")

    df = st.session_state.df_clean.copy()

    columns = list(df.columns)
    if len(columns) >= 2:
        old_name = columns[1]
        df = df.rename(columns={old_name: "Source"})
        st.markdown("Renamed column '{}' to 'Source'".format(old_name))

    col_map = {col.strip().lower(): col for col in df.columns}

    missing = [c for c in REQUIRED_COLUMNS if c.lower() not in col_map]
    if missing:
        st.warning("Missing required columns: {}".format(missing))

    cols_to_keep = [col_map[c.lower()] for c in REQUIRED_COLUMNS if c.lower() in col_map]
    df = df[cols_to_keep].copy()

    created_by_col = col_map.get("created by")
    if created_by_col:
        df["Source"] = df[created_by_col].apply(classify_source)
        smc_count = (df["Source"] == "SMC").sum()
        r4s_count = (df["Source"] == "R4S").sum()
        c1, c2, c3 = st.columns(3)
        c1.metric("Total Orders", len(df))
        c2.metric("SMC Orders", smc_count)
        c3.metric("R4S Orders", r4s_count)

    st.dataframe(df, use_container_width=True)

    col1, col2 = st.columns(2)
    with col1:
        if st.button("Back to Step 2"):
            st.session_state.step = 2
            st.rerun()
    with col2:
        if st.button("Proceed to Step 4 — External Deliveries", type="primary"):
            st.session_state.df_formatted = df
            st.session_state.step = 4
            st.rerun()

# =============================================================================
# STEP 4 — EXTERNAL DELIVERIES
# =============================================================================

elif st.session_state.step == 4:
    st.header("Step 4 — Process External Deliveries")

    df = st.session_state.df_formatted.copy()
    col_map      = {col.strip().lower(): col for col in df.columns}
    facility_col = col_map.get("destination stop facility name")
    shipper_col  = col_map.get("shipper")

    df["_is_fc"]  = df[facility_col].apply(is_fc_facility)
    external_df   = df[df["_is_fc"] == False].copy()
    internal_df   = df[df["_is_fc"] == True].copy()

    c1, c2 = st.columns(2)
    c1.metric("Internal (FC-bound) Orders", len(internal_df))
    c2.metric("External (non-FC) Orders", len(external_df))

    if not external_df.empty and shipper_col:
        external_df["_is_cst"] = external_df[shipper_col].apply(is_cst_shipper)
        cst_ext     = external_df[external_df["_is_cst"] == True].drop(columns=["_is_fc", "_is_cst"])
        non_cst_ext = external_df[external_df["_is_cst"] == False].drop(columns=["_is_fc", "_is_cst"])
    else:
        cst_ext     = pd.DataFrame(columns=REQUIRED_COLUMNS)
        non_cst_ext = pd.DataFrame(columns=REQUIRED_COLUMNS)

    c1, c2 = st.columns(2)
    c1.metric("CST External Orders", len(cst_ext))
    c2.metric("Non-CST External Orders", len(non_cst_ext))

    excel_bytes = to_excel_bytes({
        "CST External Orders": cst_ext if not cst_ext.empty else pd.DataFrame(columns=REQUIRED_COLUMNS),
        "Non-CST External Orders": non_cst_ext if not non_cst_ext.empty else pd.DataFrame(columns=REQUIRED_COLUMNS),
    })

    st.divider()
    st.subheader("Download External Orders File")
    st.download_button(
        label="Download Step4_External_Orders.xlsx",
        data=excel_bytes,
        file_name="Step4_External_Orders.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.divider()
    st.subheader("CST External Orders — copy to CST Task Sheet (Uncovered tab)")
    st.dataframe(cst_ext, use_container_width=True)

    st.subheader("Non-CST External Orders — copy to AF Scheduling Daily Task Workbook (Uncovered tab)")
    st.dataframe(non_cst_ext, use_container_width=True)

    st.divider()
    st.warning(
        "ACTION REQUIRED

"
        "1. Download the file above.
"
        "2. Copy 'CST External Orders' sheet to the CST Task Sheet (Uncovered tab).
"
        "3. Copy 'Non-CST External Orders' sheet to the AF Scheduling Daily Task Workbook (Uncovered tab).

"
        "Click 'Done' below once you have completed this step."
    )

    col1, col2 = st.columns(2)
    with col1:
        if st.button("Back to Step 3"):
            st.session_state.step = 3
            st.rerun()
    with col2:
        if st.button("Done — Proceed to Step 5", type="primary"):
            st.session_state.cst_ext     = cst_ext
            st.session_state.non_cst_ext = non_cst_ext
            st.session_state.df_step5    = internal_df.drop(columns=["_is_fc"])
            st.session_state.step        = 5
            st.rerun()

# =============================================================================
# STEP 5 — UNIFIED PORTAL ISA CHECK
# =============================================================================

elif st.session_state.step == 5:
    st.header("Step 5 — Unified Portal ISA Check")

    df_step5     = st.session_state.df_step5
    col_map      = {col.strip().lower(): col for col in df_step5.columns}
    order_id_col = col_map.get("order id")

    remaining_ids = df_step5[order_id_col].dropna().astype(str).str.strip().tolist()
    st.info("{} Order IDs need to be checked in the Unified Portal.".format(len(remaining_ids)))

    ids_excel = to_excel_bytes({"Order IDs": pd.DataFrame({"Order ID": remaining_ids})})
    st.download_button(
        label="Download Order IDs for Unified Portal",
        data=ids_excel,
        file_name="Step5_Order_IDs_for_Portal.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    with st.expander("View / Copy Order IDs (copy-paste into Unified Portal)"):
        st.code("
".join(remaining_ids), language=None)

    st.divider()
    st.subheader("Unified Portal Instructions")
    st.markdown(
        "1. Download the file above (or copy the IDs from the panel above).
"
        "2. Go to the Unified Portal and set ID type to **progressive number**.
"
        "3. Paste the IDs in batches of 50 and click **Submit**.
"
        "4. Filter results where **appointmentStatus = arrival scheduled**.
"
        "5. Note the matching Order IDs.
"
        "6. Come back here and enter the results below."
    )

    st.divider()
    st.subheader("Enter Unified Portal Results")

    input_method = st.radio(
        "How would you like to provide the matching Order IDs?",
        ["Paste IDs directly", "Upload a results file"],
        horizontal=True,
    )

    portal_ids = []

    if input_method == "Paste IDs directly":
        pasted = st.text_area(
            "Paste matching Order IDs here (one per line):",
            height=200,
            placeholder="7134408318
3992975874
7724915204
...",
        )
        if pasted.strip():
            raw_ids = [line.strip() for line in pasted.strip().splitlines() if line.strip()]
            portal_ids = list(dict.fromkeys(raw_ids))
            st.success("{} unique Order IDs entered.".format(len(portal_ids)))

    else:
        portal_file = st.file_uploader(
            "Upload your Unified Portal results file (CSV or Excel, one Order ID per row, no header):",
            type=["csv", "xlsx", "xls"],
            key="portal_upload",
        )
        if portal_file is not None:
            try:
                raw_bytes = portal_file.read()
                if portal_file.name.lower().endswith(".csv"):
                    portal_df = pd.read_csv(io.BytesIO(raw_bytes), dtype=str, header=None)
                else:
                    portal_df = pd.read_excel(io.BytesIO(raw_bytes), dtype=str, header=None, engine='openpyxl')
                raw_ids = portal_df.iloc[:, 0].dropna().astype(str).str.strip().tolist()
                portal_ids = list(dict.fromkeys(raw_ids))
                st.success("{} unique Order IDs loaded from file.".format(len(portal_ids)))
            except Exception as e:
                st.error("Error reading file: {}".format(e))

    st.divider()
    col1, col2 = st.columns(2)
    with col1:
        if st.button("Back to Step 4"):
            st.session_state.step = 4
            st.rerun()
    with col2:
        if portal_ids and st.button("Run Cross-Reference and Produce Final Results", type="primary"):
            portal_ids_set = set(portal_ids)
            df_step5 = st.session_state.df_step5.copy()
            df_step5["_in_portal"] = df_step5[order_id_col].astype(str).str.strip().isin(portal_ids_set)
            matched_df      = df_step5[df_step5["_in_portal"] == True].drop(columns=["_in_portal"])
            unmatched_count = int((df_step5["_in_portal"] == False).sum())

            shipper_col = col_map.get("shipper")
            if shipper_col and not matched_df.empty:
                matched_df = matched_df.copy()
                matched_df["_is_cst"] = matched_df[shipper_col].apply(is_cst_shipper)
                cst_final     = matched_df[matched_df["_is_cst"] == True].drop(columns=["_is_cst"])
                non_cst_final = matched_df[matched_df["_is_cst"] == False].drop(columns=["_is_cst"])
            else:
                cst_final     = pd.DataFrame(columns=REQUIRED_COLUMNS)
                non_cst_final = matched_df if not matched_df.empty else pd.DataFrame(columns=REQUIRED_COLUMNS)

            st.session_state["cst_final"]       = cst_final
            st.session_state["non_cst_final"]   = non_cst_final
            st.session_state["unmatched_count"] = unmatched_count
            st.session_state.step = 6
            st.rerun()

# =============================================================================
# STEP 6 — FINAL RESULTS
# =============================================================================

elif st.session_state.step == 6:
    st.header("Audit Complete — Final Results")
    st.balloons()

    cst_final       = st.session_state["cst_final"]
    non_cst_final   = st.session_state["non_cst_final"]
    unmatched_count = st.session_state["unmatched_count"]

    c1, c2, c3 = st.columns(3)
    c1.metric("Matched (Arrival Scheduled)", len(cst_final) + len(non_cst_final))
    c2.metric("CST Orders", len(cst_final))
    c3.metric("Non-CST Orders", len(non_cst_final))

    if unmatched_count > 0:
        st.info("{} FC-bound order(s) were not found in the Unified Portal and have been excluded.".format(unmatched_count))

    final_excel = to_excel_bytes({
        "CST Orders": cst_final if not cst_final.empty else pd.DataFrame(columns=REQUIRED_COLUMNS),
        "Non-CST Orders": non_cst_final if not non_cst_final.empty else pd.DataFrame(columns=REQUIRED_COLUMNS),
    })

    st.download_button(
        label="Download Step5_Final_Results.xlsx",
        data=final_excel,
        file_name="Step5_Final_Results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.divider()
    st.subheader("CST Orders — copy to CST Task Sheet (Uncovered tab)")
    st.dataframe(cst_final, use_container_width=True)

    st.subheader("Non-CST Orders — copy to AF Scheduling Daily Task Workbook (Uncovered tab)")
    st.dataframe(non_cst_final, use_container_width=True)

    st.divider()
    st.warning(
        "FINAL ACTION REQUIRED

"
        "1. Download the file above.
"
        "2. Copy 'CST Orders' sheet to the CST Task Sheet (Uncovered tab).
"
        "3. Copy 'Non-CST Orders' sheet to the AF Scheduling Daily Task Workbook (Uncovered tab).

"
        "Audit complete!"
    )

    st.divider()
    if st.button("Start a New Audit", type="primary"):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()

