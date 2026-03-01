import streamlit as st
import pandas as pd
import re
import io
import math

st.set_page_config(page_title='Uncovered Orders Audit', page_icon='\U0001f69b', layout='wide')

CST_SHIPPERS = [
    'Anheuser-Busch InBev Deutschland GmbH & Co KG', 'ARTSANA S.P.A.',
    'Beiersdorf Customer Supply GmbH', 'BDSK Handels GmbH & Co.KG',
    'Charles Kendall Freight', 'EST contracts B.V.', 'Coyote Logistics UK Ltd',
    'DANONE NUTRICIA SPA', 'Danone UK', 'DANONE IT SN', 'Danone UK SN',
    'DELONGHI APPLIANCES', 'Howard Tenens', 'DGL- Ingram (KSP) ES',
    'Eddie Stobart (Appleton Culina House- Culina Group) Unilever',
    'Great Bear Distribution CHORLEY', 'Great Bear Distribution Ltd (Spectrum)',
    'Great Bear Distribution MV1', 'Great Bear Distribution OLDHAM',
    'Great Bear MV1 - Helen of Troy', 'Great Bear Distribution SHEFFIELD',
    'Great Bear Port Salford', 'Great Bear Port Salford Mars',
    'H.J. Heinz BV', 'Hoover Ltd', 'iRobot UK Ltd',
    'JACOBS DOUWE EGBERTS GB LTD', 'JYSK SE',
    'Kellogg Marketing And Sales Company (UK) Limited',
    'Keter Italia S.p.A.', 'Keter Iberia S.L.U', 'Keter Germany Gmbh', 'Keter France Sas',
    'Mars PF France', 'Mars GmbH (CBT-DE)', 'Mars Multisales Spain S.L.',
    'Mars Food Europe CV', 'Mars GmbH FLOERSHEIM', 'Mars GmbH MINDEN',
    'Messaggerie Libri Spa', 'Moemax Logistik GmbH', 'Mondi Logistik GmbH',
    'Nestle Enterprises SA, Business Growth Solutions Division', 'Nestle UK',
    'Nestrade S.A. T-Hub Central', 'Nestrade SA (t-hub North)',
    'Nestrade S.A. T-Hub North', 'Nestrade S.A. T-Hub South',
    'Nestrade SA', 'Nestrade SA (Worms)', 'Nestrade T-hub West',
    'PepsiCo Deutschland GmbH', 'Philips Consumer Lifestyle BV', 'Pregis Ltd',
    'Procter & Gamble International Operations SA',
    'Robert Bosch Power Tools GmbH', 'Samsonite GmBH', 'saturn petcare gmbh',
    'Skechers EDC', 'SharkNinja Europe Ltd', 'SharkNinja Germany Gmbh',
    'SIG Combibloc GmbH - Linnich', 'SIG Combibloc GmbH - Wittenberg',
    'S.L. Systemlogistik GmbH', 'Soffass spa', 'Sofidel Germany GmbH',
    'Sofidel France SAS', 'Sofidel Spain', 'Sofidel UK',
    'Spectrum Brands (UK) Ltd', 'Tetra GmbH', 'TYRE ECO CHAIN',
    'The Book Service Limited', 'Unilever Europe B.V. (UK)',
    'Unilever Europe B.V. (DE)', 'Unilever Europe B.V. (EU)',
    'Versuni Nederland B.V', 'Walkers Snacks Distribution Ltd',
    'Wincanton (J SAINSBURY PLC)', 'Yankee Candle Co (Europe) LTD',
    'Yankee Candle Co - DE', 'Zalando SE', 'Zeitfracht Medien GmbH',
    'Danone Deutschland GmbH', 'Danone UK Waters', 'Danone FR', 'Tchibo GmbH',
    'Hachette UK Distribution', 'Wacker Chemie AG',
    'Electrolux Hausgeraete GmbH', 'Sharp Consumer Electronics Poland sp. z o.o.',
    'Brita France', 'BRITA SE - Shipments Beselich',
    'Brita Italia s.r.l. Unipersonale', 'home24 eLogistics GmbH & Co. KG',
    'Cargill Poland Sp. z o.o.', 'Nitto Advanced Film Gronau GmbH',
    'Cargill S.L.U.', 'Coca-Cola Europacific Partners Deutschland GmbH',
    'Schlaadt HighCut GmbH', 'Fressnapf Logistics Management GmbH',
    'tegut... gute Lebensmittel GmbH & Co. KG', 'Bio Springer',
    'La Palette Rouge Iberica s.a. succ.le in Italia',
    'COLGATE PALMOLIVE EUROPE', 'Hisense UK', 'ECOSCOOTING DELIVERY SL',
    'EDT BE SRL (TEMU)', 'Eddie Stobart (Appleton Culina House- Culina Group)',
    'LOreal Italy', 'Geodis D&E Normandie', 'Hager Electro SAS',
    'Heineken Deutschland GmbH', 'HarperCollins Publishers Ltd',
    'Euro Pool System UK Ltd', '3M EMEA GmbH', 'Falken Tyre Europe GmbH',
    'SACHSENMILCH Leppersdorf GmbH', 'XPO Transport Solutions UK Limited',
    'La Palette Rouge Iberica Sa', 'LPR - LA PALETTE ROUGE (GB) LTD',
    'BONDUELLE RE', 'HARIBO Sp. z o.o.', 'Groupe SEB WMF Consumer GmbH',
    'HOYER GmbH Internationale Fachspedition', 'BLACK & DECKER LIMITED BV',
    'Cycleon B.v.', 'Wella International Operations Switzerland Sarl',
]

_STOPWORDS = {
    'the','and','for','von','van','de','di','du','der','gmbh','bv','sa','ltd','llc','inc','co',
    'kg','spa','sas','sl','nv','ag','plc','ug','bvba','srl','spzoo'
}

def _normalise(s):
    if not isinstance(s, str):
        return ''
    s = re.sub(r'[^\w\s]', '', s)
    s = re.sub(r'\s+', ' ', s)
    return s.strip().lower()

def _core_tokens(s):
    words = re.sub(r'[^\w\s]', ' ', str(s)).lower().split()
    return frozenset(w for w in words if len(w) > 2 and w not in _STOPWORDS)

_CST_EXACT  = set(s.strip().lower() for s in CST_SHIPPERS)
_CST_FUZZY  = set(_normalise(s) for s in CST_SHIPPERS)
_CST_TOKENS = [_core_tokens(s) for s in CST_SHIPPERS]

def is_cst_shipper(name):
    if not isinstance(name, str) or not name.strip():
        return False
    if name.strip().lower() in _CST_EXACT:
        return True
    if _normalise(name) in _CST_FUZZY:
        return True
    t = _core_tokens(name)
    if len(t) < 2:
        return False
    for ct in _CST_TOKENS:
        if not ct:
            continue
        ov = len(t & ct)
        if ov >= 2 and ov / len(t) >= 0.8:
            return True
    return False

AMAZON_ALIAS_PATTERN = re.compile(r'^[a-z]{5,8}$')
FC_PATTERN = re.compile(r'^[A-Z]{3}\d{1,2}$')

REQUIRED_COLUMNS = [
    'Order ID','Source','Shipper','Destination Stop Date and Time',
    'Destination Stop Facility Name','Creation Date and Time','Created by'
]
REQUIRED_COLUMNS_CST = [
    'Order ID','Source','Shipper','Destination Stop Date and Time',
    'Creation Date and Time','Created by'
]

def is_fc_facility(name):
    if not isinstance(name, str):
        return False
    return bool(FC_PATTERN.match(name.strip()))

def classify_source(created_by):
    if not isinstance(created_by, str):
        return 'R4S'
    return 'SMC' if AMAZON_ALIAS_PATTERN.match(created_by.strip()) else 'R4S'

def load_smc_file(f):
    raw = f.read()
    if raw[:6] == b'Sheet0' or raw[:5] == b'Sheet':
        df = pd.read_csv(io.BytesIO(raw), sep='\t', dtype=str, skiprows=1)
        if 'Unnamed: 0' in df.columns:
            df = df.drop(columns=['Unnamed: 0'])
        return df
    return pd.read_excel(io.BytesIO(raw), dtype=str, engine='openpyxl')

def to_excel_bytes(sheets):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as w:
        for sn, df in sheets.items():
            df.to_excel(w, sheet_name=sn, index=False)
    return buf.getvalue()

def reset_index_display(df):
    df = df.copy().reset_index(drop=True)
    df.index = df.index + 1
    return df

def drop_if_exists(df, col_name):
    if df is None or df.empty:
        return df
    if col_name in df.columns:
        return df.drop(columns=[col_name])
    return df

def make_copy_block(df: pd.DataFrame, exclude_cols: list[str]) -> str:
    """
    Returns a TAB-separated text block with:
      - NO headers
      - NO index/serial
      - excluded columns removed
    Suitable for paste into Excel/Sheets.
    """
    if df is None or df.empty:
        return ""
    out = df.copy()
    for c in exclude_cols:
        if c in out.columns:
            out = out.drop(columns=[c])
    out = out.fillna("")
    lines = ["\t".join(map(str, row)) for row in out.to_numpy()]
    return "\n".join(lines)

def run_cross_reference():
    """
    Step 5 core computation: uses session_state.portal_ids + df_step5
    and advances to step 6.
    """
    ds5 = st.session_state.df_step5
    cm = {col.strip().lower(): col for col in ds5.columns}
    oic = cm.get('order id')
    if not oic:
        st.error("Missing 'Order ID' column for Step 5.")
        st.stop()

    portal_ids = st.session_state.portal_ids
    ps = set(portal_ids)

    dfc = ds5.copy()
    dfc['_in_portal'] = dfc[oic].astype(str).str.strip().isin(ps)

    matched_df = dfc[dfc['_in_portal'] == True].drop(columns=['_in_portal'])
    unmatched_count = int((dfc['_in_portal'] == False).sum())

    shc2 = cm.get('shipper')
    if shc2 and not matched_df.empty:
        matched_df = matched_df.copy()
        matched_df['_is_cst'] = matched_df[shc2].apply(is_cst_shipper)
        cst_f = matched_df[matched_df['_is_cst'] == True].drop(columns=['_is_cst'])
        non_cst_f = matched_df[matched_df['_is_cst'] == False].drop(columns=['_is_cst'])
    else:
        cst_f = pd.DataFrame(columns=REQUIRED_COLUMNS)
        non_cst_f = matched_df if not matched_df.empty else pd.DataFrame(columns=REQUIRED_COLUMNS)

    st.session_state.cst_final = cst_f
    st.session_state.non_cst_final = non_cst_f
    st.session_state.unmatched_count = unmatched_count
    st.session_state.step = 6
    st.rerun()

# ---- session state defaults ----
defaults = {
    'step': 1,
    'df_raw': None,
    'df_clean': None,
    'df_formatted': None,
    'df_step5': None,
    'cst_ext': None,
    'non_cst_ext': None,
    'portal_ids': [],
    'cst_final': None,
    'non_cst_final': None,
    'unmatched_count': 0,
    'step4_skipped': False,
    'batch_index': 0,
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ---- UI header / progress ----
st.title('Uncovered Orders Audit')
st.caption('Amazon Freight Scheduling Team - Automated Audit Workflow')
st.divider()

step_labels = ['1. Load File','2. Cleanup','3. Classify','4. External Orders','5. Portal Check','6. Final Results']
pv = (st.session_state.step - 1) / (len(step_labels) - 1)
st.progress(pv, text='Step {} of {}: {}'.format(st.session_state.step, len(step_labels), step_labels[st.session_state.step-1]))
st.divider()

# ---- Step 1 ----
if st.session_state.step == 1:
    st.header('Step 1 - Upload SMC Export File')
    st.info('Download the uncovered orders file from SMC TMS (untick LTL, intermodal) export, then upload it here to begin the audit.')

    uploaded = st.file_uploader(
        'Upload your SMC uncovered orders export (.xlsx, .xls, or .csv)',
        type=['xlsx','xls','csv'],
        key='smc_upload'
    )

    if uploaded is not None:
        try:
            df = load_smc_file(uploaded)
            st.session_state.df_raw = df
            st.success('File loaded: {} orders, {} columns detected.'.format(len(df), len(df.columns)))
            st.dataframe(df.head(10), use_container_width=True)
            st.caption('Showing first 10 of {} rows.'.format(len(df)))

            if st.button('Proceed to Step 2 - Data Cleanup', type='primary'):
                st.session_state.step = 2
                st.rerun()

        except Exception as e:
            st.error('Error reading file: {}. Please check the file and try again.'.format(e))

# ---- Step 2 ----
elif st.session_state.step == 2:
    st.header('Step 2 - Data Cleanup')
    st.info('Removing test orders (Shipper contains the word Test).')

    df = st.session_state.df_raw.copy()
    ic = len(df)
    log = []

    cm = {col.strip().lower(): col for col in df.columns}
    sc = cm.get('shipper')
    if sc:
        b = len(df)
        df = df[~df[sc].str.contains('test', case=False, na=False)]
        log.append('Removed {} test order(s).'.format(b - len(df)))

    log.append('Cleanup complete. {} removed. {} remaining.'.format(ic - len(df), len(df)))
    for msg in log:
        st.markdown(msg)

    st.dataframe(reset_index_display(df), use_container_width=True)

    c1, c2 = st.columns(2)
    with c1:
        if st.button('Back to Step 1'):
            st.session_state.step = 1
            st.rerun()
    with c2:
        if st.button('Proceed to Step 3 - Data Processing & Order Classification', type='primary'):
            st.session_state.df_clean = df
            st.session_state.step = 3
            st.rerun()

# ---- Step 3 ----
elif st.session_state.step == 3:
    st.header('Step 3 - Format Report and Classify Orders')
    st.info('Renaming Column B to Source, keeping only the 7 required columns, and classifying each order as SMC or R4S.')

    df = st.session_state.df_clean.copy()
    cols = list(df.columns)
    if len(cols) >= 2:
        old = cols[1]
        df = df.rename(columns={old: 'Source'})
        st.markdown('Renamed column {} to Source'.format(old))

    cm = {col.strip().lower(): col for col in df.columns}
    miss = [c for c in REQUIRED_COLUMNS if c.lower() not in cm]
    if miss:
        st.warning('Missing required columns: {}'.format(miss))

    keep_cols = [cm[c.lower()] for c in REQUIRED_COLUMNS if c.lower() in cm]
    df = df[keep_cols].copy()

    cbc = cm.get('created by')
    if cbc:
        df['Source'] = df[cbc].apply(classify_source)
        sc2 = (df['Source'] == 'SMC').sum()
        rc2 = (df['Source'] == 'R4S').sum()
        x1, x2, x3 = st.columns(3)
        x1.metric('Total Orders', len(df))
        x2.metric('SMC Orders', sc2)
        x3.metric('R4S Orders', rc2)

    st.dataframe(reset_index_display(df), use_container_width=True)

    c1, c2 = st.columns(2)
    with c1:
        if st.button('Back to Step 2'):
            st.session_state.step = 2
            st.rerun()
    with c2:
        if st.button('Proceed to Step 4 - External Deliveries', type='primary'):
            st.session_state.df_formatted = df
            st.session_state.step = 4
            st.rerun()

# ---- Step 4 ----
elif st.session_state.step == 4:
    st.header('Step 4 - Process External Deliveries')

    df = st.session_state.df_formatted.copy()
    cm = {col.strip().lower(): col for col in df.columns}
    fc = cm.get('destination stop facility name')
    shc = cm.get('shipper')

    if not fc or not shc:
        st.error("Missing required columns for Step 4. Ensure the export contains 'Destination Stop Facility Name' and 'Shipper'.")
        st.stop()

    df['_is_fc'] = df[fc].apply(is_fc_facility)
    ext = df[df['_is_fc'] == False].copy()
    intr = df[df['_is_fc'] == True].copy()

    # Auto-skip if no external orders
    if ext.empty:
        st.session_state.cst_ext = pd.DataFrame(columns=REQUIRED_COLUMNS_CST)
        st.session_state.non_cst_ext = pd.DataFrame(columns=REQUIRED_COLUMNS)
        st.session_state.df_step5 = intr.drop(columns=['_is_fc'])
        st.session_state.portal_ids = []
        st.session_state.batch_index = 0
        st.session_state.step4_skipped = True
        st.session_state.step = 5
        st.rerun()

    c1, c2 = st.columns(2)
    c1.metric('Internal (FC-bound) Orders', len(intr))
    c2.metric('External (non-FC) Orders', len(ext))

    ext['_is_cst'] = ext[shc].apply(is_cst_shipper)
    cst_ext_raw = ext[ext['_is_cst'] == True].drop(columns=['_is_fc', '_is_cst'])
    non_cst_ext = ext[ext['_is_cst'] == False].drop(columns=['_is_fc', '_is_cst'])

    # CST external orders remove Destination Stop Facility Name
    cst_cols_to_keep = [c for c in cst_ext_raw.columns if c.lower() != 'destination stop facility name']
    cst_ext = cst_ext_raw[cst_cols_to_keep].copy()

    c1, c2 = st.columns(2)
    c1.metric('CST External Orders', len(cst_ext))
    c2.metric('Non-CST External Orders', len(non_cst_ext))

    st.divider()
    st.subheader('CST External Orders - copy to CST Task Sheet (Uncovered tab)')
    st.dataframe(reset_index_display(cst_ext), use_container_width=True)

    st.subheader('Non-CST External Orders - copy to AF Scheduling Daily Task Workbook (Uncovered tab)')
    st.dataframe(reset_index_display(non_cst_ext), use_container_width=True)

    # Copy buttons (values only, no headers, exclude Created by)
    st.divider()
    st.subheader("Copy-ready blocks")
    cc1, cc2 = st.columns(2)
    with cc1:
        if st.button("Generate copy block: CST External Orders", key="copy_cst_ext"):
            st.session_state['_copy_block_cst_ext'] = make_copy_block(cst_ext, exclude_cols=['Created by'])
        blk = st.session_state.get('_copy_block_cst_ext', "")
        if blk:
            st.text_area("CST External copy block (Ctrl+A, Ctrl+C)", value=blk, height=220)
    with cc2:
        if st.button("Generate copy block: Non-CST External Orders", key="copy_non_cst_ext"):
            st.session_state['_copy_block_non_cst_ext'] = make_copy_block(non_cst_ext, exclude_cols=['Created by'])
        blk2 = st.session_state.get('_copy_block_non_cst_ext', "")
        if blk2:
            st.text_area("Non-CST External copy block (Ctrl+A, Ctrl+C)", value=blk2, height=220)

    st.divider()
    st.warning(
        "ACTION REQUIRED\n"
        "1. Copy CST External Orders to the CST Task Sheet (Uncovered tab).\n"
        "2. Copy Non-CST External Orders to the AF Scheduling Daily Task Workbook (Uncovered tab).\n"
        "3. Click Done below once you have completed this step."
    )

    c1, c2 = st.columns(2)
    with c1:
        if st.button('Back to Step 3'):
            st.session_state.step = 3
            st.rerun()
    with c2:
        if st.button('Done - Proceed to Step 5', type='primary'):
            st.session_state.cst_ext = cst_ext
            st.session_state.non_cst_ext = non_cst_ext
            st.session_state.df_step5 = intr.drop(columns=['_is_fc'])
            st.session_state.portal_ids = []
            st.session_state.batch_index = 0
            st.session_state.step4_skipped = False
            st.session_state.step = 5
            st.rerun()

# ---- Step 5 ----
elif st.session_state.step == 5:
    st.header('Step 5 - Unified Portal ISA Check')

    if st.session_state.step4_skipped:
        st.info("Step 4 was skipped automatically because **External Orders = 0**. Proceeding directly with FC-bound orders portal check.")

    ds5 = st.session_state.df_step5
    cm = {col.strip().lower(): col for col in ds5.columns}
    oic = cm.get('order id')
    if not oic:
        st.error("Missing 'Order ID' column for Step 5.")
        st.stop()

    rids = ds5[oic].dropna().astype(str).str.strip().tolist()
    st.info('{} Order IDs need to be checked in the Unified Portal.'.format(len(rids)))

    # Batch copy (50 per batch), NO SERIAL NUMBERS in copy block
    batch_size = 50
    total = len(rids)
    batch_count = max(1, math.ceil(total / batch_size))

    with st.expander('View / Copy Order IDs (copy-paste into Unified Portal)'):
        if total == 0:
            st.warning("No Order IDs available to copy.")
        else:
            st.caption(f"Copy/paste into Unified Portal in batches of {batch_size}. (IDs copy without serial numbers.)")

            btn_cols = st.columns(min(batch_count, 6))
            for i in range(min(batch_count, 6)):
                if btn_cols[i].button(f"Batch {i+1}", key=f"batch_btn_{i}"):
                    st.session_state.batch_index = i

            bi = int(st.session_state.batch_index or 0)
            bi = max(0, min(bi, batch_count - 1))
            st.session_state.batch_index = bi

            start = bi * batch_size
            end = min(start + batch_size, total)
            batch_ids = rids[start:end]

            st.text_area(
                f"Batch {bi+1} IDs ({start+1}–{end} of {total})",
                value="\n".join(batch_ids),
                height=260
            )

    st.divider()
    st.subheader('Unified Portal Instructions')
    st.markdown(
        "1. Copy a batch of Order IDs from the panel above.\n"
        "2. Go to the Unified Portal and set ID type to progressive number.\n"
        "3. Paste up to 50 IDs and click Submit.\n"
        "4. Filter results where appointmentStatus = arrival scheduled.\n"
        "6. Come back here and enter the results below."
    )

    st.divider()
    st.subheader('Enter Unified Portal Results')

    im = st.radio(
        'How would you like to provide the matching Order IDs?',
        ['Paste IDs directly', 'Upload a results file'],
        horizontal=True
    )

    if 'portal_ids' not in st.session_state:
        st.session_state.portal_ids = []

    if im == 'Paste Order IDs directly':
        pasted = st.text_area('Paste |arrival scheduled| Order IDs here', height=200)
        if pasted.strip():
            ri = [l.strip() for l in pasted.strip().splitlines() if l.strip()]
            st.session_state.portal_ids = list(dict.fromkeys(ri))
            st.success(f"{len(st.session_state.portal_ids)} unique Order IDs entered.")
        else:
            st.session_state.portal_ids = []

    else:
        pf = st.file_uploader(
            'Upload your Unified Portal results file (CSV or Excel, one Order ID per row, no header):',
            type=['csv', 'xlsx', 'xls'],
            key='portal_upload'
        )
        if pf is not None:
            try:
                rb = pf.read()
                if pf.name.lower().endswith('.csv'):
                    pdf = pd.read_csv(io.BytesIO(rb), dtype=str, header=None)
                else:
                    pdf = pd.read_excel(io.BytesIO(rb), dtype=str, header=None, engine='openpyxl')

                ri = pdf.iloc[:, 0].dropna().astype(str).str.strip().tolist()
                st.session_state.portal_ids = list(dict.fromkeys(ri))
                st.success(f"{len(st.session_state.portal_ids)} unique Order IDs loaded from file.")
            except Exception as e:
                st.session_state.portal_ids = []
                st.error(f"Error reading file: {e}")
        else:
            st.session_state.portal_ids = []

    st.divider()

    portal_ids = st.session_state.portal_ids
    current_count = len(portal_ids)

    c1, c2 = st.columns(2)
    with c1:
        if st.button('Back to Step 4'):
            st.session_state.step = 4
            st.rerun()

    with c2:
        run_clicked = st.button(
            'Run Cross-Reference and Produce Final Results',
            type='primary',
            disabled=(current_count == 0)
        )

    if run_clicked:
        run_cross_reference()

# ---- Step 6 ----
elif st.session_state.step == 6:
    st.header('Audit Complete - Final Results')
    st.balloons()

    cf = st.session_state.cst_final if st.session_state.cst_final is not None else pd.DataFrame(columns=REQUIRED_COLUMNS_CST)
    ncf = st.session_state.non_cst_final if st.session_state.non_cst_final is not None else pd.DataFrame(columns=REQUIRED_COLUMNS)
    uc = int(st.session_state.unmatched_count or 0)

    # CST FINAL results: remove Destination Stop Facility Name
    cf_clean = drop_if_exists(cf, 'Destination Stop Facility Name')

    c1, c2, c3 = st.columns(3)
    c1.metric('Matched (Arrival Scheduled)', len(cf) + len(ncf))
    c2.metric('CST Orders', len(cf))
    c3.metric('Non-CST Orders', len(ncf))

    if uc > 0:
        st.info('{} FC-bound order(s) were not found in the Unified Portal and have been excluded.'.format(uc))

    fb = to_excel_bytes({
        'CST Orders': cf_clean if not cf_clean.empty else pd.DataFrame(columns=[c for c in REQUIRED_COLUMNS_CST if c != 'Created by']),
        'Non-CST Orders': ncf if not ncf.empty else pd.DataFrame(columns=REQUIRED_COLUMNS)
    })

    st.download_button(
        label='Download Final_Results.xlsx',
        data=fb,
        file_name='Final_Results.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

    st.divider()
    st.subheader('CST Orders - copy to CST Task Sheet (Uncovered tab)')
    st.dataframe(reset_index_display(cf_clean), use_container_width=True)

    st.subheader('Non-CST Orders - copy to AF Scheduling Daily Task Workbook (Uncovered tab)')
    st.dataframe(reset_index_display(ncf), use_container_width=True)

    # Copy buttons (values only, no headers, exclude Created by)
    st.divider()
    st.subheader("Copy-ready blocks")
    cc1, cc2 = st.columns(2)
    with cc1:
        if st.button("Generate copy block: CST Final Orders", key="copy_cst_final"):
            st.session_state['_copy_block_cst_final'] = make_copy_block(cf_clean, exclude_cols=['Created by'])
        blk = st.session_state.get('_copy_block_cst_final', "")
        if blk:
            st.text_area("CST Final copy block (Ctrl+A, Ctrl+C)", value=blk, height=220)
    with cc2:
        if st.button("Generate copy block: Scheduling Final Orders", key="copy_non_cst_final"):
            st.session_state['_copy_block_non_cst_final'] = make_copy_block(ncf, exclude_cols=['Created by'])
        blk2 = st.session_state.get('_copy_block_non_cst_final', "")
        if blk2:
            st.text_area("Non-CST Final copy block (Ctrl+A, Ctrl+C)", value=blk2, height=220)

    st.divider()
    st.warning(
        "FINAL ACTION REQUIRED\n"
        "1. Download the file above.\n"
        "2. Copy CST Orders sheet to the CST Task Sheet (Uncovered tab).\n"
        "3. Copy Non-CST Orders sheet to the AF Scheduling Daily Task Workbook (Uncovered tab).\n"
        "Audit complete!"
    )

    st.divider()
    if st.button('Start a New Audit', type='primary'):
        keys_to_clear = list(defaults.keys()) + [
            'smc_upload', 'portal_upload',
            '_copy_block_cst_ext', '_copy_block_non_cst_ext',
            '_copy_block_cst_final', '_copy_block_non_cst_final'
        ]
        for k in keys_to_clear:
            if k in st.session_state:
                del st.session_state[k]

        for k, v in defaults.items():
            if k not in st.session_state:
                st.session_state[k] = v

        st.session_state.step = 1
        st.rerun()
