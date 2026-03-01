import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime, timedelta

st.set_page_config(page_title='Uncovered Orders Audit', page_icon='🚛', layout='wide')

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

_STOPWORDS = {'the','and','for','von','van','de','di','du','der','gmbh','bv','sa','ltd','llc','inc','co','kg','spa','sas','sl','nv','ag','plc','ug','bvba','srl','spzoo'}

def _normalise(s):
    if not isinstance(s, str): return ''
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
    t = _core_tokens(name)
    if len(t) < 2: return False
    for ct in _CST_TOKENS:
        if not ct: continue
        ov = len(t & ct)
        if ov >= 2 and ov / len(t) >= 0.8: return True
    return False

AMAZON_ALIAS_PATTERN = re.compile(r'^[a-z]{5,8}$')
FC_PATTERN = re.compile(r'^[A-Z]{3}\d{1,2}$')
REQUIRED_COLUMNS = ['Order ID','Source','Shipper','Destination Stop Date and Time','Destination Stop Facility Name','Creation Date and Time','Created by']
REQUIRED_COLUMNS_CST = ['Order ID','Source','Shipper','Destination Stop Date and Time','Creation Date and Time','Created by']

def is_fc_facility(name):
    if not isinstance(name, str): return False
    return bool(FC_PATTERN.match(name.strip()))

def classify_source(created_by):
    if not isinstance(created_by, str): return 'R4S'
    return 'SMC' if AMAZON_ALIAS_PATTERN.match(created_by.strip()) else 'R4S'

def load_smc_file(f):
    raw = f.read()
    if raw[:6] == b'Sheet0' or raw[:5] == b'Sheet':
        df = pd.read_csv(io.BytesIO(raw), sep='\t', dtype=str, skiprows=1)
        if 'Unnamed: 0' in df.columns: df = df.drop(columns=['Unnamed: 0'])
        return df
    return pd.read_excel(io.BytesIO(raw), dtype=str, engine='openpyxl')

def to_excel_bytes(sheets):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as w:
        for sn, df in sheets.items(): df.to_excel(w, sheet_name=sn, index=False)
    return buf.getvalue()

def reset_index_display(df):
    df = df.copy().reset_index(drop=True)
    df.index = df.index + 1
    return df

defaults = {
    'step': 1,
    'df_raw': None,
    'df_clean': None,
    'df_formatted': None,
    'df_step5': None,
    'cst_ext': None,
    'non_cst_ext': None,
    'portal_ids': [],
    'portal_input_mode': 'Paste IDs directly',
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

st.title('Uncovered Orders Audit')
st.caption('Amazon Freight Scheduling Team - Automated Audit Workflow')
st.divider()
step_labels = ['1. Load File','2. Cleanup','3. Classify','4. External Orders','5. Portal Check','6. Final Results']
pv = (st.session_state.step - 1) / (len(step_labels) - 1)
st.progress(pv, text='Step {} of {}: {}'.format(st.session_state.step, len(step_labels), step_labels[st.session_state.step-1]))
st.divider()

if st.session_state.step == 1:
    st.header('Step 1 - Upload SMC Export File')
    st.info('Download the uncovered orders file from SMC TMS (untick LTL, export), then upload it here to begin the audit.')
    uploaded = st.file_uploader('Upload your SMC uncovered orders export (.xlsx)', type=['xlsx','xls','csv'], key='smc_upload')
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
    for msg in log: st.markdown(msg)
    st.dataframe(reset_index_display(df), use_container_width=True)
    c1, c2 = st.columns(2)
    with c1:
        if st.button('Back to Step 1'): st.session_state.step = 1; st.rerun()
    with c2:
        if st.button('Proceed to Step 3 - Classify Orders', type='primary'):
            st.session_state.df_clean = df; st.session_state.step = 3; st.rerun()

elif st.session_state.step == 3:
    st.header('Step 3 - Format Report and Classify Orders')
    st.info('Renaming Column B to Source, keeping only the 7 required columns, and classifying each order as SMC or R4S.')
    df = st.session_state.df_clean.copy()
    cols = list(df.columns)
    if len(cols) >= 2:
        old = cols
        df = df.rename(columns={old: 'Source'})
        st.markdown('Renamed column {} to Source'.format(old))
    cm = {col.strip().lower(): col for col in df.columns}
    miss = [c for c in REQUIRED_COLUMNS if c.lower() not in cm]
    if miss: st.warning('Missing required columns: {}'.format(miss))
    ktk = [cm[c.lower()] for c in REQUIRED_COLUMNS if c.lower() in cm]
    df = df[ktk].copy()
    cbc = cm.get('created by')
    if cbc:
        df['Source'] = df[cbc].apply(classify_source)
        sc2 = (df['Source'] == 'SMC').sum()
        rc2 = (df['Source'] == 'R4S').sum()
        x1,x2,x3 = st.columns(3)
        x1.metric('Total Orders', len(df))
        x2.metric('SMC Orders', sc2)
        x3.metric('R4S Orders', rc2)
    st.dataframe(reset_index_display(df), use_container_width=True)
    c1, c2 = st.columns(2)
    with c1:
        if st.button('Back to Step 2'): st.session_state.step = 2; st.rerun()
    with c2:
        if st.button('Proceed to Step 4 - External Deliveries', type='primary'):
            st.session_state.df_formatted = df; st.session_state.step = 4; st.rerun()
elif st.session_state.step == 4:
    st.header('Step 4 - Process External Deliveries')
    df = st.session_state.df_formatted.copy()
    cm = {col.strip().lower(): col for col in df.columns}
    fc = cm.get('destination stop facility name')
    shc = cm.get('shipper')
    df['_is_fc'] = df[fc].apply(is_fc_facility)
    ext = df[df['_is_fc'] == False].copy()
    intr = df[df['_is_fc'] == True].copy()
    c1,c2 = st.columns(2)
    c1.metric('Internal (FC-bound) Orders', len(intr))
    c2.metric('External (non-FC) Orders', len(ext))
    if not ext.empty and shc:
        ext['_is_cst'] = ext[shc].apply(is_cst_shipper)
        cst_ext_raw = ext[ext['_is_cst'] == True].drop(columns=['_is_fc','_is_cst'])
        non_cst_ext = ext[ext['_is_cst'] == False].drop(columns=['_is_fc','_is_cst'])
        cst_cols_to_keep = [c for c in cst_ext_raw.columns if c.lower() != 'destination stop facility name']
        cst_ext = cst_ext_raw[cst_cols_to_keep].copy()
    else:
        cst_ext = pd.DataFrame(columns=REQUIRED_COLUMNS_CST)
        non_cst_ext = pd.DataFrame(columns=REQUIRED_COLUMNS)
    c1,c2 = st.columns(2)
    c1.metric('CST External Orders', len(cst_ext))
    c2.metric('Non-CST External Orders', len(non_cst_ext))
    st.divider()
    st.subheader('CST External Orders - copy to CST Task Sheet (Uncovered tab)')
    st.dataframe(reset_index_display(cst_ext), use_container_width=True)
    st.subheader('Non-CST External Orders - copy to AF Scheduling Daily Task Workbook (Uncovered tab)')
    st.dataframe(reset_index_display(non_cst_ext), use_container_width=True)
    st.divider()
    msg4 = 'ACTION REQUIRED' + chr(10) + chr(10)
    msg4 += '1. Copy CST External Orders above to the CST Task Sheet (Uncovered tab).' + chr(10)
    msg4 += '2. Copy Non-CST External Orders above to the AF Scheduling Daily Task Workbook (Uncovered tab).' + chr(10) + chr(10)
    msg4 += 'Click Done below once you have completed this step.'
    st.warning(msg4)
    c1,c2 = st.columns(2)
    with c1:
        if st.button('Back to Step 3'): st.session_state.step = 3; st.rerun()
    with c2:
        if st.button('Done - Proceed to Step 5', type='primary'):
            st.session_state.cst_ext = cst_ext
            st.session_state.non_cst_ext = non_cst_ext
            st.session_state.df_step5 = intr.drop(columns=['_is_fc'])
            st.session_state.portal_ids = []
            st.session_state.step = 5; st.rerun()

elif st.session_state.step == 5:
    st.header('Step 5 - Unified Portal ISA Check')
    ds5 = st.session_state.df_step5
    cm = {col.strip().lower(): col for col in ds5.columns}
    oic = cm.get('order id')
    rids = ds5[oic].dropna().astype(str).str.strip().tolist()
    st.info('{} Order IDs need to be checked in the Unified Portal.'.format(len(rids)))
    with st.expander('View / Copy Order IDs (copy-paste into Unified Portal)'):
        numbered_lines = []
        for i, oid in enumerate(rids):
            numbered_lines.append('{}. {}'.format(i + 1, oid))
        st.text(chr(10).join(numbered_lines))
    st.divider()
    st.subheader('Unified Portal Instructions')
    inst = '1. Copy the Order IDs from the panel above.' + chr(10)
    inst += '2. Go to the Unified Portal and set ID type to progressive number.' + chr(10)
    inst += '3. Paste the IDs in batches of 50 and click Submit.' + chr(10)
    inst += '4. Filter results where appointmentStatus = arrival scheduled.' + chr(10)
    inst += '5. Note the matching Order IDs.' + chr(10)
    inst += '6. Come back here and enter the results below.'
    st.markdown(inst)
    st.divider()
    st.subheader('Enter Unified Portal Results')
    im = st.radio('How would you like to provide the matching Order IDs?', ['Paste IDs directly','Upload a results file'], horizontal=True, key='portal_input_mode')
    if im == 'Paste IDs directly':
        pasted = st.text_area('Paste matching Order IDs here (one per line):', height=200, key='pasted_ids')
        if pasted.strip():
            ri = [l.strip() for l in pasted.strip().splitlines() if l.strip()]
            st.session_state.portal_ids = list(dict.fromkeys(ri))
            st.success('{} unique Order IDs entered.'.format(len(st.session_state.portal_ids)))
        else:
            st.session_state.portal_ids = []
    else:
        pf = st.file_uploader('Upload your Unified Portal results file (CSV or Excel, one Order ID per row, no header):', type=['csv','xlsx','xls'], key='portal_upload')
        if pf is not None:
            try:
                rb = pf.read()
                if pf.name.lower().endswith('.csv'):
                    pdf = pd.read_csv(io.BytesIO(rb), dtype=str, header=None)
                else:
                    pdf = pd.read_excel(io.BytesIO(rb), dtype=str, header=None, engine='openpyxl')
                ri = pdf.iloc[:,0].dropna().astype(str).str.strip().tolist()
                st.session_state.portal_ids = list(dict.fromkeys(ri))
                st.success('{} unique Order IDs loaded from file.'.format(len(st.session_state.portal_ids)))
            except Exception as e:
                st.error('Error reading file: {}'.format(e))
        else:
            st.session_state.portal_ids = []
    st.divider()
    c1,c2 = st.columns(2)
    with c1:
        if st.button('Back to Step 4'):
            st.session_state.portal_ids = []
            st.session_state.step = 4; st.rerun()
    with c2:
        if st.session_state.portal_ids:
            if st.button('Run Cross-Reference and Produce Final Results', type='primary'):
                ps = set(st.session_state.portal_ids)
                dfc = st.session_state.df_step5.copy()
                dfc['_in_portal'] = dfc[oic].astype(str).str.strip().isin(ps)
                mdf = dfc[dfc['_in_portal'] == True].drop(columns=['_in_portal'])
                uc = int((dfc['_in_portal'] == False).sum())
                shc2 = cm.get('shipper')
                if shc2 and not mdf.empty:
                    mdf = mdf.copy()
                    mdf['_is_cst'] = mdf[shc2].apply(is_cst_shipper)
                    cst_f = mdf[mdf['_is_cst'] == True].drop(columns=['_is_cst'])
                    non_cst_f = mdf[mdf['_is_cst'] == False].drop(columns=['_is_cst'])
                else:
                    cst_f = pd.DataFrame(columns=REQUIRED_COLUMNS)
                    non_cst_f = mdf if not mdf.empty else pd.DataFrame(columns=REQUIRED_COLUMNS)
                st.session_state['cst_final'] = cst_f
                st.session_state['non_cst_final'] = non_cst_f
                st.session_state['unmatched_count'] = uc
                st.session_state.step = 6; st.rerun()

elif st.session_state.step == 6:
    st.header('Audit Complete - Final Results')
    st.balloons()
    cf = st.session_state['cst_final']
    ncf = st.session_state['non_cst_final']
    uc = st.session_state['unmatched_count']
    c1,c2,c3 = st.columns(3)
    c1.metric('Matched (Arrival Scheduled)', len(cf) + len(ncf))
    c2.metric('CST Orders', len(cf))
    c3.metric('Non-CST Orders', len(ncf))
    if uc > 0:
        st.info('{} FC-bound order(s) were not found in the Unified Portal and have been excluded.'.format(uc))
    cst_out = cf if not cf.empty else pd.DataFrame(columns=REQUIRED_COLUMNS_CST)
    non_cst_out = ncf if not ncf.empty else pd.DataFrame(columns=REQUIRED_COLUMNS)
    fb = to_excel_bytes({'CST Orders': cst_out, 'Non-CST Orders': non_cst_out})
    st.download_button(label='Download Final_Results.xlsx', data=fb, file_name='Final_Results.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    st.divider()
    st.subheader('CST Orders - copy to CST Task Sheet (Uncovered tab)')
    st.dataframe(reset_index_display(cf), use_container_width=True)
    st.subheader('Non-CST Orders - copy to AF Scheduling Daily Task Workbook (Uncovered tab)')
    st.dataframe(reset_index_display(ncf), use_container_width=True)
    st.divider()
    msg6 = 'FINAL ACTION REQUIRED' + chr(10) + chr(10)
    msg6 += '1. Download the file above.' + chr(10)
    msg6 += '2. Copy CST Orders sheet to the CST Task Sheet (Uncovered tab).' + chr(10)
    msg6 += '3. Copy Non-CST Orders sheet to the AF Scheduling Daily Task Workbook (Uncovered tab).' + chr(10) + chr(10)
    msg6 += 'Audit complete!'
    st.warning(msg6)
    st.divider()
    if st.button('Start a New Audit', type='primary'):
        for k in list(st.session_state.keys()): del st.session_state[k]
        st.rerun()