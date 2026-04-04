import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import re
import io
import math
import unicodedata
import html
import uuid

st.set_page_config(page_title='Uncovered Audit Automation Tool', page_icon='\U0001f69b', layout='wide')

UNIFIED_PORTAL_URL = "https://unified-portal-eu.corp.amazon.com/#/appointment?searchType=PRO&searchCategory=appointment&searchIds=&searchIds=9524366171"

CST_SHIPPERS = [
    'Amazon Business','AEG Electrolux Hausgeräte GmbH',
    'Anheuser-Busch InBev Deutschland GmbH & Co KG',
    'ARTSANA S.P.A.',
    'Brita France',
    'BRITA SE - Shipments Beselich',
    'Brita Italia s.r.l. Unipersonale',
    'Coyote Logistics UK Ltd',
    'Danone UK',
    'DANONE IT SN',
    'Danone UK SN',
    'DELONGHI APPLIANCES',
    'Electrolux Hausgeräte GmbH',
    'Fressnapf Logistics Management GmbH',
    'Hachette UK Distribution',
    'HarperCollins Publishers Ltd',
    'Hisense UK',
    'Howard Tenens',
    'Eddie Stobart (Appleton Culina House- Culina Group)',
    'Eddie Stobart (Appleton Culina House- Culina Group) Unilever',
    'Geodis D&E Normandie',
    'Great Bear Distribution MV1',
    'Great Bear Port Salford Mars',
    'H.J. Heinz BV',
    'home24 eLogistics GmbH & Co. KG',
    'Hoover Ltd',
    'iRobot UK Ltd',
    'JACOBS DOUWE EGBERTS GB LTD',
    'Kellogg Marketing And Sales Company (UK) Limited',
    'Keter Italia S.p.A.',
    'Keter Iberia S.L.U',
    'Keter Germany Gmbh',
    'Keter France Sas',
    'Mars PF France',
    'Mars GmbH (CBT-DE)',
    'Mars GmbH FLOERSHEIM',
    'Mars GmbH MINDEN',
    'Messaggerie Libri Spa',
    'Mömax Logistik GmbH',
    'Mondi Logistik GmbH',
    'Nestlé Enterprises SA, Business Growth Solutions Division',
    'Nestlé UK',
    'Nestrade S.A. T-Hub Central',
    'Nestrade SA (t-hub North)',
    'Nestrade S.A. T-Hub South',
    'Nestrade T-hub West',
    'PepsiCo Deutschland GmbH',
    'Procter & Gamble International Operations SA',
    'REHAU Industries SE & Co. KG',
    'Robert Bosch Power Tools GmbH',
    'Skechers EDC',
    'SharkNinja Europe Ltd',
    'SharkNinja Germany Gmbh',
    'Schlaadt HighCut GmbH',
    'S.L. Systemlogistik GmbH',
    'Soffass spa',
    'Sofidel Germany GmbH',
    'Sofidel France SAS',
    'Sofidel Spain',
    'Sofidel UK',
    'tegut… gute Lebensmittel GmbH & Co. KG',
    'Tetra GmbH',
    'The Book Service Limited',
    'Unilever Europe B.V. (UK)',
    'Unilever Europe B.V. (DE)',
    'Unilever Europe B.V. (EU)',
    'Vendor Returns',
    'Versuni Nederland B.V',
    'Walkers Snacks Distribution Ltd',
    'Wincanton (J SAINSBURY PLC)',
    'Yankee Candle Co (Europe) LTD',
    'Yankee Candle Co - DE',
    'Zeitfracht Medien GmbH',
    'Danone Deutschland GmbH',
    'Danone UK Waters',
    'Danone FR',
    'Wacker Chemie AG',
    'Sharp Consumer Electronics Poland sp. z o.o.',
    'Cargill Poland Sp. z o.o.',
    'Nitto Advanced Film Gronau GmbH',
    'Cargill S.L.U.',
    'Coca-Cola Europacific Partners Deutschland GmbH',
    'Bio Springer',
    'La Palette Rouge Iberica s.a. succ.le in Italia',
    'COLGATE PALMOLIVE EUROPE',
    'ECOSCOOTING DELIVERY SL',
    'EDT BE SRL (TEMU)',
    "L 'Oreal Italy",
    'Hager Electro SAS',
    'Heineken Deutschland GmbH',
    'Euro Pool System UK Ltd',
    '3M EMEA GmbH',
    'Falken Tyre Europe GmbH',
    'SACHSENMILCH Leppersdorf GmbH',
    'XPO Transport Solutions UK Limited',
    'La Palette Rouge Iberica Sa',
    'LPR - LA PALETTE ROUGE (GB) LTD',
    'BONDUELLE RE',
    'HARIBO Sp. z o.o.',
    'Groupe SEB WMF Consumer GmbH',
    'HOYER GmbH Internationale Fachspedition',
    'BLACK & DECKER LIMITED BV',
    'Cycleon B.v.',
    'Wella International Operations Switzerland Sarl',
    'INTERFORUM',
    'Beiersdorf Customer Supply GmbH',
    'BDSK Handels GmbH & Co.KG',
    'DGL- Ingram (KSP) ES',
    'JYSK SE',
    'Mars Multisales Spain S.L.',
    'Mars Food Europe CV',
    'Nestrade SA',
    'Philips Consumer Lifestyle BV',
    'Pregis Ltd',
    'TYRE ECO CHAIN',
    'Zalando SE',
    'Tchibo GmbH',
]

_STOPWORDS = {
    'the', 'and', 'for', 'von', 'van', 'de', 'di', 'du', 'der', 'gmbh', 'bv', 'sa', 'ltd', 'llc', 'inc', 'co',
    'kg', 'spa', 'sas', 'sl', 'nv', 'ag', 'plc', 'ug', 'bvba', 'srl', 'spzoo'
}

def _de_umlaut_fold(s: str) -> str:
    if not isinstance(s, str):
        return ""
    s = (
        s.replace("Ä", "Ae").replace("Ö", "Oe").replace("Ü", "Ue")
         .replace("ä", "ae").replace("ö", "oe").replace("ü", "ue")
         .replace("ß", "ss")
    )
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s

def _normalise(s):
    if not isinstance(s, str):
        return ''
    s = _de_umlaut_fold(s)
    s = re.sub(r'[^\w\s]', '', s)
    s = re.sub(r'\s+', ' ', s)
    return s.strip().lower()

def _core_tokens(s):
    s = _de_umlaut_fold(str(s))
    words = re.sub(r'[^\w\s]', ' ', s).lower().split()
    return frozenset(w for w in words if len(w) > 2 and w not in _STOPWORDS)

_CST_EXACT = set(s.strip().lower() for s in CST_SHIPPERS)
_CST_FUZZY = set(_normalise(s) for s in CST_SHIPPERS)
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
FC_PATTERN = re.compile(r'^(?:[A-Z]{3}\d|[A-Z]{4})$')

REQUIRED_COLUMNS = [
    'Order ID', 'Source', 'Shipper', 'Destination Stop Date and Time',
    'Destination Stop Facility Name', 'Creation Date and Time', 'Created by'
]
REQUIRED_COLUMNS_CST = [
    'Order ID', 'Source', 'Shipper', 'Destination Stop Date and Time',
    'Creation Date and Time', 'Created by'
]

VISIBLE_STEP_ORDER = [1, 3, 4, 5]
VISIBLE_STEP_LABELS = {
    1: '1. Load File',
    3: '2. External Orders',
    4: '3. Portal Check',
    5: '4. Final Results',
}
VISIBLE_STEP_NUMBER = {
    1: 1,
    3: 2,
    4: 3,
    5: 4,
}

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

@st.cache_data(show_spinner=False)
def _make_copy_block_cached(df: pd.DataFrame, exclude_cols: tuple[str, ...]) -> str:
    if df is None or df.empty:
        return ""
    out = df.copy()
    for c in exclude_cols:
        if c in out.columns:
            out = out.drop(columns=[c])
    out = out.fillna("")
    lines = ["\t".join(map(str, row)) for row in out.to_numpy()]
    return "\n".join(lines)

def make_copy_block(df: pd.DataFrame, exclude_cols: list[str]) -> str:
    return _make_copy_block_cached(df, tuple(exclude_cols))

def process_step2_backend(df_raw: pd.DataFrame) -> pd.DataFrame:
    df = df_raw.copy()
    cm = {col.strip().lower(): col for col in df.columns}

    df_str = df.astype(str)
    dummy_mask = df_str.apply(lambda col: col.str.strip().str.lower().eq('dummy'), axis=0).any(axis=1)

    shipper_col = cm.get('shipper')
    if shipper_col:
        test_mask = df[shipper_col].astype(str).str.contains('test', case=False, na=False)
    else:
        test_mask = pd.Series([False] * len(df), index=df.index)

    remove_mask = dummy_mask | test_mask
    df = df.loc[~remove_mask].copy()

    cols = list(df.columns)
    if len(cols) >= 2:
        old_b = cols[1]
        df = df.rename(columns={old_b: 'Source'})

    cm2 = {col.strip().lower(): col for col in df.columns}
    keep_cols = [cm2[c.lower()] for c in REQUIRED_COLUMNS if c.lower() in cm2]
    df = df[keep_cols].copy()

    cm3 = {col.strip().lower(): col for col in df.columns}
    created_by_col = cm3.get('created by')
    if created_by_col:
        df['Source'] = df[created_by_col].apply(classify_source)

    return df

def render_inline_copy_button(text: str, button_text: str = "Copy"):
    if not text:
        return

    btn_id = f"copy_btn_{uuid.uuid4().hex}"
    text_id = f"copy_text_{uuid.uuid4().hex}"

    safe_text = html.escape(text)
    safe_button_text = html.escape(button_text)

    components.html(
        f"""
        <div style="display:flex; justify-content:flex-end; margin-top:0.25rem; margin-bottom:0.25rem;">
            <textarea id="{text_id}" readonly style="position:absolute; left:-9999px; top:-9999px;">{safe_text}</textarea>
            <button
                id="{btn_id}"
                onclick="
                    const btn = document.getElementById('{btn_id}');
                    const textarea = document.getElementById('{text_id}');
                    const originalText = btn.innerText;

                    const showCopied = () => {{
                        btn.innerText = 'Copied ✓';
                        btn.style.background = '#d1fae5';
                        btn.style.border = '1px solid #10b981';
                        setTimeout(() => {{
                            btn.innerText = originalText;
                            btn.style.background = '#f0f2f6';
                            btn.style.border = '1px solid #999';
                        }}, 1500);
                    }};

                    const showFailed = () => {{
                        btn.innerText = 'Copy failed';
                        btn.style.background = '#fee2e2';
                        btn.style.border = '1px solid #ef4444';
                        setTimeout(() => {{
                            btn.innerText = originalText;
                            btn.style.background = '#f0f2f6';
                            btn.style.border = '1px solid #999';
                        }}, 1500);
                    }};

                    const copyWithFallback = () => {{
                        textarea.focus();
                        textarea.select();
                        try {{
                            const ok = document.execCommand('copy');
                            if (ok) {{
                                showCopied();
                            }} else {{
                                showFailed();
                            }}
                        }} catch (e) {{
                            showFailed();
                        }}
                    }};

                    if (navigator.clipboard && window.isSecureContext) {{
                        navigator.clipboard.writeText(textarea.value)
                            .then(() => showCopied())
                            .catch(() => copyWithFallback());
                    }} else {{
                        copyWithFallback();
                    }}
                "
                style="
                    padding: 0.35rem 0.75rem;
                    border-radius: 0.5rem;
                    border: 1px solid #999;
                    cursor: pointer;
                    background: #f0f2f6;
                    font-size: 0.9rem;
                    font-weight: 500;
                    white-space: nowrap;
                "
            >
                {safe_button_text}
            </button>
        </div>
        """,
        height=45,
    )

def render_table_with_copy(title: str, df: pd.DataFrame, copy_text: str, button_text: str):
    left, right = st.columns([6, 1])

    with left:
        st.subheader(title)

    with right:
        if df is not None and not df.empty and copy_text:
            render_inline_copy_button(copy_text, button_text=button_text)

    st.dataframe(reset_index_display(df), use_container_width=True)

def render_portal_link(url: str):
    safe_url = html.escape(url, quote=True)
    components.html(
        f"""
        <div style="display:flex; justify-content:center; margin: 0.35rem 0 0.9rem 0; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;">
            <a
                href="{safe_url}"
                target="_blank"
                style="
                    display:inline-block;
                    text-decoration:none;
                    padding: 0.55rem 1rem;
                    border-radius: 0.6rem;
                    border: 1px solid #3b82f6;
                    background: rgba(59,130,246,0.14);
                    color: #dbeafe;
                    font-weight: 600;
                    font-size: 0.95rem;
                    text-align:center;
                "
            >
                Open Unified Portal
            </a>
        </div>
        """,
        height=60,
    )

def render_portal_batch_card(label: str, subtitle: str, text: str, button_text: str = "Copy Batch", box_height: int = 260):
    if not text:
        return

    btn_id = f"portal_copy_btn_{uuid.uuid4().hex}"
    text_id = f"portal_copy_text_{uuid.uuid4().hex}"
    safe_label = html.escape(label)
    safe_subtitle = html.escape(subtitle)
    safe_text = html.escape(text)
    safe_button_text = html.escape(button_text)

    components.html(
        f"""
        <div style="margin-bottom: 0.75rem; max-width: 280px; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;">
            <div
                style="
                    width: 100%;
                    background: #1f2937;
                    border-radius: 0.7rem;
                    padding: 0.9rem 1rem 1rem 1rem;
                    box-sizing: border-box;
                "
            >
                <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:0.65rem; gap:0.5rem;">
                    <div>
                        <div style="font-weight:700; font-size:1rem; color:#f9fafb; margin-bottom:0.18rem;">{safe_label}</div>
                        <div style="font-size:0.95rem; color:#d1d5db;">{safe_subtitle}</div>
                    </div>

                    <div style="flex-shrink:0;">
                        <textarea id="{text_id}" readonly style="position:absolute; left:-9999px; top:-9999px;">{safe_text}</textarea>
                        <button
                            id="{btn_id}"
                            onclick="
                                const btn = document.getElementById('{btn_id}');
                                const textarea = document.getElementById('{text_id}');
                                const originalText = btn.innerText;

                                const showCopied = () => {{
                                    btn.innerText = 'Copied ✓';
                                    btn.style.background = '#d1fae5';
                                    btn.style.border = '1px solid #10b981';
                                    setTimeout(() => {{
                                        btn.innerText = originalText;
                                        btn.style.background = '#f0f2f6';
                                        btn.style.border = '1px solid #999';
                                    }}, 1500);
                                }};

                                const showFailed = () => {{
                                    btn.innerText = 'Copy failed';
                                    btn.style.background = '#fee2e2';
                                    btn.style.border = '1px solid #ef4444';
                                    setTimeout(() => {{
                                        btn.innerText = originalText;
                                        btn.style.background = '#f0f2f6';
                                        btn.style.border = '1px solid #999';
                                    }}, 1500);
                                }};

                                const copyWithFallback = () => {{
                                    textarea.focus();
                                    textarea.select();
                                    try {{
                                        const ok = document.execCommand('copy');
                                        if (ok) {{
                                            showCopied();
                                        }} else {{
                                            showFailed();
                                        }}
                                    }} catch (e) {{
                                        showFailed();
                                    }}
                                }};

                                if (navigator.clipboard && window.isSecureContext) {{
                                    navigator.clipboard.writeText(textarea.value)
                                        .then(() => showCopied())
                                        .catch(() => copyWithFallback());
                                }} else {{
                                    copyWithFallback();
                                }}
                            "
                            style="
                                padding: 0.3rem 0.65rem;
                                border-radius: 0.5rem;
                                border: 1px solid #999;
                                cursor: pointer;
                                background: #f0f2f6;
                                font-size: 0.85rem;
                                font-weight: 500;
                                white-space: nowrap;
                                color: #111827;
                            "
                        >
                            {safe_button_text}
                        </button>
                    </div>
                </div>

                <div
                    style="
                        width: 100%;
                        height: {box_height}px;
                        overflow-y: auto;
                        overflow-x: hidden;
                        background: rgba(255,255,255,0.04);
                        border-radius: 0.6rem;
                        padding: 0.95rem 1rem;
                        box-sizing: border-box;
                        font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, 'Liberation Mono', 'Courier New', monospace;
                        font-size: 0.98rem;
                        line-height: 1.55;
                        white-space: pre;
                        color: #f9fafb;
                    "
                >{safe_text}</div>
            </div>
        </div>
        """,
        height=box_height + 120,
    )

def _norm_col(s: str) -> str:
    return re.sub(r'[\s_\-]+', '', str(s).strip().lower())

def _norm_val(s: str) -> str:
    return re.sub(r'\s+', ' ', str(s).strip().lower())

def extract_arrival_scheduled_ids_from_unified_portal_csv(df: pd.DataFrame):
    if df is None or df.empty:
        return []
    norm_map = {_norm_col(c): c for c in df.columns}
    search_col = norm_map.get('searchid')
    status_col = norm_map.get('appointmentstatus')
    if not search_col or not status_col:
        return None
    s_search = df[search_col].astype(str).map(str.strip)
    s_status = df[status_col].astype(str).map(_norm_val)
    mask = (s_status == 'arrival scheduled')
    ids = s_search[mask].dropna().astype(str).str.strip().tolist()
    return list(dict.fromkeys([x for x in ids if x]))

def render_wrapped_batches(batch_texts, per_row=3, box_height=260):
    if not batch_texts:
        return
    per_row = max(1, int(per_row))
    rows = math.ceil(len(batch_texts) / per_row)
    idx = 0
    for _ in range(rows):
        cols = st.columns(per_row)
        for c in range(per_row):
            if idx >= len(batch_texts):
                break
            b = batch_texts[idx]
            with cols[c]:
                render_portal_batch_card(
                    label=b["label"],
                    subtitle=b["subtitle"],
                    text=b["text"],
                    button_text="Copy Batch",
                    box_height=box_height
                )
            idx += 1

def run_cross_reference():
    ds5 = st.session_state.df_step4
    cm = {col.strip().lower(): col for col in ds5.columns}
    oic = cm.get('order id')
    if not oic:
        st.error("Missing 'Order ID' column for Portal Check.")
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
    st.session_state.step = 5
    st.rerun()

def go_back_one_step():
    cur = int(st.session_state.step or 1)

    if cur == 3:
        st.session_state.step = 1
    elif cur == 4 and st.session_state.get("step3_skipped", False):
        st.session_state.step = 1
    else:
        if cur == 5:
            st.session_state.step = 4
        elif cur == 4:
            st.session_state.step = 3
        else:
            st.session_state.step = 1

    st.rerun()

def scroll_to_top():
    components.html(
        """
        <script>
          const main = window.parent.document.querySelector('section.main');
          if (main) { main.scrollTo(0,0); }
          window.parent.scrollTo(0,0);
        </script>
        """,
        height=0
    )

defaults = {
    'step': 1,
    'df_raw': None,
    'df_formatted': None,
    'df_step4': None,
    'cst_ext': None,
    'non_cst_ext': None,
    'portal_ids': [],
    'cst_final': None,
    'non_cst_final': None,
    'unmatched_count': 0,
    'step3_skipped': False,
    'arrival_ids_ready': False,
    'portal_export_filenames': [],
    'last_step': None,
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

if st.session_state.last_step is None:
    st.session_state.last_step = st.session_state.step
elif st.session_state.step != st.session_state.last_step:
    scroll_to_top()
    st.session_state.last_step = st.session_state.step

st.title('Uncovered Orders Audit')
st.caption('Amazon Freight Scheduling Team - Automated Audit Workflow')
st.divider()

visible_step_count = len(VISIBLE_STEP_ORDER)
current_visible_step_number = VISIBLE_STEP_NUMBER.get(st.session_state.step, 1)
current_visible_step_label = VISIBLE_STEP_LABELS.get(st.session_state.step, VISIBLE_STEP_LABELS[1])
pv = (current_visible_step_number - 1) / (visible_step_count - 1)

st.progress(
    pv,
    text='Step {} of {}: {}'.format(
        current_visible_step_number,
        visible_step_count,
        current_visible_step_label.split('. ', 1)[1]
    )
)
st.divider()

if st.session_state.step == 1:
    st.header('Step 1 - Upload SMC Export File')
    st.info('Download the uncovered orders file from SMC TMS (untick LTL, intermodal) export, then upload it here to begin the audit.')

    uploaded = st.file_uploader(
        'Upload your SMC uncovered orders export (.xlsx, .xls, or .csv)',
        type=['xlsx', 'xls', 'csv'],
        key='smc_upload'
    )

    if uploaded is not None:
        try:
            df = load_smc_file(uploaded)
            st.session_state.df_raw = df
            st.success('File loaded: {} orders, {} columns detected.'.format(len(df), len(df.columns)))
            st.dataframe(df.head(10), use_container_width=True)
            st.caption('Showing first 10 of {} rows.'.format(len(df)))

            if st.button('Proceed to Step 2 - External Orders', type='primary'):
                st.session_state.df_formatted = process_step2_backend(df)
                st.session_state.step = 3
                st.rerun()

        except Exception as e:
            st.error('Error reading file: {}. Please check the file and try again.'.format(e))

elif st.session_state.step == 3:
    st.header('Step 2 - Process External Orders')

    df = st.session_state.df_formatted.copy()
    cm = {col.strip().lower(): col for col in df.columns}
    fc = cm.get('destination stop facility name')
    shc = cm.get('shipper')

    if not fc or not shc:
        st.error("Missing required columns for this step. Ensure the export contains 'Destination Stop Facility Name' and 'Shipper'.")
        st.stop()

    df['_is_fc'] = df[fc].apply(is_fc_facility)
    ext = df[df['_is_fc'] == False].copy()
    intr = df[df['_is_fc'] == True].copy()

    if ext.empty:
        st.session_state.cst_ext = pd.DataFrame(columns=REQUIRED_COLUMNS_CST)
        st.session_state.non_cst_ext = pd.DataFrame(columns=REQUIRED_COLUMNS)
        st.session_state.df_step4 = intr.drop(columns=['_is_fc'])
        st.session_state.portal_ids = []
        st.session_state.arrival_ids_ready = False
        st.session_state.portal_export_filenames = []
        st.session_state.step3_skipped = True
        st.session_state.step = 4
        st.rerun()

    c1, c2 = st.columns(2)
    c1.metric('Internal (FC-bound) Orders', len(intr))
    c2.metric('External (non-FC) Orders', len(ext))

    ext['_is_cst'] = ext[shc].apply(is_cst_shipper)
    cst_ext_raw = ext[ext['_is_cst'] == True].drop(columns=['_is_fc', '_is_cst'])
    non_cst_ext = ext[ext['_is_cst'] == False].drop(columns=['_is_fc', '_is_cst'])

    cst_cols_to_keep = [c for c in cst_ext_raw.columns if c.lower() != 'destination stop facility name']
    cst_ext = cst_ext_raw[cst_cols_to_keep].copy()

    c1, c2 = st.columns(2)
    c1.metric('CST External Orders', len(cst_ext))
    c2.metric('Non-CST External Orders', len(non_cst_ext))

    st.divider()
    cst_ext_copy = make_copy_block(cst_ext, exclude_cols=['Created by'])
    render_table_with_copy(
        title='CST External Orders - copy to CST Task Sheet (Uncovered tab)',
        df=cst_ext,
        copy_text=cst_ext_copy,
        button_text='Copy to CST Sheet'
    )

    non_cst_ext_copy = make_copy_block(non_cst_ext, exclude_cols=['Created by'])
    render_table_with_copy(
        title='Non-CST External Orders - copy to AF Scheduling Daily Task Workbook (Uncovered tab)',
        df=non_cst_ext,
        copy_text=non_cst_ext_copy,
        button_text='Copy to Scheduling Sheet'
    )

    st.divider()
    st.warning(
        "ACTION REQUIRED\n"
        "1. Copy CST External Orders to the CST Task Sheet (Uncovered tab).\n"
        "2. Copy Non-CST External Orders to the AF Scheduling Daily Task Workbook (Uncovered tab).\n"
        "3. Click Done below once you have completed this step."
    )

    c1, c2 = st.columns(2)
    with c1:
        if st.button('Back a step'):
            go_back_one_step()
    with c2:
        if st.button('Done - Proceed to Step 3', type='primary'):
            st.session_state.cst_ext = cst_ext
            st.session_state.non_cst_ext = non_cst_ext
            st.session_state.df_step4 = intr.drop(columns=['_is_fc'])
            st.session_state.portal_ids = []
            st.session_state.arrival_ids_ready = False
            st.session_state.portal_export_filenames = []
            st.session_state.step3_skipped = False
            st.session_state.step = 4
            st.rerun()

elif st.session_state.step == 4:
    st.header('Step 3 - Unified Portal ISA Check')

    if st.session_state.step3_skipped:
        st.info("External Orders = 0, so the tool skipped straight to Portal Check.")

    ds4 = st.session_state.df_step4
    cm = {col.strip().lower(): col for col in ds4.columns}
    oic = cm.get('order id')
    if not oic:
        st.error("Missing 'Order ID' column for Portal Check.")
        st.stop()

    rids = ds4[oic].dropna().astype(str).str.strip().tolist()
    st.info('{} Order IDs need to be checked in the Unified Portal.'.format(len(rids)))

    render_portal_link(UNIFIED_PORTAL_URL)

    batch_size = 50
    total = len(rids)
    batch_count = max(1, math.ceil(total / batch_size))

    with st.expander('View / Copy Order IDs (copy-paste into Unified Portal)'):
        if total == 0:
            st.warning("No Order IDs available to copy.")
        else:
            st.caption("Batches are displayed in a wrapped grid. Use the copy button on each batch.")

            batches = []
            for i in range(batch_count):
                start = i * batch_size
                end = min(start + batch_size, total)
                batch_ids = rids[start:end]
                if not batch_ids:
                    continue
                batches.append({
                    "label": f"Batch {i+1}",
                    "subtitle": f"{start+1}–{end} of {total}",
                    "text": "\n".join(batch_ids),
                })

            render_wrapped_batches(batches, per_row=3, box_height=260)

    st.divider()
    st.subheader('Upload Unified Portal Results CSV(s)')

    if st.button("Reset Portal Inputs", key="reset_step4"):
        st.session_state.portal_ids = []
        st.session_state.arrival_ids_ready = False
        st.session_state.portal_export_filenames = []
        if 'portal_export_upload_multi' in st.session_state:
            del st.session_state['portal_export_upload_multi']
        if 'manual_arrivals_paste' in st.session_state:
            del st.session_state['manual_arrivals_paste']
        st.rerun()

    method = st.radio(
        "How do you want to provide Unified Portal results?",
        ["Upload Unified Portal CSV export(s) (recommended)", "Paste Arrival Scheduled Order IDs manually (fallback)"],
        horizontal=True
    )

    if method.startswith("Upload"):
        portal_csvs = st.file_uploader(
            "Upload Unified Portal export CSV file(s). You can upload multiple files (one per 50-ID search batch). "
            "Each CSV must contain columns: searchId and appointmentStatus.",
            type=['csv'],
            accept_multiple_files=True,
            key='portal_export_upload_multi'
        )

        if portal_csvs:
            all_arrival_ids = []
            files_ok = 0
            files_missing_cols = 0
            files_read_error = 0
            zero_arrivals = 0
            filenames = []

            for f in portal_csvs:
                filenames.append(getattr(f, "name", "unknown.csv"))
                try:
                    raw = f.read()
                    pdf = pd.read_csv(io.BytesIO(raw), dtype=str)

                    arrival_ids = extract_arrival_scheduled_ids_from_unified_portal_csv(pdf)
                    if arrival_ids is None:
                        files_missing_cols += 1
                        continue

                    files_ok += 1
                    if len(arrival_ids) == 0:
                        zero_arrivals += 1
                    else:
                        all_arrival_ids.extend(arrival_ids)

                except Exception:
                    files_read_error += 1

            all_arrival_ids = list(dict.fromkeys([x for x in all_arrival_ids if x]))

            if files_missing_cols > 0:
                st.error(
                    f"{files_missing_cols} file(s) did not contain required columns "
                    f"('searchId' and 'appointmentStatus')."
                )
            if files_read_error > 0:
                st.error(f"{files_read_error} file(s) could not be read as CSV.")

            st.info(
                f"Files uploaded: {len(portal_csvs)} | Parsed OK: {files_ok} | "
                f"0 arrivals: {zero_arrivals} | Arrival Scheduled IDs extracted: {len(all_arrival_ids)}"
            )

            with st.expander("Show uploaded filenames"):
                st.write(filenames)

            if len(all_arrival_ids) == 0:
                st.session_state.portal_ids = []
                st.session_state.arrival_ids_ready = False
                st.session_state.portal_export_filenames = filenames
                st.warning("No 'Arrival Scheduled' rows were found across the uploaded file(s).")
            else:
                st.session_state.portal_ids = all_arrival_ids
                st.session_state.arrival_ids_ready = True
                st.session_state.portal_export_filenames = filenames
                st.success(
                    f"Done. Combined {len(all_arrival_ids)} unique 'Arrival Scheduled' Order IDs "
                    f"from {files_ok} CSV file(s)."
                )
                with st.expander("Preview extracted Arrival Scheduled IDs"):
                    st.caption("One-click copy via copy icon (top-right of code box)")
                    st.code("\n".join(all_arrival_ids), language=None)

    else:
        pasted = st.text_area(
            "Paste (Arrival Scheduled) Order IDs here:",
            height=200,
            key="manual_arrivals_paste"
        )
        if pasted.strip():
            ri = [l.strip() for l in pasted.strip().splitlines() if l.strip()]
            ids = list(dict.fromkeys(ri))
            if ids:
                st.session_state.portal_ids = ids
                st.session_state.arrival_ids_ready = True
                st.session_state.portal_export_filenames = []
                st.success(f"{len(ids)} unique Order IDs entered.")
                with st.expander("Preview / Copy pasted IDs"):
                    st.caption("One-click copy via copy icon (top-right of code box)")
                    st.code("\n".join(ids), language=None)
        else:
            st.info("Paste IDs to enable the cross-reference button.")

    st.divider()
    ready = bool(st.session_state.arrival_ids_ready and len(st.session_state.portal_ids) > 0)
    if ready:
        st.info("Arrival Scheduled Order IDs are ready. You can now run the final cross-reference.")

    c1, c2 = st.columns(2)
    with c1:
        if st.button('Back a step'):
            go_back_one_step()
    with c2:
        run_clicked = st.button(
            'Run Cross-Reference and Produce Final Results',
            type='primary',
            disabled=(not ready)
        )

    if run_clicked:
        run_cross_reference()

elif st.session_state.step == 5:
    st.header('Step 4 - Audit Complete')
    st.balloons()

    cf = st.session_state.cst_final if st.session_state.cst_final is not None else pd.DataFrame(columns=REQUIRED_COLUMNS_CST)
    ncf = st.session_state.non_cst_final if st.session_state.non_cst_final is not None else pd.DataFrame(columns=REQUIRED_COLUMNS)
    uc = int(st.session_state.unmatched_count or 0)

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
    cf_copy = make_copy_block(cf_clean, exclude_cols=['Created by'])
    render_table_with_copy(
        title='CST Orders - copy to CST Task Sheet (Uncovered tab)',
        df=cf_clean,
        copy_text=cf_copy,
        button_text='Copy to CST Sheet'
    )

    ncf_copy = make_copy_block(ncf, exclude_cols=['Created by'])
    render_table_with_copy(
        title='Non-CST Orders - copy to AF Scheduling Daily Task Workbook (Uncovered tab)',
        df=ncf,
        copy_text=ncf_copy,
        button_text='Copy to Scheduling Sheet'
    )

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
            'smc_upload',
            'portal_export_upload_multi',
            'manual_arrivals_paste',
        ]
        for k in keys_to_clear:
            if k in st.session_state:
                del st.session_state[k]
        for k, v in defaults.items():
            if k not in st.session_state:
                st.session_state[k] = v
        st.session_state.step = 1
        st.rerun()
