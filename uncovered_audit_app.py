
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
        cst_ext = ext[ext['_is_cst'] == True].drop(columns=['_is_fc','_is_cst'])
        non_cst_ext = ext[ext['_is_cst'] == False].drop(columns=['_is_fc','_is_cst'])
    else:
        cst_ext = pd.DataFrame(columns=REQUIRED_COLUMNS)
        non_cst_ext = pd.DataFrame(columns=REQUIRED_COLUMNS)
    c1,c2 = st.columns(2)
    c1.metric('CST External Orders', len(cst_ext))
    c2.metric('Non-CST External Orders', len(non_cst_ext))
    eb = to_excel_bytes({'CST External Orders': cst_ext if not cst_ext.empty else pd.DataFrame(columns=REQUIRED_COLUMNS),'Non-CST External Orders': non_cst_ext if not non_cst_ext.empty else pd.DataFrame(columns=REQUIRED_COLUMNS)})
    st.divider()
    st.subheader('Download External Orders File')
    st.download_button(label='Download Step4_External_Orders.xlsx', data=eb, file_name='Step4_External_Orders.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    st.divider()
    st.subheader('CST External Orders - copy to CST Task Sheet (Uncovered tab)')
    st.dataframe(cst_ext, use_container_width=True)
    st.subheader('Non-CST External Orders - copy to AF Scheduling Daily Task Workbook (Uncovered tab)')
    st.dataframe(non_cst_ext, use_container_width=True)
    st.divider()
    st.warning(chr(10).join(['ACTION REQUIRED','','1. Download the file above.','2. Copy CST External Orders sheet to the CST Task Sheet (Uncovered tab).','3. Copy Non-CST External Orders sheet to the AF Scheduling Daily Task Workbook (Uncovered tab).','','Click Done below once you have completed this step.']))
    c1,c2 = st.columns(2)
    with c1:
        if st.button('Back to Step 3'): st.session_state.step = 3; st.rerun()
    with c2:
        if st.button('Done - Proceed to Step 5', type='primary'):
            st.session_state.cst_ext = cst_ext
            st.session_state.non_cst_ext = non_cst_ext
            st.session_state.df_step5 = intr.drop(columns=['_is_fc'])
            st.session_state.step = 5; st.rerun()

elif st.session_state.step == 5:
    st.header('Step 5 - Unified Portal ISA Check')
    ds5 = st.session_state.df_step5
    cm = {col.strip().lower(): col for col in ds5.columns}
    oic = cm.get('order id')
    rids = ds5[oic].dropna().astype(str).str.strip().tolist()
    st.info('{} Order IDs need to be checked in the Unified Portal.'.format(len(rids)))
    ie = to_excel_bytes({'Order IDs': pd.DataFrame({'Order ID': rids})})
    st.download_button(label='Download Order IDs for Unified Portal', data=ie, file_name='Step5_Order_IDs_for_Portal.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    with st.expander('View / Copy Order IDs (copy-paste into Unified Portal)'):
        st.code(chr(10).join(rids), language=None)
    st.divider()
    st.subheader('Unified Portal Instructions')
    st.markdown(chr(10).join(['1. Download the file above (or copy the IDs from the panel above).','2. Go to the Unified Portal and set ID type to progressive number.','3. Paste the IDs in batches of 50 and click Submit.','4. Filter results where appointmentStatus = arrival scheduled.','5. Note the matching Order IDs.','6. Come back here and enter the results below.']))
    st.divider()
    st.subheader('Enter Unified Portal Results')
    im = st.radio('How would you like to provide the matching Order IDs?', ['Paste IDs directly','Upload a results file'], horizontal=True)
    portal_ids = []
    if im == 'Paste IDs directly':
        pasted = st.text_area('Paste matching Order IDs here (one per line):', height=200)
        if pasted.strip():
            ri = [l.strip() for l in pasted.strip().splitlines() if l.strip()]
            portal_ids = list(dict.fromkeys(ri))
            st.success('{} unique Order IDs entered.'.format(len(portal_ids)))
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
                portal_ids = list(dict.fromkeys(ri))
                st.success('{} unique Order IDs loaded from file.'.format(len(portal_ids)))
            except Exception as e:
                st.error('Error reading file: {}'.format(e))
    st.divider()
    c1,c2 = st.columns(2)
    with c1:
        if st.button('Back to Step 4'): st.session_state.step = 4; st.rerun()
    with c2:
        if portal_ids and st.button('Run Cross-Reference and Produce Final Results', type='primary'):
            ps = set(portal_ids)
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
    fb = to_excel_bytes({'CST Orders': cf if not cf.empty else pd.DataFrame(columns=REQUIRED_COLUMNS),'Non-CST Orders': ncf if not ncf.empty else pd.DataFrame(columns=REQUIRED_COLUMNS)})
    st.download_button(label='Download Step5_Final_Results.xlsx', data=fb, file_name='Step5_Final_Results.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    st.divider()
    st.subheader('CST Orders - copy to CST Task Sheet (Uncovered tab)')
    st.dataframe(cf, use_container_width=True)
    st.subheader('Non-CST Orders - copy to AF Scheduling Daily Task Workbook (Uncovered tab)')
    st.dataframe(ncf, use_container_width=True)
    st.divider()
    st.warning(chr(10).join(['FINAL ACTION REQUIRED','','1. Download the file above.','2. Copy CST Orders sheet to the CST Task Sheet (Uncovered tab).','3. Copy Non-CST Orders sheet to the AF Scheduling Daily Task Workbook (Uncovered tab).','','Audit complete!']))
    st.divider()
    if st.button('Start a New Audit', type='primary'):
        for k in list(st.session_state.keys()): del st.session_state[k]
        st.rerun()