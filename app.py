import streamlit as st
import pandas as pd
import json, os, glob, datetime, io, zipfile
import engine

def create_zip_of_images(png_dict, all_png):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        if all_png and os.path.exists(all_png):
            zf.write(all_png, os.path.basename(all_png))
        for c, p in png_dict.items():
            if p and os.path.exists(p):
                zf.write(p, os.path.basename(p))
    buf.seek(0)
    return buf

# Page configuration
st.set_page_config(
    page_title="Harley's OOS Automation Suite",
    page_icon="🍰",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for polished styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1B365D;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        font-size: 1.05rem;
        color: #555555;
        margin-bottom: 1.5rem;
    }
    .metric-card {
        background-color: #F8F9FA;
        border-radius: 8px;
        padding: 16px;
        border-left: 5px solid #1B365D;
        box-shadow: 0 1px 3px rgba(0,0,0,0.08);
    }
    div[data-testid="stMetricValue"] {
        font-size: 1.8rem;
        font-weight: 700;
        color: #1B365D;
    }
</style>
""", unsafe_allow_html=True)

# Sidebar
st.sidebar.markdown("# 🍰 Harley's OOS Suite")
st.sidebar.caption("Automated Morning & Evening Intelligence")

nav_choice = st.sidebar.radio(
    "Navigation",
    ["🚀 Run OOS Reports", "🕒 Outlet Timings Manager", "📁 Browse Past Reports"],
    index=0
)

# -------------------------------------------------------------
# TAB 1: RUN OOS REPORTS
# -------------------------------------------------------------
if nav_choice == "🚀 Run OOS Reports":
    st.markdown('<div class="main-header">🚀 Run Automated OOS Reports</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">Upload your daily change history CSV to generate Morning & Evening reports, multi-tab Excel workbooks, and city-wise screenshots.</div>', unsafe_allow_html=True)

    col_up, col_opts = st.columns([1.6, 1.0])
    
    with col_up:
        uploaded_file = st.file_uploader(
            "Upload 'Item sold out change history' CSV",
            type=["csv"],
            help="Drag and drop or select the daily change history export file"
        )
        
        # Also check if recent CSVs exist in Downloads
        recent_csvs = glob.glob(os.path.join(engine.DOWNLOADS_DIR, "*change history*.csv"))
        recent_csvs.sort(key=os.path.getmtime, reverse=True)
        
        selected_local_file = None
        if not uploaded_file and recent_csvs:
            st.info("💡 Or pick a recent export file:")
            selected_local_file = st.selectbox(
                "Recent CSV Exports:",
                options=[None] + recent_csvs[:10],
                format_func=lambda x: "Choose export..." if x is None else os.path.basename(x)
            )

    active_csv = uploaded_file or selected_local_file

    with col_opts:
        st.markdown("#### Run Settings")
        
        detected_date = None
        if active_csv:
            try:
                if uploaded_file:
                    preview_df = pd.read_csv(uploaded_file, nrows=20)
                    uploaded_file.seek(0)
                else:
                    preview_df = pd.read_csv(selected_local_file, nrows=20)
                if 'Date and Time' in preview_df.columns:
                    detected_date = pd.to_datetime(preview_df['Date and Time']).dt.date.mode()[0]
            except Exception:
                pass
                
        target_date = st.date_input(
            "Report Date:",
            value=detected_date or datetime.date.today(),
            help="The business date for this report."
        )
        
        day_of_week = target_date.strftime('%A')
        st.markdown(f"**Day of Week:** `{day_of_week}` *(outlet opening timings automatically adjust)*")
        
        report_scope = st.radio(
            "Generate Scope:",
            ["Both Morning & Evening", "Morning Only", "Evening Only"],
            index=0
        )

    st.markdown("---")
    
    if active_csv:
        if st.button("⚡ Run Automated Analysis Now", type="primary", use_container_width=True):
            run_m = report_scope in ["Both Morning & Evening", "Morning Only"]
            run_e = report_scope in ["Both Morning & Evening", "Evening Only"]
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            def update_progress(msg):
                status_text.info(msg)
            
            try:
                progress_bar.progress(20)
                csv_payload = uploaded_file if uploaded_file else selected_local_file
                
                results = engine.run_oos_pipeline(
                    csv_input=csv_payload,
                    target_date=target_date,
                    run_morning=run_m,
                    run_evening=run_e,
                    progress_cb=update_progress
                )
                
                progress_bar.progress(100)
                status_text.success("✅ Analysis completed successfully!")
                st.session_state['last_results'] = results
                st.rerun()
                
            except Exception as e:
                status_text.error(f"❌ Error during processing: {str(e)}")
                st.exception(e)

    # Display Results if available
    if 'last_results' in st.session_state:
        res = st.session_state['last_results']
        st.markdown(f"### 📋 Results Summary for `{res['date_str']}` ({res['day_name']})")
        
        # Metric KPI cards
        kpi_cols = st.columns(4)
        if res.get('morning'):
            kpi_cols[0].metric("Morning Outlets", res['morning']['total_outlets'])
            kpi_cols[1].metric("Morning OOS Items", res['morning']['total_items'])
        if res.get('evening'):
            kpi_cols[2].metric("Evening Still OUT (5 PM)", res['evening']['still_out'])
            kpi_cols[3].metric("Evening Restored IN", res['evening']['restored_in'])

        if hasattr(os, 'startfile'):
            if st.button("📂 Open Output Folder"):
                os.startfile(res['date_dir'])

        # Detailed Report Tabs
        tab_m, tab_e = st.tabs(["🌅 Morning Report", "🌆 Evening Report"])
        
        # Morning Tab
        with tab_m:
            if res.get('morning'):
                m_data = res['morning']
                st.markdown(f"#### Morning OOS ({m_data['total_items']} items across {m_data['total_outlets']} outlets)")
                
                # Download actions bar
                dl_m1, dl_m2 = st.columns(2)
                with dl_m1:
                    with open(m_data['excel_path'], "rb") as fp:
                        st.download_button(
                            label="📥 Download Morning Excel Workbook (.xlsx)",
                            data=fp,
                            file_name=os.path.basename(m_data['excel_path']),
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                with dl_m2:
                    m_zip_buf = create_zip_of_images(m_data['city_pngs'], m_data['png_all'])
                    st.download_button(
                        label="📦 Download All Morning Screenshots (.zip)",
                        data=m_zip_buf,
                        file_name=f"Morning_OOS_Screenshots_{res['date_dash']}.zip",
                        mime="application/zip",
                        use_container_width=True
                    )

                city_sel_m = st.selectbox(
                    "View Morning Data & Screenshot for:",
                    ["ALL CITIES", "AP", "BLR", "HYD", "MUM", "PUN"],
                    key="morning_city_select"
                )
                
                df_m = pd.DataFrame(m_data['items'])
                if city_sel_m != "ALL CITIES":
                    df_m = df_m[df_m['city'] == city_sel_m]
                
                st.dataframe(
                    df_m[['city', 'odoo', 'open_time', 'rank', 'item', 'resolved_in']].rename(
                        columns={'city': 'City', 'odoo': 'Outlet Name', 'open_time': 'Opening Timing', 'rank': 'Top 20 Rank', 'item': 'Item Name', 'resolved_in': 'Restored At'}
                    ),
                    use_container_width=True,
                    hide_index=True
                )
                
                img_path = m_data['png_all'] if city_sel_m == "ALL CITIES" else m_data['city_pngs'].get(city_sel_m)
                if img_path and os.path.exists(img_path):
                    with open(img_path, "rb") as img_f:
                        st.download_button(
                            label=f"📸 Download {city_sel_m} Screenshot ({os.path.basename(img_path)})",
                            data=img_f,
                            file_name=os.path.basename(img_path),
                            mime="image/png",
                            type="secondary",
                            use_container_width=True
                        )
                    st.image(img_path, caption=f"Morning Screenshot: {city_sel_m}", use_container_width=True)
            else:
                st.info("Morning report was not selected for this run.")

        # Evening Tab
        with tab_e:
            if res.get('evening'):
                e_data = res['evening']
                st.markdown(f"#### Evening OOS ({e_data['total_items']} total items: {e_data['still_out']} Still OUT, {e_data['restored_in']} Restored IN)")
                
                # Download actions bar
                dl_e1, dl_e2 = st.columns(2)
                with dl_e1:
                    with open(e_data['excel_path'], "rb") as fp:
                        st.download_button(
                            label="📥 Download Evening Excel Workbook (.xlsx)",
                            data=fp,
                            file_name=os.path.basename(e_data['excel_path']),
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                with dl_e2:
                    e_zip_buf = create_zip_of_images(e_data['city_pngs'], e_data['png_all'])
                    st.download_button(
                        label="📦 Download All Evening Screenshots (.zip)",
                        data=e_zip_buf,
                        file_name=f"Evening_OOS_Screenshots_{res['date_dash']}.zip",
                        mime="application/zip",
                        use_container_width=True
                    )

                city_sel_e = st.selectbox(
                    "View Evening Data & Screenshot for:",
                    ["ALL CITIES", "AP", "BLR", "HYD", "MUM", "PUN"],
                    key="evening_city_select"
                )
                
                df_e = pd.DataFrame(e_data['items'])
                if city_sel_e != "ALL CITIES":
                    df_e = df_e[df_e['city'] == city_sel_e]
                
                st.dataframe(
                    df_e[['city', 'odoo', 'rank', 'item', 'status', 'note']].rename(
                        columns={'city': 'City', 'odoo': 'Outlet Name', 'rank': 'Top 20 Rank', 'item': 'Item Name', 'status': 'Status (5 PM)', 'note': 'Details'}
                    ),
                    use_container_width=True,
                    hide_index=True
                )
                
                img_path_e = e_data['png_all'] if city_sel_e == "ALL CITIES" else e_data['city_pngs'].get(city_sel_e)
                if img_path_e and os.path.exists(img_path_e):
                    with open(img_path_e, "rb") as img_f_e:
                        st.download_button(
                            label=f"📸 Download {city_sel_e} Screenshot ({os.path.basename(img_path_e)})",
                            data=img_f_e,
                            file_name=os.path.basename(img_path_e),
                            mime="image/png",
                            type="secondary",
                            use_container_width=True
                        )
                    st.image(img_path_e, caption=f"Evening Screenshot with Status: {city_sel_e}", use_container_width=True)
            else:
                st.info("Evening report was not selected for this run.")

# -------------------------------------------------------------
# TAB 2: OUTLET TIMINGS MANAGER
# -------------------------------------------------------------
elif nav_choice == "🕒 Outlet Timings Manager":
    st.markdown('<div class="main-header">🕒 Outlet Opening Timings Manager</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">View and update opening timings across all outlets. Configure specific hours for different days of the week (e.g. weekends vs weekdays).</div>', unsafe_allow_html=True)

    timings_db = engine.load_timings_db()
    
    if not timings_db:
        st.warning("No timings database found! Generating initial database from repository...")
        engine.load_mappings_and_history()
        timings_db = engine.load_timings_db()

    records = []
    for k, v in timings_db.items():
        records.append({
            'City': v.get('city', engine.get_city(k)),
            'Outlet Name': k,
            'Default Opening': v.get('default', '09:00'),
            'Monday': v.get('Monday', v.get('default', '09:00')),
            'Tuesday': v.get('Tuesday', v.get('default', '09:00')),
            'Wednesday': v.get('Wednesday', v.get('default', '09:00')),
            'Thursday': v.get('Thursday', v.get('default', '09:00')),
            'Friday': v.get('Friday', v.get('default', '09:00')),
            'Saturday': v.get('Saturday', v.get('default', '09:00')),
            'Sunday': v.get('Sunday', v.get('default', '09:00'))
        })

    df_timings = pd.DataFrame(records)
    diff_day_count = sum(1 for r in records if len(set([r['Monday'], r['Tuesday'], r['Wednesday'], r['Thursday'], r['Friday'], r['Saturday'], r['Sunday']])) > 1)
    
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Configured Outlets", len(df_timings))
    c2.metric("Outlets with Day Variations", diff_day_count)
    c3.metric("Supported Days", "Mon - Sun (7 Days)")

    st.markdown("---")
    
    f_col1, f_col2, f_col3 = st.columns([1, 2, 1])
    with f_col1:
        city_filter = st.selectbox("Filter City:", ["All Cities"] + sorted(list(df_timings['City'].unique())))
    with f_col2:
        search_query = st.text_input("🔍 Search Outlet Name:", placeholder="e.g. HSR, Arekere, Bhavanipuram...")
    with f_col3:
        show_diff_only = st.checkbox("Show only outlets with day variations", value=False)

    df_display = df_timings.copy()
    if city_filter != "All Cities":
        df_display = df_display[df_display['City'] == city_filter]
    if search_query:
        df_display = df_display[df_display['Outlet Name'].str.contains(search_query, case=False, na=False)]
    if show_diff_only:
        df_display = df_display[df_display.apply(lambda r: len(set([r['Monday'], r['Tuesday'], r['Wednesday'], r['Thursday'], r['Friday'], r['Saturday'], r['Sunday']])) > 1, axis=1)]

    st.info("💡 **Tip:** Edit values directly in the table below (format `HH:MM`, e.g. `08:00`, `09:00`, `10:00`, or `00:00` for 24h). Click **Save Changes** when done.")

    edited_df = st.data_editor(
        df_display,
        use_container_width=True,
        num_rows="dynamic",
        key="timings_editor"
    )

    col_btn1, col_btn2 = st.columns([1, 4])
    with col_btn1:
        if st.button("💾 Save All Changes", type="primary", use_container_width=True):
            for idx, r in edited_df.iterrows():
                b_name = r['Outlet Name']
                if b_name:
                    timings_db[b_name] = {
                        'odoo': b_name,
                        'city': r['City'],
                        'default': r['Default Opening'],
                        'Monday': r['Monday'],
                        'Tuesday': r['Tuesday'],
                        'Wednesday': r['Wednesday'],
                        'Thursday': r['Thursday'],
                        'Friday': r['Friday'],
                        'Saturday': r['Saturday'],
                        'Sunday': r['Sunday']
                    }
            engine.save_timings_db(timings_db)
            st.success("✅ Outlet timings saved successfully! Future runs will automatically use these updated hours.")
            st.rerun()

# -------------------------------------------------------------
# TAB 3: BROWSE PAST REPORTS
# -------------------------------------------------------------
elif nav_choice == "📁 Browse Past Reports":
    st.markdown('<div class="main-header">📁 Past Reports Archive</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">Browse previously generated OOS reports and screenshots.</div>', unsafe_allow_html=True)

    past_folders = [d for d in glob.glob(os.path.join(engine.OOS_ROOT, "*")) if os.path.isdir(d) and os.path.basename(d) not in ['tool', '__pycache__', '.streamlit']]
    past_folders.sort(key=os.path.getmtime, reverse=True)

    if not past_folders:
        st.info("No past folders found.")
    else:
        for fpath in past_folders:
            fname = os.path.basename(fpath)
            with st.expander(f"📁 {fname} (Modified: {datetime.datetime.fromtimestamp(os.path.getmtime(fpath)).strftime('%Y-%m-%d %H:%M')})", expanded=(fname == "Oct - 05")):
                sub_cols = st.columns([2, 1])
                with sub_cols[0]:
                    m_path = os.path.join(fpath, "Morning")
                    e_path = os.path.join(fpath, "Evening")
                    if os.path.exists(m_path):
                        st.markdown("- 🌅 **Morning Report:** Present")
                    if os.path.exists(e_path):
                        st.markdown("- 🌆 **Evening Report:** Present")
                    city_subdirs = [c for c in ['ALL CITIES', 'AP', 'BLR', 'HYD', 'MUM', 'PUN'] if os.path.isdir(os.path.join(fpath, c))]
                    if city_subdirs:
                        st.markdown(f"- 🏙️ **City Subfolders:** {', '.join(city_subdirs)}")
                    master_f = os.path.join(fpath, "Oct-2026 OOS SUMMARY.xlsx")
                    if os.path.exists(master_f):
                        st.markdown("- 📊 **Summary Workbook:** Present")
                        
                with sub_cols[1]:
                    if hasattr(os, 'startfile'):
                        if st.button(f"Open {fname} Folder", key=f"open_{fname}"):
                            os.startfile(fpath)

