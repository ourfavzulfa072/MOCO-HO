import streamlit as st
import pandas as pd
import plotly.express as px
import io

# -----------------------------------------------------------------------------
# 1. PAGE CONFIGURATION
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="BIMA NUSA INT MOCO - Audit",
    page_icon="⛏️",
    layout="wide"
)

# Custom CSS untuk styling header & card
st.markdown("""
    <style>
    .main-header {
        font-size: 2.2rem;
        font-weight: 700;
        color: #F3F4F6;
        margin-bottom: 0px;
    }
    .sub-header {
        font-size: 1rem;
        color: #9CA3AF;
        margin-bottom: 20px;
    }
    </style>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# 2. HELPER FUNCTIONS FOR EXCEL PARSING
# -----------------------------------------------------------------------------
def process_aturan_1_summary(file):
    try:
        df_raw = pd.read_excel(file, header=None)
        
        header_row = 1
        for idx, row in df_raw.head(15).iterrows():
            row_str = [str(x).upper() for x in row.values if pd.notna(x)]
            if any("WORKGROUP" in item for item in row_str):
                header_row = idx
                break

        df = pd.read_excel(file, header=[header_row, header_row + 1])
        
        cols = []
        for c in df.columns:
            top = str(c[0]).strip() if pd.notna(c[0]) and not str(c[0]).startswith("Unnamed") else ""
            bot = str(c[1]).strip() if pd.notna(c[1]) and not str(c[1]).startswith("Unnamed") else ""
            if top and bot:
                cols.append(f"{top}_{bot}".upper())
            elif top:
                cols.append(top.upper())
            elif bot:
                cols.append(bot.upper())
            else:
                cols.append(f"COL_{len(cols)}")
        df.columns = cols
        
        col_map = {}
        for c in df.columns:
            if "WORKGROUP" in c: col_map[c] = "WORKGROUP"
            elif "SITE" in c: col_map[c] = "SITE"
            elif "UNIT" in c or "CN" in c or "EQUIPMENT" in c: col_map[c] = "UNITNO"
            elif "DATE" in c or "TANGGAL" in c: col_map[c] = "DATE"
            elif c.endswith("_EWH") or c == "EWH": col_map[c] = "EWH"
            elif c.endswith("_STB") or c == "STB": col_map[c] = "STB"
            elif c.endswith("_BD") or c == "BD": col_map[c] = "BD"
            elif c.endswith("_MOHH") or c == "MOHH": col_map[c] = "MOHH"
            elif "USER" in c and ("DAY" in c or "_D" in c): col_map[c] = "USERNAMES DAY"
            elif "USER" in c and ("NIGHT" in c or "_N" in c): col_map[c] = "USERNAMES N"

        df = df.rename(columns=col_map)
        
        if "DATE" in df.columns:
            df["DATE"] = pd.to_datetime(df["DATE"], errors="coerce").dt.strftime("%Y-%m-%d")

        if "WORKGROUP" in df.columns:
            df["WORKGROUP_STR"] = df["WORKGROUP"].astype(str).str.strip().str.upper()
            df = df[df["WORKGROUP_STR"].str.contains("OB|COAL|OVERBURDEN", regex=True, na=False)].copy()
            df = df.drop(columns=["WORKGROUP_STR"])
            
        for num_col in ["EWH", "STB", "BD", "MOHH"]:
            if num_col in df.columns:
                df[num_col] = pd.to_numeric(df[num_col], errors="coerce").fillna(0)
            else:
                df[num_col] = 0.0

        df["MOHH_CALC"] = (df["EWH"] + df["STB"] + df["BD"]).round(2)
        df["MOHH"] = df.apply(lambda r: r["MOHH_CALC"] if r["MOHH"] == 0 else round(r["MOHH"], 2), axis=1)
        
        df_anomali = df[df["MOHH"] > 24.001].copy()
        
        target_cols = ["DATE", "SITE", "UNITNO", "WORKGROUP", "EWH", "STB", "BD", "MOHH", "USERNAMES DAY", "USERNAMES N"]
        existing_target = [c for c in target_cols if c in df_anomali.columns]
        
        return df_anomali[existing_target], df[existing_target], None
    except Exception as e:
        return None, None, str(e)


def process_aturan_2_input_time(file):
    try:
        df_raw = pd.read_excel(file, header=None)
        
        sub_header_row = None
        for idx, row in df_raw.head(15).iterrows():
            row_str = [str(x).upper() for x in row.values if pd.notna(x)]
            if any("DEV (HOURS TEXT)" in item or "DEV (HOURS)" in item for item in row_str):
                sub_header_row = idx
                break

        if sub_header_row is not None:
            df = pd.read_excel(file, header=sub_header_row)
        else:
            df = pd.read_excel(file)

        df.columns = [str(c).strip() for c in df.columns]

        date_col = None
        for c in df.columns:
            if "DATE" in str(c).upper() or "TANGGAL" in str(c).upper():
                date_col = c
                break
        
        if date_col:
            df[date_col] = pd.to_datetime(df[date_col], errors="coerce").dt.strftime("%Y-%m-%d")

        dev_txt_col = None
        for c in df.columns:
            c_upper = str(c).upper()
            if "DEV" in c_upper and "TEXT" in c_upper:
                dev_txt_col = c
                break

        if not dev_txt_col:
            dev_cols = [c for c in df.columns if "DEV" in str(c).upper()]
            if len(dev_cols) >= 2:
                dev_txt_col = dev_cols[1]
            elif len(dev_cols) == 1:
                dev_txt_col = dev_cols[0]

        if not dev_txt_col:
            return None, None, "Kolom 'Dev (Hours text)' tidak ditemukan di dalam file Excel."

        mask_delay = df[dev_txt_col].astype(str).str.lower().str.contains("jam", na=False)
        df_delay = df[mask_delay].copy()

        return df_delay, df, None
    except Exception as e:
        return None, None, str(e)

# Helper untuk Export Excel
def convert_df_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Audit Report')
    processed_data = output.getvalue()
    return processed_data

# -----------------------------------------------------------------------------
# 3. SIDEBAR BRANDING & UPLOAD
# -----------------------------------------------------------------------------
with st.sidebar:
    # Contoh tempat logo (Ganti URL dengan file logo Bima Nusa/MOCO jika ada)
    st.image("https://img.icons8.com/color/96/mine-cart.png", width=70)
    st.title("Control Panel")
    st.caption("BIMA NUSA INT - MOCO Systems")
    st.markdown("---")
    
    st.subheader("📁 Submit Daily Operational Files")
    file_prod = st.file_uploader("Summary Productivity (.xlsx)", type=["xlsx", "xls"])
    file_time = st.file_uploader("Input Time / Time Entry (.xlsx)", type=["xlsx", "xls"])

# -----------------------------------------------------------------------------
# 4. MAIN AUDIT DISPLAY & VISUALIZATION
# -----------------------------------------------------------------------------
# Header Utama
col_logo, col_title = st.columns([1, 8])
with col_title:
    st.markdown('<p class="main-header">BIMA NUSA INT MOCO - Mining Operational</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">MOHH Anomaly & Latensi Input User (OB & COAL)</p>', unsafe_allow_html=True)

if file_prod is not None and file_time is not None:
    df_m_anomali, df_m_all, err1 = process_aturan_1_summary(file_prod)
    df_t_delay, df_t_all, err2 = process_aturan_2_input_time(file_time)
    
    if err1:
        st.error(f"Error Aturan 1 (Summary Productivity): {err1}")
    elif err2:
        st.error(f"Error Aturan 2 (Input Time): {err2}")
    else:
        st.success("✅ File Berhasil Diproses! Menampilkan Dashboard & Hasil Audit.")
        
        tab1, tab2 = st.tabs(["🚨 Aturan 1: Anomali MOHH (>24 Jam)", "⏱️ Aturan 2: Keterlambatan Input User (>1 Jam)"])
        
        # ==========================================
        # TAB 1: MOHH ANOMALY
        # ==========================================
        with tab1:
            st.subheader("🚨 Dashboard & Tabel Anomali MOHH (> 24 Jam)")
            
            # Metrics Row
            col_m1, col_m2, col_m3 = st.columns(3)
            col_m1.metric("Total Anomali MOHH Found", f"{len(df_m_anomali)} Record")
            col_m2.metric("Total Evaluated Units", f"{len(df_m_all)} Record")
            percent_anomali = (len(df_m_anomali) / len(df_m_all) * 100) if len(df_m_all) > 0 else 0
            col_m3.metric("% Anomali", f"{percent_anomali:.1f}%")
            
            st.markdown("---")
            
            # VISUALISASI ATURAN 1
            if len(df_m_all) > 0:
                col_chart1, col_chart2 = st.columns(2)
                
                with col_chart1:
                    st.markdown("##### 📊 Komposisi Rata-Rata Jam Operational (Total Data)")
                    avg_hours = pd.DataFrame({
                        'Kategori': ['EWH', 'STB', 'BD'],
                        'Jam': [df_m_all['EWH'].mean(), df_m_all['STB'].mean(), df_m_all['BD'].mean()]
                    })
                    fig_pie = px.pie(avg_hours, values='Jam', names='Kategori', hole=0.4,
                                     color_discrete_sequence=['#22c55e', '#eab308', '#ef4444'])
                    fig_pie.update_layout(margin=dict(t=20, b=20, l=20, r=20), height=300)
                    st.plotly_chart(fig_pie, use_container_width=True)
                
                with col_chart2:
                    st.markdown("##### 🏗️ Unit Anomali per Site")
                    if len(df_m_anomali) > 0 and 'SITE' in df_m_anomali.columns:
                        site_counts = df_m_anomali['SITE'].value_counts().reset_index()
                        site_counts.columns = ['SITE', 'Jumlah Anomali']
                        fig_bar = px.bar(site_counts, x='SITE', y='Jumlah Anomali', color='Jumlah Anomali',
                                         color_continuous_scale='Reds')
                        fig_bar.update_layout(margin=dict(t=20, b=20, l=20, r=20), height=300)
                        st.plotly_chart(fig_bar, use_container_width=True)
                    else:
                        st.info("Tidak ada grafik site karena 0 anomali ditemukan.")

            st.markdown("### 📋 Detail Data Anomali")
            if len(df_m_anomali) > 0:
                st.dataframe(df_m_anomali, use_container_width=True)
                excel_data1 = convert_df_to_excel(df_m_anomali)
                st.download_button(
                    label="📥 Download Data Anomali MOHH (.xlsx)",
                    data=excel_data1,
                    file_name="Anomali_MOHH_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.info("🎉 Tidak ditemukan anomali MOHH > 24 jam pada file ini.")

        # ==========================================
        # TAB 2: INPUT TIME LATENCY
        # ==========================================
        with tab2:
            st.subheader("⏱️ Dashboard & Tabel User Keterlambatan Input (> 1 Jam)")
            
            col_t1, col_t2, col_t3 = st.columns(3)
            col_t1.metric("Total Terlambat (>1 Jam)", f"{len(df_t_delay)} Record")
            col_t2.metric("Total Entry Evaluated", f"{len(df_t_all)} Record")
            percent_delay = (len(df_t_delay) / len(df_t_all) * 100) if len(df_t_all) > 0 else 0
            col_t3.metric("% Latensi Input", f"{percent_delay:.1f}%")

            st.markdown("---")
            
            # VISUALISASI ATURAN 2
            if len(df_t_delay) > 0:
                col_chart3, col_chart4 = st.columns(2)
                
                with col_chart3:
                    st.markdown("##### 👤 Top 5 User Terlambat Input")
                    # Cari kolom user secara fleksibel
                    user_col = [c for c in df_t_delay.columns if "USER" in str(c).upper()]
                    user_col_name = user_col[0] if user_col else None
                    
                    if user_col_name:
                        top_users = df_t_delay[user_col_name].value_counts().head(5).reset_index()
                        top_users.columns = ['User', 'Frekuensi']
                        fig_user = px.bar(top_users, y='User', x='Frekuensi', orientation='h',
                                          color='Frekuensi', color_continuous_scale='Oranges')
                        fig_user.update_layout(yaxis={'categoryorder':'total ascending'}, margin=dict(t=20, b=20, l=20, r=20), height=300)
                        st.plotly_chart(fig_user, use_container_width=True)
                
                with col_chart4:
                    st.markdown("##### ☀️ vs 🌙 Keterlambatan Berdasarkan Shift")
                    shift_col = [c for c in df_t_delay.columns if "SHIFT" in str(c).upper()]
                    if shift_col:
                        shift_counts = df_t_delay[shift_col[0]].value_counts().reset_index()
                        shift_counts.columns = ['Shift', 'Jumlah']
                        fig_shift = px.pie(shift_counts, values='Jumlah', names='Shift', color_discrete_sequence=px.colors.qualitative.Set2)
                        fig_shift.update_layout(margin=dict(t=20, b=20, l=20, r=20), height=300)
                        st.plotly_chart(fig_shift, use_container_width=True)

            st.markdown("### 📋 Detail Log Keterlambatan User")
            if len(df_t_delay) > 0:
                st.dataframe(df_t_delay, use_container_width=True)
                excel_data2 = convert_df_to_excel(df_t_delay)
                st.download_button(
                    label="📥 Download Data Latensi Input (.xlsx)",
                    data=excel_data2,
                    file_name="Latensi_Input_User_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.info("🎉 Tidak ditemukan keterlambatan input user > 1 jam pada file ini.")

else:
    # Tampilan saat belum ada file yang diunggah
    st.info("👋 Silakan unggah **kedua file Excel di sidebar kiri** untuk memulai audit Aturan 1 dan Aturan 2.")
    
    # Placeholder visual / Info Card agar ruang kosong tetap menarik
    st.markdown("---")
    col_info1, col_info2 = st.columns(2)
    with col_info1:
        st.markdown("""
        ### 🚨 Aturan 1: Audit MOHH Anomaly
        * **Syarat Target:** Workgroup **OB** dan **COAL**.
        * **Kalkulasi:** Total MOHH = $\text{EWH} + \text{STB} + \text{BD}$.
        * **Kriteria Anomali:** Total MOHH $> 24.0$ jam.
        """)
    with col_info2:
        st.markdown("""
        ### ⏱️ Aturan 2: Latensi Input User
        * **Syarat Target:** Log pengetikan user/dispatcher.
        * **Deteksi:** Membaca kolom `Dev (Hours text)`.
        * **Kriteria Anomali:** Memuat teks berdurasi jam (terlambat $> 1$ jam).
        """)
