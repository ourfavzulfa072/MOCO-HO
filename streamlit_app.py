import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# -----------------------------------------------------------------------------
# 1. PAGE CONFIGURATION
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="MOCO - Control Room Mining Dashboard",
    page_icon="🚜",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("🚜 MOCO - Mining Operational Control Room")
st.markdown("""
**Single-Site Audit Session:** Unggah 2 file ERP harian (*Summary Productivity* & *Input Time*) untuk audit otomatis integritas Unit No, jam kerja unit (MOHH, HM/EWH), dan latensi pengetikan dispatcher.
""")

# -----------------------------------------------------------------------------
# 2. HELPER FUNCTIONS FOR EXCEL PARSING
# -----------------------------------------------------------------------------
def load_summary_productivity(file):
    """
    Membaca & membersihkan file Summary Productivity secara aman dari NaN/float.
    Mendukung deteksi UNIT NO, WORKGROUP (OB & COAL), dan grup jam kerja (HM, EWH, STB, BD, MOHH).
    """
    try:
        # 1. Read raw excel
        df_raw = pd.read_excel(file, header=None)
        
        # 2. Cari baris header secara aman (mencari kata 'WORKGROUP' atau 'UNIT' atau 'MODEL')
        header_row_idx = None
        for idx, row in df_raw.head(15).iterrows():
            row_vals = [str(x).upper() for x in row.values if pd.notna(x)]
            if any("WORKGROUP" in item or "MODEL" in item for item in row_vals):
                header_row_idx = idx
                break
        
        if header_row_idx is None:
            header_row_idx = 1  # Fallback ke baris ke-2
            
        # 3. Read dengan header bertingkat
        df = pd.read_excel(file, header=[header_row_idx, header_row_idx + 1])
        
        # 4. Flatten MultiIndex columns secara aman
        flat_cols = []
        for col in df.columns:
            l1 = str(col[0]).strip() if pd.notna(col[0]) and not str(col[0]).startswith("Unnamed") else ""
            l2 = str(col[1]).strip() if pd.notna(col[1]) and not str(col[1]).startswith("Unnamed") else ""
            
            if l1 and l2:
                flat_cols.append(f"{l1}_{l2}".upper())
            elif l1:
                flat_cols.append(l1.upper())
            elif l2:
                flat_cols.append(l2.upper())
            else:
                flat_cols.append(f"COL_{len(flat_cols)}")
        
        df.columns = flat_cols
        
        # 5. Cari dan standardisasi nama kolom penting (WORKGROUP & UNIT NO)
        wg_cols = [c for c in df.columns if "WORKGROUP" in c]
        if wg_cols:
            df = df.rename(columns={wg_cols[0]: "WORKGROUP"})
        else:
            df.rename(columns={df.columns[0]: "WORKGROUP"}, inplace=True)

        unit_cols = [c for c in df.columns if "UNIT" in c or "CN" in c or "NO" in c or "EQ" in c]
        if unit_cols:
            df = df.rename(columns={unit_cols[0]: "UNIT_NO"})

        # Clean & Filter WORKGROUP (Hanya OB dan COAL)
        df['WORKGROUP'] = df['WORKGROUP'].astype(str).str.strip().str.upper()
        df_filtered = df[df['WORKGROUP'].isin(['OB', 'COAL'])].copy()
        
        # Bersihkan string UNIT_NO jika ada
        if "UNIT_NO" in df_filtered.columns:
            df_filtered['UNIT_NO'] = df_filtered['UNIT_NO'].astype(str).str.strip()
            
        return df_filtered, None
    except Exception as e:
        return None, str(e)

def load_input_time(file):
    """
    Membaca & membersihkan file Input Time (Time Entry).
    """
    try:
        df = pd.read_excel(file)
        # Flatten string columns
        df.columns = [str(c).strip().upper() for c in df.columns]
        return df, None
    except Exception as e:
        return None, str(e)

# -----------------------------------------------------------------------------
# 3. SIDEBAR UPLOAD CONTROL PANEL
# -----------------------------------------------------------------------------
st.sidebar.header("📁 Upload Operational Files")
file_summary = st.sidebar.file_uploader(
    "1. Summary Productivity (.xlsx)", 
    type=["xlsx", "xls"],
    help="Upload file Summary Productivity (memuat Unit No, MOHH, HM, EWH, BD, STB, dan Workgroup)."
)

file_input_time = st.sidebar.file_uploader(
    "2. Input Time / Time Entry (.xlsx)", 
    type=["xlsx", "xls"],
    help="Upload file Time Entry (memuat jam entri data dan nama User/Dispatcher)."
)

# -----------------------------------------------------------------------------
# 4. DASHBOARD PROCESSING LOGIC
# -----------------------------------------------------------------------------
if file_summary is not None and file_input_time is not None:
    df_prod, err_prod = load_summary_productivity(file_summary)
    df_time, err_time = load_input_time(file_input_time)
    
    if err_prod:
        st.error(f"Gagal membaca file Summary Productivity: {err_prod}")
    elif err_time:
        st.error(f"Gagal membaca file Input Time: {err_time}")
    else:
        st.success("✅ File berhasil diunggah & difilter! Menampilkan analisa site operasional.")
        
        # ---------------------------------------------------------
        # A. EXECUTIVE SUMMARY CARDS
        # ---------------------------------------------------------
        st.subheader("📌 Executive Operational Overview (OB & COAL Only)")
        
        col1, col2, col3, col4 = st.columns(4)
        
        total_records = len(df_prod)
        total_ob = len(df_prod[df_prod['WORKGROUP'] == 'OB'])
        total_coal = len(df_prod[df_prod['WORKGROUP'] == 'COAL'])
        total_units = df_prod['UNIT_NO'].nunique() if 'UNIT_NO' in df_prod.columns else total_records
        
        with col1:
            st.metric("Total Active Units", f"{total_units:,} Units")
        with col2:
            st.metric("OB Units", f"{total_ob:,}")
        with col3:
            st.metric("COAL Units", f"{total_coal:,}")
        with col4:
            st.metric("Total Records Filtered", f"{total_records:,}")
            
        st.markdown("---")
        
        # ---------------------------------------------------------
        # B. DATA PREVIEW & WORKGROUP DISTRIBUTION
        # ---------------------------------------------------------
        left_col, right_col = st.columns([2, 1])
        
        with left_col:
            st.subheader("📊 Summary Productivity Data (OB & COAL)")
            st.dataframe(df_prod.head(100), use_container_width=True)
            
        with right_col:
            st.subheader("🥧 Workgroup Distribution")
            wg_counts = df_prod['WORKGROUP'].value_counts().reset_index()
            wg_counts.columns = ['WORKGROUP', 'COUNT']
            fig_pie = px.pie(
                wg_counts, 
                values='COUNT', 
                names='WORKGROUP', 
                color='WORKGROUP',
                color_discrete_map={'OB': '#FFA500', 'COAL': '#333333'},
                hole=0.4
            )
            st.plotly_chart(fig_pie, use_container_width=True)

        st.markdown("---")
        
        # ---------------------------------------------------------
        # C. TIME ENTRY & DISPATCHER LATENCY AUDIT
        # ---------------------------------------------------------
        st.subheader("⏱️ Dispatcher Input Latency & User Audit")
        
        st.dataframe(df_time.head(100), use_container_width=True)
        
        # Identifikasi kolom User / Dispatcher di file Time Entry
        user_cols = [c for c in df_time.columns if "USER" in c or "DISPATCHER" in c or "CREATED" in c]
        if user_cols:
            user_col_name = user_cols[0]
            st.subheader("🏆 Leaderboard Inputter / Dispatcher Activity")
            user_summary = df_time[user_col_name].value_counts().reset_index()
            user_summary.columns = ['User / Dispatcher', 'Total Input Entries']
            
            fig_user = px.bar(
                user_summary.head(15),
                x='Total Input Entries',
                y='User / Dispatcher',
                orientation='h',
                title="Top 15 Most Active Inputters",
                color='Total Input Entries',
                color_continuous_scale='Blues'
            )
            fig_user.update_layout(yaxis={'categoryorder': 'total ascending'})
            st.plotly_chart(fig_user, use_container_width=True)

else:
    # State awal sebelum upload
    st.info("👋 **Selamat Datang di MOCO Dashboard!**")
    st.warning(" Silakan **unggah 2 file Excel harian** di sidebar sebelah kiri untuk mulai membaca data operasional site.")
    
    st.markdown("""
    ### 📋 Petunjuk Unggah File:
    1. **Summary Productivity (.xlsx):**
       * Menampilkan **Unit No**, **Workgroup** (`OB` dan `COAL`), serta jam operasional (`HM`, `EWH`, `STB`, `BD`, `MOHH`).
    2. **Input Time / Time Entry (.xlsx):**
       * Menampilkan log aktivitas dan ketepatan waktu input dari Dispatcher/User.
    """)
