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
**Single-Site Audit Session:** Unggah 2 file ERP harian (*Summary Productivity* & *Input Time*) untuk audit otomatis integritas jam kerja unit (MOHH, HM/EWH) dan latensi pengetikan dispatcher.
""")

# -----------------------------------------------------------------------------
# 2. HELPER FUNCTIONS FOR EXCEL PARSING
# -----------------------------------------------------------------------------
def load_summary_productivity(file):
    """
    Membaca & membersihkan file Summary Productivity.
    Menggabungkan multi-header dan memfilter WORKGROUP hanya OB & COAL.
    """
    try:
        # Read raw excel without initial header to locate header rows
        df_raw = pd.read_excel(file, header=None)
        
        # Cari baris yang mengandung 'WORKGROUP' atau 'MODEL'
        header_row_idx = None
        for idx, row in df_raw.head(15).iterrows():
            row_str = row.astype(str).str.upper().tolist()
            if any("WORKGROUP" in item for item in row_str):
                header_row_idx = idx
                break
        
        if header_row_idx is None:
            header_row_idx = 1  # Fallback ke baris ke-2
            
        # Re-read dengan header bertingkat
        df = pd.read_excel(file, header=[header_row_idx, header_row_idx + 1])
        
        # Cleaning column names (flatten MultiIndex tuple)
        flat_cols = []
        for col in df.columns:
            l1 = str(col[0]).strip() if not str(col[0]).startswith("Unnamed") else ""
            l2 = str(col[1]).strip() if not str(col[1]).startswith("Unnamed") else ""
            if l1 and l2:
                flat_cols.append(f"{l1}_{l2}".upper())
            elif l1:
                flat_cols.append(l1.upper())
            elif l2:
                flat_cols.append(l2.upper())
            else:
                flat_cols.append("UNKNOWN")
        
        df.columns = flat_cols
        
        # Identifikasi Kolom WORKGROUP
        wg_col = [c for c in df.columns if "WORKGROUP" in c]
        if wg_col:
            df = df.rename(columns={wg_col[0]: "WORKGROUP"})
        else:
            # Dropdown/Fallback pencarian kolom
            df.columns.values[0] = "WORKGROUP"
            
        # Standardisasi data WORKGROUP
        df['WORKGROUP'] = df['WORKGROUP'].astype(str).str.strip().str.upper()
        
        # FILTER KETAT: Hanya WORKGROUP 'OB' dan 'COAL'
        df_filtered = df[df['WORKGROUP'].isin(['OB', 'COAL'])].copy()
        
        return df_filtered, None
    except Exception as e:
        return None, str(e)

def load_input_time(file):
    """
    Membaca & membersihkan file Input Time (Time Entry).
    Mengakses timestamp entri data & dispatcher/user.
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
    help="Upload file Summary Productivity yang memuat MOHH, HM, EWH, BD, STB, dan Workgroup."
)

file_input_time = st.sidebar.file_uploader(
    "2. Input Time / Time Entry (.xlsx)", 
    type=["xlsx", "xls"],
    help="Upload file Time Entry yang memuat jam entri data dan nama User/Dispatcher."
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
        total_time_entries = len(df_time)
        
        with col1:
            st.metric("Total Records Filtered", f"{total_records:,} Units")
        with col2:
            st.metric("OB Units", f"{total_ob:,}")
        with col3:
            st.metric("COAL Units", f"{total_coal:,}")
        with col4:
            st.metric("Total Time Entries", f"{total_time_entries:,}")
            
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
        
        # Jika terdapat kolom USER / DISPATCHER di file Time Entry
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
    # State saat file belum diunggah
    st.info("👋 **Selamat Datang di MOCO Dashboard!**")
    st.warning(" Silakan **unggah 2 file Excel harian** di sidebar sebelah kiri untuk mulai membaca data operasional site.")
    
    st.markdown("""
    ### 📋 Petunjuk Unggah File:
    1. **Summary Productivity (.xlsx):**
       * File yang berisi *grup HM* (`HM D`, `HM N`, `HM TTL`, `EWH`, `STB`, `BD`, `MOHH`).
       * Sistem akan memfilter material secara otomatis, **hanya mengambil `OB` dan `COAL`**.
    2. **Input Time / Time Entry (.xlsx):**
       * File log aktivitas input waktu dari Dispatcher/User.
    """)
