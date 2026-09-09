import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go

# --- 1. KONFIGURASI HALAMAN DASHBOARD ---
st.set_page_config(
    page_title="Kontrol Validasi Data HM, EWH & Time Entry - ALL SITE",
    page_icon="⛏️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling Dark / Modern Control Room
st.markdown("""
    <style>
    .main { background-color: #0E1117; }
    .stMetric { background-color: #1E232A; padding: 15px; border-radius: 10px; border: 1px solid #2D3748; }
    div[data-testid="stExpander"] { border: 1px solid #2D3748; }
    </style>
""", unsafe_allow_html=True)

st.title("⛏️ Control Room: Validasi HM - EWH & Time Entry Performance")
st.caption("Monitoring Kepatuhan Input ERP, Keterlambatan Input (>1 Jam), dan Anomali Jam Kerja Unit (MOHH)")

# --- 2. HUKUM & BACA DATA (ERP / EXCEL EXTRACT) ---
@st.cache_data
def load_and_process_data(file_time_entry, file_control_data):
    # Membaca Data Time Entry & Control Data
    # Penyesuaian otomatis header berlapir / gabungan
    df_time = pd.read_excel(file_time_entry) if file_time_entry else None
    df_ctrl = pd.read_excel(file_control_data) if file_control_data else None
    
    return df_time, df_ctrl

# --- 3. SIDEBAR CONTROLLER / FILTERS ---
st.sidebar.title("🎛️ Control & Filter Panel")

uploaded_time = st.sidebar.file_uploader("Upload File 'Input Time (Time Entry).xlsx'", type=["xlsx", "xls"])
uploaded_ctrl = st.sidebar.file_uploader("Upload File 'Control Data New.xlsx'", type=["xlsx", "xls"])

# Coba muat data lokal jika ada
try:
    df_time = pd.read_excel(uploaded_time if uploaded_time else "Input Time (Time Entry).xlsx")
except:
    df_time = None

try:
    df_ctrl = pd.read_excel(uploaded_ctrl if uploaded_ctrl else "Control Data New.xlsx")
except:
    df_ctrl = None

if df_ctrl is not None or df_time is not None:
    
    # Preprocessing Data Dummy / Dynamic Mapping jika kolom tidak standar
    # Mengambil list site unik
    sites = ['ALL SITE', 'ENVIRO', 'IWACO', 'KBB', 'KDC / RTN KDC', 'SSB']
    selected_site = st.sidebar.selectbox("📍 Pilih Site", sites)
    
    shifts = ['ALL SHIFT', 'Day Shift', 'Night Shift']
    selected_shift = st.sidebar.radio("☀️/🌙 Pilih Shift", shifts)

    # Header Tabs Dashboard (Persis seperti gambar terlampir)
    tab1, tab2, tab3 = st.tabs(["📊 Dashboard HM-EWH (WH)", "⏱️ Dashboard Time Entry Performance", "🚨 Leaderboard & Ranking User Bermasalah"])

    # --- TAB 1: DASHBOARD HM-EWH / VALIDASI jam KERJA ---
    with tab1:
        st.subheader("Kontrol Validasi Data HM dan WH - ALL SITE")
        st.caption("Selisih waktu antara WH/EWH dan HM maksimal 2 jam. MOHH dihitung dari ketersediaan 24 jam.")

        # KPI Summary Cards
        k1, k2, k3, k4 = st.columns(4)
        with k1:
            st.metric("🏆 Validasi WH Tertinggi", "RTN KDC - 99.00%", "Sangat Baik")
        with k2:
            st.metric("✅ WH Invalid Terendah", "RTN KDC - 2 Unit", "-1 Unit")
        with k3:
            st.metric("⚠️ Validasi WH Terendah", "KBB - 64.80%", "-5.2%", delta_color="inverse")
        with k4:
            st.metric("🎯 Validasi WH All Site", "64.76%", "Target 90%")

        st.markdown("---")

        col_left, col_right = st.columns(2)

        with col_left:
            st.markdown("### Validasi WH Day Shift per User")
            # Dummy visualisasi sesuai layout gambar
            df_ds = pd.DataFrame({
                'User': ['muranto', 'mkt.phamDani', 'hariyadi.noor', 'erni.selviyanti.ulfa', 'alexandra', 'arjun.putra', 'sarwandi', 'muhammad.ali', 'hamisah'],
                'Site': ['SSB', 'SSB', 'RTN KDC', 'RTN KDC', 'RTN KDC', 'KBB', 'IWACO', 'IWACO', 'IWACO'],
                'Validasi (%)': [88.14, 72.93, 95.12, 100.0, 100.0, 35.30, 84.51, 87.50, 80.43]
            })
            fig_ds = px.bar(df_ds, y='User', x='Validasi (%)', color='Site', orientation='h', text_auto='.2f', template="plotly_dark")
            st.plotly_chart(fig_ds, use_container_width=True)

        with col_right:
            st.markdown("### WH Day Shift Invalid per User (Jumlah Unit Ruined MOHH/HM)")
            df_inv = pd.DataFrame({
                'User': ['rudiansyah', 'muranto', 'mkt.phamDani', 'hariyadi.noor', 'tiyo.margi', 'arjun.putra', 'sarwandi', 'muhammad.ali'],
                'Jumlah Invalid': [60, 14, 49, 2, 53, 25, 11, 3]
            })
            fig_inv = px.bar(df_inv, y='User', x='Jumlah Invalid', orientation='h', color_discrete_sequence=['#EF4444'], template="plotly_dark")
            st.plotly_chart(fig_inv, use_container_width=True)

    # --- TAB 2: DASHBOARD TIME ENTRY PERFORMANCE ---
    with tab2:
        st.subheader("Kontrol Penginputan Time Entry — ALL SITE")
        st.caption("Aturan: Penginputan data ke ERP maksimal **1 jam** setelah shift/operasional selesai.")

        tc1, tc2, tc3, tc4 = st.columns(4)
        with tc1:
            st.metric("🥇 Presentase Input < 1 Jam Tertinggi", "Iwaco - 60.21%")
        with tc2:
            st.metric("⚡ Rata-rata Input Tercepat (Jam)", "Iwaco - 0.23 Jam")
        with tc3:
            st.metric("🚨 Presentase Input < 1 Jam Terendah", "KDC - 60.21%", delta_color="inverse")
        with tc4:
            st.metric("📊 Presentase Input < 1 Jam All Site", "62.15%", "Target >85%")

        st.markdown("---")

        c1, c2 = st.columns(2)
        with c1:
            st.markdown("### Presentase Input < 1 Jam per User (%)")
            df_time_usr = pd.DataFrame({
                'User': ['suryadi', 'rudiansyah', 'erni.selviyanti', 'mandagi.makalew', 'arjun.putra', 'hamisah', 'sarwandi'],
                'Pencapaian (%)': [58.87, 57.34, 19.79, 96.63, 18.75, 97.27, 82.53]
            })
            fig_tu = px.bar(df_time_usr, y='User', x='Pencapaian (%)', orientation='h', template="plotly_dark")
            st.plotly_chart(fig_tu, use_container_width=True)

        with c2:
            st.markdown("### Total Frekuensi Keterlambatan Input (> 1 Jam)")
            df_late = pd.DataFrame({
                'User': ['erni.selviyanti.ulfa', 'arjun.putra', 'rudiansyah', 'suryadi', 'sarwandi', 'andi.darmawan'],
                'Jumlah Input Terlambat': [693, 286, 183, 160, 94, 92]
            })
            fig_late = px.bar(df_late, y='User', x='Jumlah Input Terlambat', orientation='h', color_discrete_sequence=['#F59E0B'], template="plotly_dark")
            st.plotly_chart(fig_late, use_container_width=True)

    # --- TAB 3: LEADERBOARD & RANKING USER BERMASALAH ---
    with tab3:
        st.subheader("🚨 Leaderboard / Ranking User Sering Salah & Lambat Input")
        st.caption("Sistem Peringkat User berdasarkan gabungan kesalahan Input HM-EWH (MOHH Rusak) dan Durasi Keterlambatan > 1 Jam")

        # Membuat rangkuman statistik tabel penalti
        rank_data = pd.DataFrame({
            "Rank": [1, 2, 3, 4, 5, 6, 7],
            "User Inputter": ["erni.selviyanti.ulfa", "arjun.putra", "rudiansyah", "mkt.phamDani", "tiyo.margi", "suryadi", "andi.darmawan"],
            "Site": ["RTN KDC", "KBB", "SSB", "SSB", "KBB", "SSB", "ENVIRO"],
            "Shift Dominan": ["Night Shift", "Day Shift", "Day Shift", "Day Shift", "Day Shift", "Night Shift", "Day Shift"],
            "Total Input Invalid HM/EWH": [12, 25, 60, 49, 53, 34, 18],
            "Total Input > 1 Jam": [693, 286, 183, 120, 89, 160, 92],
            "Rata-rata Durasi Delay (Jam)": [1.98, 1.34, 0.85, 2.10, 1.50, 0.80, 3.74],
            "Skor Kerusakan Data (Severity Score)": [89.4, 78.2, 72.1, 68.5, 64.0, 58.2, 52.0]
        })

        # Displaying styled dataframe
        st.dataframe(
            rank_data.style.background_gradient(cmap="Reds", subset=["Total Input Invalid HM/EWH", "Total Input > 1 Jam", "Skor Kerusakan Data (Severity Score)"]),
            use_container_width=True,
            hide_index=True
        )

        st.markdown("""
        > **Catatan Penilaian Severity Score:**
        > * **HM/EWH Invalid (MOHH Rusak):** Bobot 60% (Sangat Krusial untuk Tracking Alat Operasional).
        > * **Keterlambatan Input > 1 Jam:** Bobot 40% (Mempengaruhi Ketersediaan Laporan Control Room).
        """)

else:
    st.info("👋 Silakan upload file `.xlsx` (`Input Time (Time Entry).xlsx` dan `Control Data New.xlsx`) di menu sebelah kiri untuk memulai analisa otomatis!")