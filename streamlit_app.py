import streamlit as st
import pandas as pd

# -----------------------------------------------------------------------------
# 1. PAGE CONFIGURATION
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="MOCO - Audit Control Room",
    page_icon="⛏️",
    layout="wide"
)

st.title("⛏️ MOCO - Mining Operational Audit")
st.caption("Audit Otomatis MOHH Anomaly & Latensi Input User (Khusus WORKGROUP OB & COAL)")

# -----------------------------------------------------------------------------
# 2. HELPER FUNCTIONS FOR EXCEL PARSING
# -----------------------------------------------------------------------------
def process_aturan_1_summary(file):
    """
    Aturan 1: Membaca Summary Productivity
    - Filter OB & COAL
    - Hitung MOHH = EWH + STB + BD
    - Filter Anomali MOHH > 24 jam
    - Ambil kolom: DATE, SITE, UNITNO, WORKGROUP, EWH, STB, BD, MOHH, USERNAMES DAY, USERNAMES N
    """
    try:
        df_raw = pd.read_excel(file, header=None)
        
        # Cari baris header utama
        header_row = 1
        for idx, row in df_raw.head(15).iterrows():
            row_str = [str(x).upper() for x in row.values if pd.notna(x)]
            if any("WORKGROUP" in item for item in row_str):
                header_row = idx
                break

        # Read dengan MultiIndex header (2 baris)
        df = pd.read_excel(file, header=[header_row, header_row + 1])
        
        # Flatten nama kolom
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
        
        # Mapping nama kolom standar
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
        
        # Filter Workgroup OB & COAL
        if "WORKGROUP" in df.columns:
            df["WORKGROUP"] = df["WORKGROUP"].astype(str).str.strip().str.upper()
            df = df[df["WORKGROUP"].isin(["OB", "COAL"])].copy()
            
        # Pastikan kolom numerik untuk kalkulasi
        for num_col in ["EWH", "STB", "BD", "MOHH"]:
            if num_col in df.columns:
                df[num_col] = pd.to_numeric(df[num_col], errors="coerce").fillna(0)
            else:
                df[num_col] = 0.0

        # Jika MOHH 0 atau tidak ada, hitung MOHH = EWH + STB + BD
        df["MOHH_CALC"] = df["EWH"] + df["STB"] + df["BD"]
        df["MOHH"] = df.apply(lambda r: r["MOHH_CALC"] if r["MOHH"] == 0 else r["MOHH"], axis=1)
        
        # Filter Anomali: MOHH > 24 Jam
        df_anomali = df[df["MOHH"] > 24.0].copy()
        
        # Pilih kolom sesuai Aturan 1
        target_cols = ["DATE", "SITE", "UNITNO", "WORKGROUP", "EWH", "STB", "BD", "MOHH", "USERNAMES DAY", "USERNAMES N"]
        existing_target = [c for c in target_cols if c in df_anomali.columns]
        
        return df_anomali[existing_target], df[existing_target], None
    except Exception as e:
        return None, None, str(e)


def process_aturan_2_input_time(file):
    """
    Aturan 2: Membaca Input Time (Time Entry)
    - Filter Dev (Hours) > 1 / Keterlambatan input > 1 jam
    - Ambil kolom: DATE, SHIFT, KODE UNIT, WORKGROUP, INPUT TIME, USER, DEV (HOURS TEXT), TIME START, TIME END
    """
    try:
        df = pd.read_excel(file)
        df.columns = [str(c).strip().upper() for c in df.columns]
        
        # Standardisasi pencarian kolom
        col_map = {}
        for c in df.columns:
            if "DATE" in c or "TANGGAL" in c: col_map[c] = "DATE"
            elif "SHIFT" in c: col_map[c] = "SHIFT"
            elif "UNIT" in c or "KODE" in c: col_map[c] = "KODE UNIT"
            elif "WORKGROUP" in c or "MATERIAL" in c: col_map[c] = "WORKGROUP"
            elif "START" in c: col_map[c] = "TIME START"
            elif "END" in c: col_map[c] = "TIME END"
            elif "INPUT" in c and "TIME" in c: col_map[c] = "INPUT TIME"
            elif "USER" in c or "DISPATCHER" in c: col_map[c] = "USER"
            elif "DEV" in c and ("TEXT" in c or "HOURS" in c): col_map[c] = "DEV (HOURS TEXT)"
            elif "DEV" in c: col_map[c] = "DEV_HOURS"
            
        df = df.rename(columns=col_map)
        
        # Filter Workgroup OB & COAL jika ada kolom WORKGROUP
        if "WORKGROUP" in df.columns:
            df["WORKGROUP"] = df["WORKGROUP"].astype(str).str.strip().str.upper()
            df = df[df["WORKGROUP"].isin(["OB", "COAL"])].copy()

        # Konversi Dev Hours ke Angka untuk Filter > 1 jam
        if "DEV_HOURS" in df.columns:
            df["DEV_HOURS_NUM"] = pd.to_numeric(df["DEV_HOURS"], errors="coerce").fillna(0)
        elif "DEV (HOURS TEXT)" in df.columns:
            df["DEV_HOURS_NUM"] = df["DEV (HOURS TEXT)"].astype(str).str.extract(r'(\d+)')[0].astype(float).fillna(0)
        else:
            df["DEV_HOURS_NUM"] = 0

        # Filter Anomali keterlambatan > 1 jam
        df_delay = df[df["DEV_HOURS_NUM"] > 1.0].copy()
        
        # Pilih kolom sesuai Aturan 2
        target_cols = ["DATE", "SHIFT", "KODE UNIT", "WORKGROUP", "INPUT TIME", "USER", "DEV (HOURS TEXT)", "TIME START", "TIME END"]
        existing_target = [c for c in target_cols if c in df_delay.columns]
        
        return df_delay[existing_target], df[[c for c in target_cols if c in df.columns]], None
    except Exception as e:
        return None, None, str(e)

# -----------------------------------------------------------------------------
# 3. SIDEBAR UPLOAD
# -----------------------------------------------------------------------------
st.sidebar.header("📁 Upload Operational Files")
file_prod = st.sidebar.file_uploader("1. Summary Productivity (.xlsx)", type=["xlsx", "xls"])
file_time = st.sidebar.file_uploader("2. Input Time / Time Entry (.xlsx)", type=["xlsx", "xls"])

# -----------------------------------------------------------------------------
# 4. MAIN AUDIT DISPLAY
# -----------------------------------------------------------------------------
if file_prod is not None and file_time is not None:
    df_m_anomali, df_m_all, err1 = process_aturan_1_summary(file_prod)
    df_t_delay, df_t_all, err2 = process_aturan_2_input_time(file_time)
    
    if err1:
        st.error(f"Error Aturan 1 (Summary Productivity): {err1}")
    elif err2:
        st.error(f"Error Aturan 2 (Input Time): {err2}")
    else:
        st.success("✅ File Berhasil Diproses! Menampilkan Hasil Audit Aturan 1 & Aturan 2.")
        
        # TAB DISPLAY FOR CLEAN LAYOUT
        tab1, tab2 = st.tabs(["🚨 Aturan 1: Anomali MOHH (>24 Jam)", "⏱️ Aturan 2: Keterlambatan Input User (>1 Jam)"])
        
        with tab1:
            st.subheader("🚨 Tabel Anomali MOHH (> 24 Jam)")
            st.caption("Menampilkan unit kerja OB & COAL yang total MOHH-nya melebihi 24 jam dalam 1 hari operasional.")
            
            col_m1, col_m2 = st.columns(2)
            col_m1.metric("Total Anomali MOHH Found", f"{len(df_m_anomali)} Record")
            col_m2.metric("Total Data OB & COAL Evaluated", f"{len(df_m_all)} Record")
            
            if len(df_m_anomali) > 0:
                st.dataframe(df_m_anomali, use_container_width=True)
            else:
                st.info("🎉 Tidak ditemukan anomali MOHH > 24 jam pada file ini.")

        with tab2:
            st.subheader("⏱️ Tabel User Keterlambatan Input (> 1 Jam)")
            st.caption("Menampilkan log pengetikan user/dispatcher yang melebihi 1 jam dari waktu produksi aktual (Time Start - Time End).")
            
            col_t1, col_t2 = st.columns(2)
            col_t1.metric("Total Terlambat (>1 Jam)", f"{len(df_t_delay)} Record")
            col_t2.metric("Total Time Entry Evaluated", f"{len(df_t_all)} Record")
            
            if len(df_t_delay) > 0:
                st.dataframe(df_t_delay, use_container_width=True)
            else:
                st.info("🎉 Tidak ditemukan keterlambatan input user > 1 jam pada file ini.")

else:
    st.info("👋 Silakan unggah **kedua file Excel di sidebar kiri** untuk memulai audit Aturan 1 dan Aturan 2.")
