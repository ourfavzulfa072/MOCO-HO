import streamlit as st
import pandas as pd

# -----------------------------------------------------------------------------
# 1. PAGE CONFIGURATION
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="MOCO - HO ALL SITE",
    page_icon="⛏️",
    layout="wide"
)

st.title("⛏️ MOCO - Mining Operational")
st.caption("MOHH Anomaly & Latensi Input User (OB & COAL)")

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
            df["WORKGROUP_STR"] = df["WORKGROUP"].astype(str).str.strip().str.upper()
            df = df[df["WORKGROUP_STR"].str.contains("OB|COAL|OVERBURDEN", regex=True, na=False)].copy()
            df = df.drop(columns=["WORKGROUP_STR"])
            
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
    - Otomatis mendeteksi baris header kolom (Input, user, Dev (hours), Dev (Hours text), Time Start, dll.)
    - Menyaring data yang pada kolom Dev (Hours text) mengandung kata 'jam'
    - Mengambil dan menyajikan SELURUH kolom dari baris yang terlambat tersebut (termasuk kolom Input/tanggal jam)
    """
    try:
        df_raw = pd.read_excel(file, header=None)
        
        # Deteksi baris yang berisi header tabel aktual
        header_row = 0
        for idx, row in df_raw.head(15).iterrows():
            row_str = [str(x).upper() for x in row.values if pd.notna(x)]
            if any("DEV" in item or "INPUT" in item for item in row_str):
                header_row = idx
                break

        # Baca ulang file Excel mulai dari baris header yang ditemukan
        df = pd.read_excel(file, header=header_row)

        # Rapikan nama kolom dari spasi liar/unnamed
        df.columns = [str(c).strip() for c in df.columns if not str(c).startswith("Unnamed")]

        # Identifikasi kolom Dev (Hours text)
        dev_txt_col = None
        for c in df.columns:
            if "DEV" in c.upper() and "TEXT" in c.upper():
                dev_txt_col = c
                break
        
        # Fallback jika kolom teks tidak terdeteksi eksplisit
        if not dev_txt_col:
            for c in df.columns:
                if "DEV" in c.upper():
                    dev_txt_col = c

        if not dev_txt_col:
            return None, None, "Kolom 'Dev (Hours text)' tidak ditemukan di dalam file Excel."

        # FILTER UTAMA: Ambil baris yang kolom Dev (Hours text)-nya mengandung kata 'jam'
        mask_delay = df[dev_txt_col].astype(str).str.lower().str.contains("jam", na=False)
        
        # Ambil seluruh baris yang terlambat beserta semua kolomnya (Input, user, Dev, Time Start, dst)
        df_delay = df[mask_delay].copy()

        return df_delay, df, None
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
            st.caption("Menampilkan log pengetikan user/dispatcher (termasuk tanggal & jam input) yang keterlambatannya memuat durasi jam.")
            
            col_t1, col_t2 = st.columns(2)
            col_t1.metric("Total Terlambat (>1 Jam)", f"{len(df_t_delay)} Record")
            col_t2.metric("Total Time Entry Evaluated", f"{len(df_t_all)} Record")
            
            if len(df_t_delay) > 0:
                st.dataframe(df_t_delay, use_container_width=True)
            else:
                st.info("🎉 Tidak ditemukan keterlambatan input user > 1 jam pada file ini.")

else:
    st.info("👋 Silakan unggah **kedua file Excel di sidebar kiri** untuk memulai audit Aturan 1 dan Aturan 2.")
