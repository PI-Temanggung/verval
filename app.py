import io
import openpyxl
import pandas as pd
import requests
import streamlit as st

# Konfigurasi Halaman (Lebar responsif untuk PC dan HP)
st.set_page_config(
    page_title="Panel Verifikasi Kariyawan Tiri", page_icon="📱", layout="wide"
)

# 1. Link export cepat untuk membaca data Excel dari Google Spreadsheet
SHEET_URL = "https://docs.google.com/spreadsheets/d/1ikC39Z3V9w5yypVDGVgMsfiuSInRRgvR/export?format=xlsx"

# 2. URL Web App dari Google Apps Script Anda
APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbx5mnQ71mpsG_K66m0-4ASG_aj0X7xKUDJXoMIHNvr5c5J0oAqgx25aVeepzaL2Qh5LRQ/exec"

# Memuat data live dari Google Spreadsheet (Cache 30 detik agar update cepat)
@st.cache_data(ttl=30)
def load_excel_data():
    xls = pd.ExcelFile(SHEET_URL)
    sheet_name = xls.sheet_names[0]
    df = pd.read_excel(SHEET_URL, sheet_name=sheet_name, header=0, dtype=str)
    return df, sheet_name

try:
    df_original, active_sheet = load_excel_data()
except Exception as e:
    st.error(
        "Gagal membaca data dari Google Spreadsheet. Pastikan link Google Sheet sudah diset ke 'Anyone with the link can view'.\n\nError: "
        f"{e}"
    )
    st.stop()

# --- FUNGSI MENGIRIM STATUS KE GOOGLE APPS SCRIPT ---
def update_status_to_gsheet(trx_key, status_value):
    if not APPS_SCRIPT_URL:
        return False
    try:
        payload = {"transaksi": str(trx_key), "status": str(status_value)}
        response = requests.post(APPS_SCRIPT_URL, json=payload, timeout=10)
        if response.status_code == 200:
            st.cache_data.clear()  # Bersihkan cache agar data ter-refresh
            return True
    except Exception as ex:
        st.warning(f"Gagal sinkronisasi otomatis ke Spreadsheet: {ex}")
    return False

# --- FUNGSI MENGUBAH LINK GOOGLE DRIVE MENJADI GAMBAR LANGSUNG ---
def convert_gdrive_url(url):
    if not url or not isinstance(url, str):
        return None
    url = url.strip()
    file_id = None
    
    if "drive.google.com" in url:
        if "/file/d/" in url:
            try:
                file_id = url.split("/file/d/")[1].split("/")[0]
            except: pass
        elif "open?id=" in url:
            try:
                file_id = url.split("open?id=")[1].split("&")[0]
            except: pass
        elif "id=" in url:
            try:
                file_id = url.split("id=")[1].split("&")[0]
            except: pass
            
        if file_id:
            # Menggunakan endpoint thumbnail/direct image dari Google Drive agar bisa langsung dirender st.image
            return f"https://lh3.googleusercontent.com/d/{file_id}=s800"
            
    return url

# --- PENCARIAN KOLOM OTOMATIS ---
cols = list(df_original.columns)

def find_col(keywords):
    for c in cols:
        for kw in keywords:
            if kw.lower() in str(c).lower():
                return c
    return None

col_kec = find_col(["kecamatan"])
col_trx = find_col(["no transaksi", "kode trx", "transaksi"])
col_petani = find_col(["nama petani", "petani"])
col_nik = find_col(["nik"])
col_url = find_col(["url bukti", "link", "url"])

if len(cols) >= 8:
    col_kode_kios_pos = cols[6]
    col_nama_kios_pos = cols[7]
else:
    col_kode_kios_pos = cols[-2] if len(cols) > 1 else cols[0]
    col_nama_kios_pos = cols[-1]

if len(cols) >= 24:
    col_tipe_tebus = cols[23]
else:
    col_tipe_tebus = find_col(["tipe", "tebus", "jenis"])

if not col_kec or not col_trx:
    st.error(
        "Kolom penting ('Kecamatan' atau 'No Transaksi') tidak ditemukan. Kolom terdeteksi: "
        f"{cols}"
    )
    st.stop()

df_original["Kios_Gabungan"] = (
    df_original[col_kode_kios_pos].astype(str)
    + " - "
    + df_original[col_nama_kios_pos].astype(str)
)

col_status_existing = find_col(["status_verifikasi", "status verifikasi"])

if "verifikasi_dict" not in st.session_state:
    st.session_state.verifikasi_dict = {}
    if col_status_existing:
        for _, r in df_original.iterrows():
            t_key = str(r[col_trx])
            s_val = str(r[col_status_existing])
            if s_val and s_val != "nan" and s_val != "Belum Dicek":
                st.session_state.verifikasi_dict[t_key] = s_val

# --- HEADER UTAMA ---
st.markdown(
    """
    <div style="background: linear-gradient(135deg, #002b80 0%, #0055ff 100%); padding: 20px; border-radius: 12px; color: white; margin-bottom: 20px; box-shadow: 0 4px 10px rgba(0,0,0,0.1);">
        <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px;">
            <h1 style="margin: 0; font-size: 28px; font-weight: 800; letter-spacing: 0.5px;">KARIYAWAN TIRI</h1>
            <p style="margin: 0; font-size: 14px; background: rgba(255,255,255,0.2); padding: 6px 14px; border-radius: 20px; font-weight: 500;">Pengecekan Transaksi Ipubers</p>
        </div>
        <p style="margin: 5px 0 15px 0; font-size: 13px; opacity: 0.9;">Rekapitulasi Total Keseluruhan Transaksi Kabupaten Temanggung</p>
    </div>
""",
    unsafe_allow_html=True,
)

total_kab = len(df_original)
dicek_kab = sum(1 for _, r in df_original.iterrows() if str(r[col_trx]) in st.session_state.verifikasi_dict)
terima_kab = sum(1 for _, r in df_original.iterrows() if st.session_state.verifikasi_dict.get(str(r[col_trx])) == "TERIMA")
ragu_kab = sum(1 for _, r in df_original.iterrows() if st.session_state.verifikasi_dict.get(str(r[col_trx])) == "RAGU-RAGU")
tolak_kab = sum(1 for _, r in df_original.iterrows() if st.session_state.verifikasi_dict.get(str(r[col_trx])) == "TOLAK")

mk1, mk2, mk3, mk4, mk5 = st.columns(5)
mk1.metric("Total Kab. Temanggung", total_kab)
mk2.metric("Sudah Dicek", dicek_kab)
mk3.metric("Terima", terima_kab)
mk4.metric("Ragu-Ragu", ragu_kab)
mk5.metric("Tolak", tolak_kab)

prog_kab = float(dicek_kab) / total_kab if total_kab > 0 else 0.0
st.progress(prog_kab, text=f"Progress Kabupaten Temanggung: {dicek_kab} dari {total_kab} Nota")
st.markdown("---")

# --- SIDEBAR: FILTER NAVIGASI ---
st.sidebar.markdown("<h2 style='color: #002b80; font-size: 18px;'>🎛️ Filter Data</h2>", unsafe_allow_html=True)
st.sidebar.markdown("---")

kecamatan_list = sorted(df_original[col_kec].dropna().unique().tolist())
selected_kecamatan = st.sidebar.selectbox("1. Pilih Kecamatan", ["-- Pilih Kecamatan --"] + kecamatan_list)

if selected_kecamatan != "-- Pilih Kecamatan --":
    df_filtered = df_original[df_original[col_kec] == selected_kecamatan]
else:
    df_filtered = df_original

kios_list = sorted(df_filtered["Kios_Gabungan"].dropna().astype(str).unique().tolist())
selected_nama_kios = st.sidebar.selectbox("2. Pilih Kios (Kode - Nama)", ["-- Pilih Kios --"] + kios_list)

if selected_nama_kios != "-- Pilih Kios --":
    df_filtered = df_filtered[df_filtered["Kios_Gabungan"].astype(str) == selected_nama_kios]

if col_tipe_tebus and col_tipe_tebus in df_filtered.columns:
    tipe_list = sorted(df_filtered[col_tipe_tebus].dropna().astype(str).unique().tolist())
    selected_tipe_tebus = st.sidebar.selectbox("3. Tipe Tebus", ["-- Semua Tipe --"] + tipe_list)
    if selected_tipe_tebus != "-- Semua Tipe --":
        df_filtered = df_filtered[df_filtered[col_tipe_tebus].astype(str) == selected_tipe_tebus]
else:
    selected_tipe_tebus = "-- Semua Tipe --"

st.sidebar.markdown("---")

if selected_nama_kios != "-- Pilih Kios --":
    df_kios_all = df_filtered

    st.sidebar.markdown("#### 🔎 Filter Status Nota:")
    status_filter_options = ["Semua Nota", "Belum Dicek", "TERIMA", "RAGU-RAGU", "TOLAK"]
    selected_status_filter = st.sidebar.radio("Tampilkan berdasarkan:", status_filter_options)

    filtered_indices = []
    for idx, row in df_kios_all.iterrows():
        trx_key = str(row[col_trx])
        status = st.session_state.verifikasi_dict.get(trx_key, "Belum Dicek")

        if selected_status_filter == "Belum Dicek" and status == "Belum Dicek":
            filtered_indices.append(idx)
        elif selected_status_filter == "TERIMA" and status == "TERIMA":
            filtered_indices.append(idx)
        elif selected_status_filter == "RAGU-RAGU" and status == "RAGU-RAGU":
            filtered_indices.append(idx)
        elif selected_status_filter == "TOLAK" and status == "TOLAK":
            filtered_indices.append(idx)
        elif selected_status_filter == "Semua Nota":
            filtered_indices.append(idx)

    if len(filtered_indices) == 0:
        st.warning(f"Tidak ada data nota dengan status '{selected_status_filter}' untuk filter ini.")
    else:
        if "current_pos" not in st.session_state:
            st.session_state.current_pos = 0
        if "last_filter" not in st.session_state or st.session_state.last_filter != selected_nama_kios:
            st.session_state.current_pos = 0
            st.session_state.last_filter = selected_nama_kios

        if st.session_state.current_pos >= len(filtered_indices):
            st.session_state.current_pos = len(filtered_indices) - 1
        if st.session_state.current_pos < 0:
            st.session_state.current_pos = 0

        pos = st.session_state.current_pos
        row_idx = filtered_indices[pos]
        row_data = df_original.loc[row_idx]
        current_trx_key = str(row_data[col_trx])

        total_nota_filtered = len(df_kios_all)
        sudah_cek_kios = sum(1 for _, r in df_kios_all.iterrows() if str(r[col_trx]) in st.session_state.verifikasi_dict)
        diterima_kios = sum(1 for _, r in df_kios_all.iterrows() if st.session_state.verifikasi_dict.get(str(r[col_trx])) == "TERIMA")
        ragu_kios = sum(1 for _, r in df_kios_all.iterrows() if st.session_state.verifikasi_dict.get(str(r[col_trx])) == "RAGU-RAGU")
        ditolak_kios = sum(1 for _, r in df_kios_all.iterrows() if st.session_state.verifikasi_dict.get(str(r[col_trx])) == "TOLAK")

        progress_val = float(sudah_cek_kios) / total_nota_filtered if total_nota_filtered > 0 else 0.0
        st.sidebar.markdown("---")
        st.sidebar.markdown(f"**Progress Kios Ini:** {sudah_cek_kios}/{total_nota_filtered} Nota")
        st.sidebar.progress(progress_val)

        st.markdown("<h4 style='color: #002b80; font-size: 15px;'>📌 Ringkasan Kios Terpilih</h4>", unsafe_allow_html=True)
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Total Kios", total_nota_filtered)
        m2.metric("Dicek Kios", sudah_cek_kios)
        m3.metric("Terima Kios", diterima_kios)
        m4.metric("Ragu Kios", ragu_kios)
        m5.metric("Tolak Kios", ditolak_kios)

        st.markdown("---")

        col_kiri, col_kanan = st.columns([1, 1.3], gap="medium")

        with col_kiri:
            st.markdown("<h3 style='color: #002b80; font-size: 16px;'>📄 Detail Transaksi</h3>", unsafe_allow_html=True)
            trx_val = row_data.get(col_trx, "-")
            kios_gabung_val = row_data.get("Kios_Gabungan", "-")
            petani_val = row_data.get(col_petani, "-") if col_petani else "-"

            nik_val = str(row_data[col_nik]).strip() if col_nik and pd.notna(row_data.get(col_nik)) else "-"
            if nik_val.endswith(".0"):
                nik_val = nik_val[:-2]

            tgl_val = row_data.get("Tanggal Tebus", "-")
            tipe_val = row_data.get(col_tipe_tebus, "-") if col_tipe_tebus else "-"

            st.markdown(f"**Kios:** `{kios_gabung_val}`")
            st.markdown(f"**No Transaksi:** `{trx_val}`")
            st.markdown(f"**Nama Petani:** {petani_val}")
            st.markdown(f"**NIK:** `{nik_val}`")
            st.markdown(f"**Tanggal:** {tgl_val}")
            st.markdown(f"**Tipe Tebus:** <span style='color:blue; font-weight:bold;'>{tipe_val}</span>", unsafe_allow_html=True)

            pupuk_info = []
            for p in ["Urea", "NPK", "SP36", "ZA", "Organik"]:
                if p in df_original.columns and pd.notna(row_data.get(p)):
                    val_p = str(row_data.get(p)).strip()
                    if val_p.endswith(".0"):
                        val_p = val_p[:-2]
                    pupuk_info.append(f"- {p} : **{val_p} kg**")

            if pupuk_info:
                st.markdown(f"🌾 **ALOKASI:**\n\n" + "\n".join(pupuk_info))

            current_status = st.session_state.verifikasi_dict.get(current_trx_key, "Belum Dicek")
            if current_status == "TERIMA":
                status_color = "green"
            elif current_status == "RAGU-RAGU":
                status_color = "darkorange"
            elif current_status == "TOLAK":
                status_color = "red"
            else:
                status_color = "gray"

            st.markdown(f"Status: <span style='color:{status_color}; font-weight:bold;'>{current_status}</span>", unsafe_allow_html=True)

            st.markdown("---")
            st.markdown("#### Aksi Verifikasi:")

            col_btn1, col_btn2, col_btn3 = st.columns(3)
            with col_btn1:
                if st.button("✅ TERIMA", type="primary", key=f"terima_{row_idx}"):
                    st.session_state.verifikasi_dict[current_trx_key] = "TERIMA"
                    update_status_to_gsheet(current_trx_key, "TERIMA")
                    if st.session_state.current_pos < len(filtered_indices) - 1:
                        st.session_state.current_pos += 1
                    st.rerun()

            with col_btn2:
                if st.button("⚠️ RAGU", key=f"ragu_{row_idx}"):
                    st.session_state.verifikasi_dict[current_trx_key] = "RAGU-RAGU"
                    update_status_to_gsheet(current_trx_key, "RAGU-RAGU")
                    if st.session_state.current_pos < len(filtered_indices) - 1:
                        st.session_state.current_pos += 1
                    st.rerun()

            with col_btn3:
                if st.button("❌ TOLAK", key=f"tolak_{row_idx}"):
                    st.session_state.verifikasi_dict[current_trx_key] = "TOLAK"
                    update_status_to_gsheet(current_trx_key, "TOLAK")
                    if st.session_state.current_pos < len(filtered_indices) - 1:
                        st.session_state.current_pos += 1
                    st.rerun()

            if st.button("🔄 Reset Status", key=f"reset_{row_idx}"):
                if current_trx_key in st.session_state.verifikasi_dict:
                    del st.session_state.verifikasi_dict[current_trx_key]
                    update_status_to_gsheet(current_trx_key, "Belum Dicek")
                st.rerun()

        with col_kanan:
            st.markdown("<h3 style='color: #002b80; font-size: 16px;'>🖼️ Preview Nota Bukti</h3>", unsafe_allow_html=True)
            nota_url = row_data.get(col_url, None) if col_url else None

            if pd.notna(nota_url) and str(nota_url).startswith("http"):
                raw_url = str(nota_url).strip()
                direct_img_url = convert_gdrive_url(raw_url)
                
                # Menampilkan gambar langsung di aplikasi Streamlit
                try:
                    st.image(direct_img_url, caption=f"Nota No. Transaksi: {current_trx_key}", use_container_width=True)
                except Exception:
                    st.warning("Gagal merender gambar langsung. Anda dapat membuka link berikut:")
                
                # Sediakan juga link cadangan untuk buka tab baru jika diperlukan
                st.markdown(f"<div style='text-align: center; margin-top: 8px;'><a href='{raw_url}' target='_blank'>🔗 Buka Gambar Ukuran Penuh di Tab Baru</a></div>", unsafe_allow_html=True)
            else:
                st.warning("Link bukti nota tidak tersedia pada baris ini.")

        st.markdown("---")
        nav_c1, nav_c2, nav_c3 = st.columns([1, 2, 1])
        with nav_c1:
            if st.button("⬅️ Sebelumnya", key=f"prev_{row_idx}"):
                if st.session_state.current_pos > 0:
                    st.session_state.current_pos -= 1
                    st.rerun()
        with nav_c2:
            st.markdown(f"<p style='text-align: center; font-weight: bold; margin: 5px 0;'>Nota ke-{pos + 1} dari {len(filtered_indices)}</p>", unsafe_allow_html=True)
        with nav_c3:
            if st.button("Selanjutnya ➡️", key=f"next_{row_idx}"):
                if st.session_state.current_pos < len(filtered_indices) - 1:
                    st.session_state.current_pos += 1
                    st.rerun()

st.markdown("---")
st.markdown("<h3 style='color: #002b80; font-size: 16px;'>📥 Download Hasil Verifikasi Excel</h3>", unsafe_allow_html=True)

if st.button("📊 Download Rekap Data Lengkap (Excel)", type="primary"):
    df_export = df_original.copy()
    status_list = [
        st.session_state.verifikasi_dict.get(str(row[col_trx]), "Belum Dicek")
        for _, row in df_export.iterrows()
    ]
    df_export["Status_Verifikasi"] = status_list

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_export.to_excel(writer, sheet_name=active_sheet, index=False)

    st.download_button(
        label="⬇️ Klik Disini untuk Menyimpan File Excel",
        data=output.getvalue(),
        file_name="Hasil_Verifikasi_Nota.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
