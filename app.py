"""
Panel Verifikasi Kariyawan Tiri
Versi rebuild: lebih ringan, tampilan profesional, fokus preview nota.
"""

import io
import time

import pandas as pd
import requests
import streamlit as st

# ============================================================
# KONFIGURASI
# ============================================================
SHEET_URL = "https://docs.google.com/spreadsheets/d/1ikC39Z3V9w5yypVDGVgMsfiuSInRRgvR/export?format=xlsx"
APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbx5mnQ71mpsG_K66m0-4ASG_aj0X7xKUDJXoMIHNvr5c5J0oAqgx25aVeepzaL2Qh5LRQ/exec"
CACHE_TTL = 30  # detik

STATUS_TERIMA = "TERIMA"
STATUS_RAGU = "RAGU-RAGU"
STATUS_TOLAK = "TOLAK"
STATUS_BELUM = "Belum Dicek"

STATUS_COLOR = {
    STATUS_TERIMA: "#16a34a",
    STATUS_RAGU: "#d97706",
    STATUS_TOLAK: "#dc2626",
    STATUS_BELUM: "#6b7280",
}

st.set_page_config(
    page_title="Panel Verifikasi Kariyawan Tiri",
    page_icon="🧾",
    layout="wide",
)

# ============================================================
# CSS PROFESIONAL (ringan, tanpa library tambahan)
# ============================================================
st.markdown(
    """
    <style>
    .block-container {padding-top: 1.4rem; padding-bottom: 2rem; max-width: 1200px;}
    #MainMenu, footer {visibility: hidden;}

    .kt-header {
        background: linear-gradient(120deg, #0f172a 0%, #1d4ed8 100%);
        padding: 22px 28px;
        border-radius: 14px;
        color: white;
        margin-bottom: 18px;
    }
    .kt-header h1 {margin: 0; font-size: 24px; font-weight: 800; letter-spacing: .3px;}
    .kt-header p {margin: 4px 0 0 0; font-size: 13px; opacity: .85;}

    div[data-testid="stMetric"] {
        background: #ffffff;
        border: 1px solid #e5e7eb;
        border-radius: 10px;
        padding: 10px 6px;
    }

    .kt-card {
        background: #ffffff;
        border: 1px solid #e5e7eb;
        border-radius: 12px;
        padding: 18px 20px;
    }
    .kt-label {color: #64748b; font-size: 12px; font-weight: 600; text-transform: uppercase; letter-spacing: .04em;}
    .kt-value {color: #0f172a; font-size: 15px; font-weight: 600; margin-bottom: 10px;}
    .kt-status-pill {
        display: inline-block; padding: 4px 12px; border-radius: 999px;
        font-size: 12px; font-weight: 700; color: white;
    }
    .kt-nota-count {text-align:center; font-weight:600; color:#334155; font-size: 13px;}
    </style>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# DATA LOADING
# ============================================================
@st.cache_data(ttl=CACHE_TTL, show_spinner="Memuat data dari Google Spreadsheet...")
def load_excel_data():
    xls = pd.ExcelFile(SHEET_URL)
    sheet_name = xls.sheet_names[0]
    df = pd.read_excel(SHEET_URL, sheet_name=sheet_name, header=0, dtype=str)
    return df, sheet_name


def find_col(cols, keywords):
    for c in cols:
        for kw in keywords:
            if kw.lower() in str(c).lower():
                return c
    return None


def gdrive_image_candidates(url: str):
    """Kembalikan kandidat URL gambar thumbnail untuk link Google Drive."""
    if not url or not isinstance(url, str):
        return []
    url = url.strip()
    if "drive.google.com" not in url:
        return [url]

    file_id = None
    if "/file/d/" in url:
        try:
            file_id = url.split("/file/d/")[1].split("/")[0]
        except IndexError:
            pass
    elif "id=" in url:
        try:
            file_id = url.split("id=")[1].split("&")[0]
        except IndexError:
            pass

    if not file_id:
        return [url]

    return [f"https://lh3.googleusercontent.com/d/{file_id}=s1000"]


def resolve_image_url(url: str) -> str:
    """Kalau link Google Drive, ubah ke format thumbnail. Selain itu (mis. link
    nota dari aplikasi lain), pakai URL aslinya apa adanya — dimuat langsung
    lewat browser pengguna, bukan lewat server, supaya tidak kena proteksi
    anti-bot di sisi server."""
    if url and "drive.google.com" in url:
        candidates = gdrive_image_candidates(url)
        return candidates[0] if candidates else url
    return url


def update_status_to_gsheet(trx_key, status_value):
    if not APPS_SCRIPT_URL:
        return False
    try:
        payload = {"transaksi": str(trx_key), "status": str(status_value)}
        r = requests.post(APPS_SCRIPT_URL, json=payload, timeout=10)
        if r.status_code == 200:
            st.cache_data.clear()
            return True
    except Exception as ex:
        st.toast(f"Gagal sinkronisasi ke Spreadsheet: {ex}", icon="⚠️")
    return False


# ============================================================
# LOAD DATA
# ============================================================
try:
    df_original, active_sheet = load_excel_data()
except Exception as e:
    st.error(
        "Gagal membaca data dari Google Spreadsheet. Pastikan link sudah diset ke "
        f"'Anyone with the link can view'.\n\nError: {e}"
    )
    st.stop()

cols = list(df_original.columns)
col_kec = find_col(cols, ["kecamatan"])
col_trx = find_col(cols, ["no transaksi", "kode trx", "transaksi"])
col_petani = find_col(cols, ["nama petani", "petani"])
col_nik = find_col(cols, ["nik"])
col_url = find_col(cols, ["url bukti", "link", "url"])
col_status_existing = find_col(cols, ["status_verifikasi", "status verifikasi"])

col_kode_kios_pos = cols[6] if len(cols) >= 8 else (cols[-2] if len(cols) > 1 else cols[0])
col_nama_kios_pos = cols[7] if len(cols) >= 8 else cols[-1]
col_tipe_tebus = cols[23] if len(cols) >= 24 else find_col(cols, ["tipe", "tebus", "jenis"])

if not col_kec or not col_trx:
    st.error(f"Kolom 'Kecamatan' atau 'No Transaksi' tidak ditemukan. Kolom terdeteksi: {cols}")
    st.stop()

df_original["Kios_Gabungan"] = (
    df_original[col_kode_kios_pos].astype(str) + " - " + df_original[col_nama_kios_pos].astype(str)
)

# state verifikasi (di-load sekali dari kolom status eksisting bila ada)
if "verifikasi_dict" not in st.session_state:
    st.session_state.verifikasi_dict = {}
    if col_status_existing:
        mask = df_original[col_status_existing].notna() & ~df_original[col_status_existing].isin(["nan", STATUS_BELUM])
        for trx, status in zip(df_original.loc[mask, col_trx], df_original.loc[mask, col_status_existing]):
            st.session_state.verifikasi_dict[str(trx)] = str(status)

verif = st.session_state.verifikasi_dict


def status_of(trx_key):
    return verif.get(str(trx_key), STATUS_BELUM)


# hitung status sekali untuk semua baris (lebih ringan daripada iterrows berulang)
df_original["_status"] = df_original[col_trx].astype(str).map(lambda k: verif.get(k, STATUS_BELUM))

# ============================================================
# HEADER
# ============================================================
st.markdown(
    """
    <div class="kt-header">
        <h1>🧾 KARIYAWAN TIRI — Panel Verifikasi</h1>
        <p>Pengecekan transaksi Ipubers · Kabupaten Temanggung</p>
    </div>
    """,
    unsafe_allow_html=True,
)

total_kab = len(df_original)
dicek_kab = int((df_original["_status"] != STATUS_BELUM).sum())
terima_kab = int((df_original["_status"] == STATUS_TERIMA).sum())
ragu_kab = int((df_original["_status"] == STATUS_RAGU).sum())
tolak_kab = int((df_original["_status"] == STATUS_TOLAK).sum())

mk1, mk2, mk3, mk4, mk5 = st.columns(5)
mk1.metric("Total Nota", total_kab)
mk2.metric("Sudah Dicek", dicek_kab)
mk3.metric("Terima", terima_kab)
mk4.metric("Ragu-Ragu", ragu_kab)
mk5.metric("Tolak", tolak_kab)

prog_kab = (dicek_kab / total_kab) if total_kab else 0.0
st.progress(prog_kab, text=f"Progress Kabupaten: {dicek_kab} dari {total_kab} nota")
st.markdown("")

# ============================================================
# SIDEBAR FILTER
# ============================================================
with st.sidebar:
    st.markdown("### 🎛️ Filter Data")
    st.markdown("---")

    kecamatan_list = sorted(df_original[col_kec].dropna().unique().tolist())
    selected_kecamatan = st.selectbox("1. Kecamatan", ["-- Pilih Kecamatan --"] + kecamatan_list)

    df_filtered = (
        df_original[df_original[col_kec] == selected_kecamatan]
        if selected_kecamatan != "-- Pilih Kecamatan --"
        else df_original
    )

    kios_list = sorted(df_filtered["Kios_Gabungan"].dropna().astype(str).unique().tolist())
    selected_nama_kios = st.selectbox("2. Kios (Kode - Nama)", ["-- Pilih Kios --"] + kios_list)

    if selected_nama_kios != "-- Pilih Kios --":
        df_filtered = df_filtered[df_filtered["Kios_Gabungan"].astype(str) == selected_nama_kios]

    if col_tipe_tebus and col_tipe_tebus in df_filtered.columns:
        tipe_list = sorted(df_filtered[col_tipe_tebus].dropna().astype(str).unique().tolist())
        selected_tipe = st.selectbox("3. Tipe Tebus", ["-- Semua Tipe --"] + tipe_list)
        if selected_tipe != "-- Semua Tipe --":
            df_filtered = df_filtered[df_filtered[col_tipe_tebus].astype(str) == selected_tipe]

    st.markdown("---")

    if selected_nama_kios != "-- Pilih Kios --":
        st.markdown("#### 🔎 Status Nota")
        status_filter = st.radio(
            "Tampilkan berdasarkan:",
            ["Semua Nota", "Belum Dicek", STATUS_TERIMA, STATUS_RAGU, STATUS_TOLAK],
            label_visibility="collapsed",
        )

# ============================================================
# AREA UTAMA: DETAIL + AKSI
# ============================================================
if selected_nama_kios != "-- Pilih Kios --":
    df_kios_all = df_filtered

    if status_filter == "Semua Nota":
        df_view = df_kios_all
    else:
        df_view = df_kios_all[df_kios_all["_status"] == status_filter]

    filtered_indices = df_view.index.tolist()

    total_nota_filtered = len(df_kios_all)
    sudah_cek_kios = int((df_kios_all["_status"] != STATUS_BELUM).sum())
    diterima_kios = int((df_kios_all["_status"] == STATUS_TERIMA).sum())
    ragu_kios = int((df_kios_all["_status"] == STATUS_RAGU).sum())
    ditolak_kios = int((df_kios_all["_status"] == STATUS_TOLAK).sum())
    progress_val = (sudah_cek_kios / total_nota_filtered) if total_nota_filtered else 0.0

    with st.sidebar:
        st.markdown("---")
        st.markdown(f"**Progress Kios Ini:** {sudah_cek_kios}/{total_nota_filtered} nota")
        st.progress(progress_val)

    if not filtered_indices:
        st.warning(f"Tidak ada data nota dengan status '{status_filter}' untuk filter ini.")
    else:
        # reset posisi bila filter/kios berubah
        if "current_pos" not in st.session_state:
            st.session_state.current_pos = 0
        filter_key = f"{selected_nama_kios}|{status_filter}"
        if st.session_state.get("last_filter") != filter_key:
            st.session_state.current_pos = 0
            st.session_state.last_filter = filter_key
        st.session_state.current_pos = max(0, min(st.session_state.current_pos, len(filtered_indices) - 1))

        pos = st.session_state.current_pos
        row_idx = filtered_indices[pos]
        row_data = df_original.loc[row_idx]
        current_trx_key = str(row_data[col_trx])
        current_status = status_of(current_trx_key)

        st.markdown("##### 📌 Ringkasan Kios Terpilih")
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Total", total_nota_filtered)
        m2.metric("Dicek", sudah_cek_kios)
        m3.metric("Terima", diterima_kios)
        m4.metric("Ragu", ragu_kios)
        m5.metric("Tolak", ditolak_kios)
        st.markdown("---")

        col_kiri, col_kanan = st.columns([1, 1.35], gap="large")

        # ---------------- DETAIL + AKSI ----------------
        with col_kiri:
            st.markdown('<div class="kt-card">', unsafe_allow_html=True)

            nik_val = str(row_data.get(col_nik, "-")).strip() if col_nik else "-"
            if nik_val.endswith(".0"):
                nik_val = nik_val[:-2]
            tgl_val = row_data.get("Tanggal Tebus", "-")
            tipe_val = row_data.get(col_tipe_tebus, "-") if col_tipe_tebus else "-"
            petani_val = row_data.get(col_petani, "-") if col_petani else "-"

            st.markdown('<div class="kt-label">Kios</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="kt-value">{row_data.get("Kios_Gabungan", "-")}</div>', unsafe_allow_html=True)

            st.markdown('<div class="kt-label">No. Transaksi</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="kt-value">{current_trx_key}</div>', unsafe_allow_html=True)

            st.markdown('<div class="kt-label">Nama Petani / NIK</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="kt-value">{petani_val} &nbsp;·&nbsp; {nik_val}</div>', unsafe_allow_html=True)

            st.markdown('<div class="kt-label">Tanggal / Tipe Tebus</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="kt-value">{tgl_val} &nbsp;·&nbsp; {tipe_val}</div>', unsafe_allow_html=True)

            pupuk_info = []
            for p in ["Urea", "NPK", "SP36", "ZA", "Organik"]:
                if p in df_original.columns and pd.notna(row_data.get(p)):
                    val_p = str(row_data.get(p)).strip()
                    if val_p.endswith(".0"):
                        val_p = val_p[:-2]
                    pupuk_info.append(f"{p}: **{val_p} kg**")
            if pupuk_info:
                st.markdown('<div class="kt-label">Alokasi</div>', unsafe_allow_html=True)
                st.markdown(" &nbsp;·&nbsp; ".join(pupuk_info))

            color = STATUS_COLOR.get(current_status, "#6b7280")
            st.markdown(
                f'<br><span class="kt-status-pill" style="background:{color};">{current_status}</span>',
                unsafe_allow_html=True,
            )
            st.markdown("</div>", unsafe_allow_html=True)

            st.markdown("#### Aksi Verifikasi")
            b1, b2, b3 = st.columns(3)

            def set_status(new_status):
                st.session_state.verifikasi_dict[current_trx_key] = new_status
                ok = update_status_to_gsheet(current_trx_key, new_status)
                if not ok:
                    st.toast("Status tersimpan lokal, tapi sinkron ke Sheet gagal.", icon="⚠️")
                if st.session_state.current_pos < len(filtered_indices) - 1:
                    st.session_state.current_pos += 1
                st.rerun()

            with b1:
                if st.button("✅ TERIMA", type="primary", use_container_width=True, key=f"terima_{row_idx}"):
                    set_status(STATUS_TERIMA)
            with b2:
                if st.button("⚠️ RAGU", use_container_width=True, key=f"ragu_{row_idx}"):
                    set_status(STATUS_RAGU)
            with b3:
                if st.button("❌ TOLAK", use_container_width=True, key=f"tolak_{row_idx}"):
                    set_status(STATUS_TOLAK)

            if st.button("🔄 Reset Status", key=f"reset_{row_idx}"):
                if current_trx_key in st.session_state.verifikasi_dict:
                    del st.session_state.verifikasi_dict[current_trx_key]
                    update_status_to_gsheet(current_trx_key, STATUS_BELUM)
                st.rerun()

        # ---------------- PREVIEW NOTA (fokus utama) ----------------
        with col_kanan:
            st.markdown("#### 🖼️ Preview Nota Bukti")
            raw_url = row_data.get(col_url, None) if col_url else None

            if pd.notna(raw_url) and str(raw_url).strip().startswith("http"):
                raw_url = str(raw_url).strip()
                img_url = resolve_image_url(raw_url)
                safe_key = "".join(ch for ch in current_trx_key if ch.isalnum())
                st.markdown(
                    f"""
                    <div style="text-align:center;">
                      <img src="{img_url}" referrerpolicy="no-referrer"
                           style="max-width:100%; border-radius:10px;
                           border:1px solid #e5e7eb;"
                           onerror="this.onerror=null; this.style.display='none';
                                    document.getElementById('nota-err-{safe_key}').style.display='block';">
                      <div id="nota-err-{safe_key}" style="display:none; padding:16px;
                           background:#fef2f2; border:1px solid #fecaca; border-radius:8px;
                           color:#b91c1c; margin-top:8px; font-size:13px;">
                          Gambar nota tidak bisa ditampilkan langsung di sini
                          (situs sumber mungkin membatasi akses tersemat).
                      </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                st.markdown(
                    f"<div style='text-align:center; margin-top:8px;'>"
                    f"<a href='{raw_url}' target='_blank'>🔗 Buka Nota di Tab Baru</a></div>",
                    unsafe_allow_html=True,
                )
            else:
                st.info("Link bukti nota tidak tersedia pada baris ini.")

        st.markdown("---")
        nav1, nav2, nav3 = st.columns([1, 2, 1])
        with nav1:
            if st.button("⬅️ Sebelumnya", key=f"prev_{row_idx}", disabled=(pos == 0)):
                st.session_state.current_pos -= 1
                st.rerun()
        with nav2:
            st.markdown(f'<p class="kt-nota-count">Nota ke-{pos + 1} dari {len(filtered_indices)}</p>', unsafe_allow_html=True)
        with nav3:
            if st.button("Selanjutnya ➡️", key=f"next_{row_idx}", disabled=(pos == len(filtered_indices) - 1)):
                st.session_state.current_pos += 1
                st.rerun()
else:
    st.info("👈 Pilih Kecamatan lalu Kios pada sidebar untuk mulai memverifikasi nota.")

# ============================================================
# DOWNLOAD REKAP
# ============================================================
st.markdown("---")
st.markdown("#### 📥 Download Hasil Verifikasi")
if st.button("📊 Download Rekap Data Lengkap (Excel)", type="primary"):
    df_export = df_original.drop(columns=["_status"], errors="ignore").copy()
    df_export["Status_Verifikasi"] = [
        status_of(row[col_trx]) for _, row in df_export.iterrows()
    ]
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_export.to_excel(writer, sheet_name=active_sheet, index=False)
    st.download_button(
        label="⬇️ Klik Disini untuk Menyimpan File Excel",
        data=output.getvalue(),
        file_name=f"Hasil_Verifikasi_Nota_{int(time.time())}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
