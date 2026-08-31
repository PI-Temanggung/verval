import io
import openpyxl
import pandas as pd
import streamlit as st

# Konfigurasi Halaman (Lebar responsif untuk PC dan HP)
st.set_page_config(
    page_title="Panel Verifikasi Joko Winarno", page_icon="📱", layout="wide"
)

# Link export otomatis dari Spreadsheet Anda
SHEET_URL = (
    "https://docs.google.com/spreadsheets/d/1ikC39Z3V9w5yypVDGVgMsfiuSInRRgvR/export?format=xlsx"
)


# Memuat data live dari Google Spreadsheet (Cache 30 detik agar update cepat)
@st.cache_data(ttl=30)
def load_excel_data():
  xls = pd.ExcelFile(SHEET_URL)
  sheet_name = xls.sheet_names[0]
  df = pd.read_excel(SHEET_URL, sheet_name=sheet_name, header=0)
  return df, sheet_name


try:
  df_original, active_sheet = load_excel_data()
except Exception as e:
  st.error(
      "Gagal membaca data dari Google Spreadsheet. Pastikan akses sharing diset"
      " ke 'Anyone with the link can view'.\n\nError: "
      f"{e}"
  )
  st.stop()

# --- HEADER UTAMA (Ramah Tampilan HP) ---
st.markdown(
    """
    <div style="background-color: #f8f9fa; padding: 15px; border-radius: 10px; border-left: 6px solid #0055ff; margin-bottom: 15px;">
        <h1 style="color: #002b80; margin: 0; font-size: 22px;">Joko Winarno</h1>
        <p style="color: #555555; margin: 3px 0 0 0; font-size: 13px;">Panel Verifikasi Nota Pupuk Bersubsidi (Mobile Friendly)</p>
    </div>
""",
    unsafe_allow_html=True,
)

# --- PENCARIAN KOLOM SECARA OTOMATIS ---
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
col_url = find_col(["url bukti", "link", "url"])

# Kolom G (indeks 6) & H (indeks 7) untuk Kios
if len(cols) >= 8:
  col_kode_kios_pos = cols[6]
  col_nama_kios_pos = cols[7]
else:
  col_kode_kios_pos = cols[-2] if len(cols) > 1 else cols[0]
  col_nama_kios_pos = cols[-1]

# Kolom X (indeks 23) untuk Tipe Tebus (Individu / Kelompok)
if len(cols) >= 24:
  col_tipe_tebus = cols[23]
else:
  col_tipe_tebus = find_col(["tipe", "tebus", "jenis"])

if not col_kec or not col_trx:
  st.error(
      "Kolom penting ('Kecamatan' atau 'No Transaksi') tidak ditemukan. Kolom"
      f" terdeteksi: {cols}"
  )
  st.stop()

# --- GABUNGKAN KODE KIOS (G) DAN NAMA KIOS (H) ---
df_original["Kios_Gabungan"] = (
    df_original[col_kode_kios_pos].astype(str)
    + " - "
    + df_original[col_nama_kios_pos].astype(str)
)

# --- SIDEBAR: FILTER NAVIGASI ---
st.sidebar.markdown(
    "<h2 style='color: #002b80; font-size: 18px;'>🎛️ Filter Data</h2>",
    unsafe_allow_html=True,
)
st.sidebar.markdown("---")

kecamatan_list = sorted(df_original[col_kec].dropna().unique().tolist())
selected_kecamatan = st.sidebar.selectbox(
    "1. Pilih Kecamatan", ["-- Pilih Kecamatan --"] + kecamatan_list
)

if selected_kecamatan != "-- Pilih Kecamatan --":
  df_filtered = df_original[df_original[col_kec] == selected_kecamatan]
else:
  df_filtered = df_original

# Filter Kios Gabungan
kios_list = sorted(
    df_filtered["Kios_Gabungan"].dropna().astype(str).unique().tolist()
)
selected_nama_kios = st.sidebar.selectbox(
    "2. Pilih Kios (Kode - Nama)", ["-- Pilih Kios --"] + kios_list
)

if selected_nama_kios != "-- Pilih Kios --":
  df_filtered = df_filtered[
      df_filtered["Kios_Gabungan"].astype(str) == selected_nama_kios
  ]

# Filter Tipe Tebus (Kolom X: Individu / Kelompok)
if col_tipe_tebus and col_tipe_tebus in df_filtered.columns:
  tipe_list = sorted(
      df_filtered[col_tipe_tebus].dropna().astype(str).unique().tolist()
  )
  selected_tipe_tebus = st.sidebar.selectbox(
      "3. Tipe Tebus", ["-- Semua Tipe --"] + tipe_list
  )
  if selected_tipe_tebus != "-- Semua Tipe --":
    df_filtered = df_filtered[
        df_filtered[col_tipe_tebus].astype(str) == selected_tipe_tebus
    ]
else:
  selected_tipe_tebus = "-- Semua Tipe --"

st.sidebar.markdown("---")

# Inisialisasi Session State Verifikasi Lokal
if "verifikasi_dict" not in st.session_state:
  st.session_state.verifikasi_dict = {}

if selected_nama_kios != "-- Pilih Kios --":
  df_kios_all = df_filtered

  st.sidebar.markdown("#### 🔎 Filter Status Nota:")
  status_filter_options = [
      "Semua Nota",
      "Belum Dicek",
      "TERIMA",
      "RAGU-RAGU",
      "TOLAK",
  ]
  selected_status_filter = st.sidebar.radio(
      "Tampilkan berdasarkan:", status_filter_options
  )

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
    st.warning(
        f"Tidak ada data nota dengan status '{selected_status_filter}' untuk"
        " filter ini."
    )
  else:
    if "current_pos" not in st.session_state:
      st.session_state.current_pos = 0
    if (
        "last_filter" not in st.session_state
        or st.session_state.last_filter != selected_nama_kios
    ):
      st.session_state.current_pos = 0
      st.session_state.last_filter = selected_nama_kios

    if st.session_state.current_pos >= len(filtered_indices):
      st.session_state.current_pos = 0

    pos = st.session_state.current_pos
    row_idx = filtered_indices[pos]
    row_data = df_original.loc[row_idx]
    current_trx_key = str(row_data[col_trx])

    # Hitung Statistik
    total_nota_filtered = len(df_kios_all)
    sudah_cek = sum(
        1
        for _, r in df_kios_all.iterrows()
        if str(r[col_trx]) in st.session_state.verifikasi_dict
    )
    diterima = sum(
        1
        for _, r in df_kios_all.iterrows()
        if st.session_state.verifikasi_dict.get(str(r[col_trx])) == "TERIMA"
    )
    ragu = sum(
        1
        for _, r in df_kios_all.iterrows()
        if st.session_state.verifikasi_dict.get(str(r[col_trx])) == "RAGU-RAGU"
    )
    ditolak = sum(
        1
        for _, r in df_kios_all.iterrows()
        if st.session_state.verifikasi_dict.get(str(r[col_trx])) == "TOLAK"
    )

    # Progress Bar di Sidebar
    progress_val = (
        float(sudah_cek) / total_nota_filtered
        if total_nota_filtered > 0
        else 0.0
    )
    st.sidebar.markdown("---")
    st.sidebar.markdown(f"**Progress:** {sudah_cek}/{total_nota_filtered} Nota")
    st.sidebar.progress(progress_val)

    # Metrik Ringkasan (5 kolom dengan Ragu-Ragu)
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Total", total_nota_filtered)
    m2.metric("Dicek", sudah_cek)
    m3.metric("Terima", diterima)
    m4.metric("Ragu", ragu)
    m5.metric("Tolak", ditolak)

    st.markdown("---")

    # --- TAMPILAN UTAMA (Preview Nota Diperlebar/Diperpanjang & Detail Kiri) ---
    col_kiri, col_kanan = st.columns([1, 1.3], gap="medium")

    with col_kiri:
      st.markdown(
          "<h3 style='color: #002b80; font-size: 16px;'>📄 Detail"
          " Transaksi</h3>",
          unsafe_allow_html=True,
      )
      trx_val = row_data.get(col_trx, "-")
      kios_gabung_val = row_data.get("Kios_Gabungan", "-")
      petani_val = row_data.get(col_petani, "-") if col_petani else "-"
      nik_val = row_data.get("NIK", "-")
      tgl_val = row_data.get("Tanggal Tebus", "-")
      tipe_val = (
          row_data.get(col_tipe_tebus, "-") if col_tipe_tebus else "-"
      )

      st.markdown(f"**Kios:** `{kios_gabung_val}`")
      st.markdown(f"**No Transaksi:** `{trx_val}`")
      st.markdown(f"**Nama Petani:** {petani_val}")
      st.markdown(f"**NIK:** {nik_val}")
      st.markdown(f"**Tanggal:** {tgl_val}")
      st.markdown(
          f"**Tipe Tebus:** <span style='color:blue;"
          f" font-weight:bold;'>{tipe_val}</span>",
          unsafe_allow_html=True,
      )

      # Format alokasi pupuk ke bawah (vertikal)
      pupuk_info = []
      for p in ["Urea", "NPK", "SP36", "ZA", "Organik"]:
        if p in df_original.columns and pd.notna(row_data.get(p)):
          pupuk_info.append(f"- {p} : **{row_data.get(p)} kg**")

      if pupuk_info:
        st.markdown(f"🌾 **ALOKASI:**\n\n" + "\n".join(pupuk_info))

      current_status = st.session_state.verifikasi_dict.get(
          current_trx_key, "Belum Dicek"
      )
      if current_status == "TERIMA":
        status_color = "green"
      elif current_status == "RAGU-RAGU":
        status_color = "darkorange"
      elif current_status == "TOLAK":
        status_color = "red"
      else:
        status_color = "gray"

      st.markdown(
          f"Status: <span"
          f" style='color:{status_color}; font-weight:bold;'>{current_status}</span>",
          unsafe_allow_html=True,
      )

      st.markdown("---")
      st.markdown("#### Aksi Verifikasi:")

      col_btn1, col_btn2, col_btn3 = st.columns(3)
      with col_btn1:
        if st.button("✅ TERIMA", type="primary", key=f"terima_{row_idx}"):
          st.session_state.verifikasi_dict[current_trx_key] = "TERIMA"
          if pos < len(filtered_indices) - 1:
            st.session_state.current_pos += 1
          st.rerun()

      with col_btn2:
        if st.button("⚠️ RAGU", key=f"ragu_{row_idx}"):
          st.session_state.verifikasi_dict[current_trx_key] = "RAGU-RAGU"
          if pos < len(filtered_indices) - 1:
            st.session_state.current_pos += 1
          st.rerun()

      with col_btn3:
        if st.button("❌ TOLAK", key=f"tolak_{row_idx}"):
          st.session_state.verifikasi_dict[current_trx_key] = "TOLAK"
          if pos < len(filtered_indices) - 1:
            st.session_state.current_pos += 1
          st.rerun()

      if st.button("🔄 Reset Status", key=f"reset_{row_idx}"):
        if current_trx_key in st.session_state.verifikasi_dict:
          del st.session_state.verifikasi_dict[current_trx_key]
        st.rerun()

    with col_kanan:
      st.markdown(
          "<h3 style='color: #002b80; font-size: 16px;'>🖼️ Preview Nota"
          " Bukti (Lebih Luas)</h3>",
          unsafe_allow_html=True,
      )
      nota_url = row_data.get(col_url, None) if col_url else None

      if pd.notna(nota_url) and str(nota_url).startswith("http"):
        # Ukuran iframe diperpanjang menjadi 650px agar preview lebih optimal
        st.markdown(
            f'<iframe src="{nota_url}" width="100%" height="650px"'
            ' style="border: 2px solid #0055ff; border-radius: 8px;'
            ' background-color: white;"></iframe>',
            unsafe_allow_html=True,
        )
        st.markdown(f"🔗 [Buka Gambar Full di Tab Baru]({nota_url})")
      else:
        st.warning("Link bukti nota tidak tersedia.")

    # Navigasi Antar Nota di Bawah
    st.markdown("---")
    nav_c1, nav_c2, nav_c3 = st.columns([1, 2, 1])
    with nav_c1:
      if st.button("⬅️ Sebelumnya", key=f"prev_{row_idx}"):
        if pos > 0:
          st.session_state.current_pos -= 1
          st.rerun()
    with nav_c2:
      st.markdown(
          f"<p style='text-align: center; font-weight: bold; margin: 5px"
          f" 0;'>Nota ke-{pos + 1} dari {len(filtered_indices)}</p>",
          unsafe_allow_html=True,
      )
    with nav_c3:
      if st.button("Selanjutnya ➡️", key=f"next_{row_idx}"):
        if pos < len(filtered_indices) - 1:
          st.session_state.current_pos += 1
          st.rerun()

# --- PANEL DOWNLOAD HASIL REKAP ---
st.markdown("---")
st.markdown(
    "<h3 style='color: #002b80; font-size: 16px;'>📥 Download Hasil"
    " Verifikasi Excel</h3>",
    unsafe_allow_html=True,
)

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
