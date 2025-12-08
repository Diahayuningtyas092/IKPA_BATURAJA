import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io
import numpy as np
import os
import base64
import requests
import re
import calendar
from pathlib import Path
from datetime import datetime
from github import Github
from github import Auth
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

# define month order map
MONTH_ORDER = {
    'JANUARI': 1,
    'FEBRUARI': 2, 'PEBRUARI': 2, 'PEBRUARY': 2,
    'MARET': 3, 'MAR': 3, 'MRT': 3,
    'APRIL': 4,
    'MEI': 5,
    'JUNI': 6,
    'JULI': 7,
    'AGUSTUS': 8, 'AGUSTUSS': 8,
    'SEPTEMBER': 9, 'SEPT': 9, 'SEP': 9,
    'OKTOBER': 10,
    'NOVEMBER': 11, 'NOPEMBER': 11,
    'DESEMBER': 12
}

# Konfigurasi halaman
st.set_page_config(
    page_title="Dashboard IKPA KPPN Baturaja",
    page_icon="📊",
    layout="wide"
)

# Path ke file template (akan diatur di session state)
TEMPLATE_PATH = r"C:\Users\KEMENKEU\Desktop\INDIKATOR PELAKSANAAN ANGGARAN.xlsx"

# Inisialisasi session state untuk menyimpan data dan aktivitas
if 'data_storage' not in st.session_state:
    st.session_state.data_storage = {}

if 'activity_log' not in st.session_state:
    st.session_state.activity_log = []  # Each entry: dict with timestamp, action, period, status

# ------------------------------

def normalize_kode_satker(k, width=6):
    """
    Pastikan Kode Satker sebagai string digit dengan leading zero.
    Jika input None/empty -> return ''.
    """
    if pd.isna(k):
        return ''
    s = str(k).strip()
    # ambil hanya digit
    digits = re.findall(r'\d+', s)
    if not digits:
        return ''
    # biasanya kode terletak di awal; gabungkan semua digit found
    kod = digits[0]
    # pad left with zeros hingga panjang width
    kod = kod.zfill(width)
    return kod


def extract_kode_from_satker_field(s, width=6):
    """
    Jika kolom 'Satker' mengandung '001234 – NAMA SATKER', ambil angka di awal.
    Jika hanya angka (sebagai int/str), return padded string.
    """
    if pd.isna(s):
        return ''
    stxt = str(s).strip()
    # cari angka di awal baris (atau angka pertama)
    m = re.match(r'^\s*0*\d+', stxt)
    if m:
        return normalize_kode_satker(m.group(0), width=width)
    # fallback: cari first group of digits anywhere
    m2 = re.search(r'(\d+)', stxt)
    if m2:
        return normalize_kode_satker(m2.group(1), width=width)
    return ''


# =============================================================
# ✅ FUNGSI BARU 1: Memastikan DIPA Unik per Satker
# =============================================================
def ensure_unique_dipa_per_satker(df_dipa):
    """
    Memastikan setiap Kode Satker hanya punya 1 row dengan revisi terbaru.
    """
    if df_dipa is None or df_dipa.empty:
        return df_dipa
    
    df = df_dipa.copy()
    
    # Normalize Kode Satker
    if 'Kode Satker' in df.columns:
        df['Kode Satker'] = df['Kode Satker'].apply(normalize_kode_satker)
    
    # Pastikan Tanggal Posting Revisi dalam format datetime
    if 'Tanggal Posting Revisi' in df.columns:
        df['Tanggal Posting Revisi'] = pd.to_datetime(
            df['Tanggal Posting Revisi'], 
            errors='coerce'
        )
        
        # Sort by Kode Satker dan Tanggal (terbaru di bawah)
        df = df.sort_values(
            by=['Kode Satker', 'Tanggal Posting Revisi'], 
            ascending=[True, True]
        )
        
        # Ambil row terakhir (terbaru) per Kode Satker
        df_unique = df.groupby('Kode Satker', as_index=False).last()
    else:
        # Jika tidak ada tanggal, ambil row terakhir per Kode Satker
        df_unique = df.groupby('Kode Satker', as_index=False).last()
    
    return df_unique.reset_index(drop=True)


# =============================================================
# ✅ FUNGSI BARU 2: Menambahkan Kolom Jenis Satker
# =============================================================
def add_jenis_satker_column(df_merged):
    """
    Menambahkan kolom 'Jenis Satker' berdasarkan Total Pagu DIPA.
    - Top 30% = "Satker Besar"
    - 40%-70% = "Satker Sedang"  
    - 0%-40% = "Satker Kecil"
    """
    if df_merged is None or df_merged.empty:
        return df_merged
    
    df = df_merged.copy()
    
    # Pastikan kolom Total Pagu DIPA ada dan numeric
    if 'Total Pagu DIPA' not in df.columns:
        df['Jenis Satker'] = 'Tidak Ada Data Pagu'
        return df
    
    # Convert ke numeric, handle errors
    df['Total Pagu DIPA'] = pd.to_numeric(
        df['Total Pagu DIPA'], 
        errors='coerce'
    ).fillna(0)
    
    # Hitung persentil (hanya dari nilai > 0)
    valid_pagu = df[df['Total Pagu DIPA'] > 0]['Total Pagu DIPA']
    
    if len(valid_pagu) == 0:
        df['Jenis Satker'] = 'Tidak Ada Data Pagu'
        return df
    
    # Persentil 70 (top 30%)
    p70 = valid_pagu.quantile(0.70)
    # Persentil 40
    p40 = valid_pagu.quantile(0.40)
    
    # Fungsi kategorisasi
    def kategorisasi(pagu):
        if pd.isna(pagu) or pagu == 0:
            return "Tidak Ada Data Pagu"
        elif pagu >= p70:
            return "Satker Besar"
        elif pagu >= p40:
            return "Satker Sedang"
        else:
            return "Satker Kecil"
    
    df['Jenis Satker'] = df['Total Pagu DIPA'].apply(kategorisasi)
    
    return df


def process_dipa_dataframe(df, source_name=None, date_col_candidates=None):
    """
    Proses dataframe DIPA mentah -> normalisasi Kode Satker, parse tanggal revisi,
    ekstrak Year dari tanggal revisi (atau fallback ke kolom Tahun jika tersedia),
    lalu pilih revisi terbaru per Kode Satker per Year.
    Mengembalikan dataframe bersih dengan kolom minimal:
      ['Kode Satker','Tahun','Tanggal Posting Revisi', ... kolom asli lainnya ...]
    """
    if df is None or not isinstance(df, pd.DataFrame):
        return None

    # normalize headers whitespace
    df = df.rename(columns={c: c.strip() for c in df.columns})

    # 1) Temukan kolom tanggal revisi (gunakan header yang diberikan: 'Tanggal Posting Revisi')
    date_col = None
    if date_col_candidates is None:
        date_col_candidates = ['Tanggal Posting Revisi','Tanggal Revisi','Tgl Revisi','Tanggal','TGL REVISI']
    # check exact presence first
    for name in date_col_candidates:
        if name in df.columns:
            date_col = name
            break
    # fallback: coba cari kolom yang mengandung kata 'revisi' atau 'tanggal'
    if date_col is None:
        for c in df.columns:
            if 'revisi' in c.lower() or 'tanggal' in c.lower() or 'tgl' in c.lower():
                date_col = c
                break

    # 2) Temukan kolom kode/kolom satker: prioritas 'Satker', 'Kode Satker', 'Nama Satker'
    kode_col = None
    if 'Kode Satker' in df.columns:
        kode_col = 'Kode Satker'
    else:
        for c in ['Satker','Nama Satker','Nama','Satker Nama','No']:
            if c in df.columns:
                kode_col = c
                break
        # fallback: cari kolom yang mengandung 'satker' di nama
        if kode_col is None:
            for c in df.columns:
                if 'satker' in c.lower():
                    kode_col = c
                    break

    # 3) Pastikan kolom kode sebagai string, dan buat kolom 'Kode Satker' standar
    df_work = df.copy()

    if kode_col is not None:
        # Some satker columns may contain "001234 - NAMA" -> use extraction
        df_work['Kode Satker'] = df_work[kode_col].astype(str).fillna('').apply(lambda s: extract_kode_from_satker_field(s))
    else:
        # if no satker-like column, try to find any column with many digits
        found = None
        for c in df_work.columns:
            sample = df_work[c].dropna().astype(str).head(10).tolist()
            if sample and all(re.search(r'\d', x) for x in sample):
                found = c
                break
        if found:
            df_work['Kode Satker'] = df_work[found].astype(str).fillna('').apply(lambda s: extract_kode_from_satker_field(s))
        else:
            df_work['Kode Satker'] = ''

    # 4) Parse tanggal revisi column (if exists) -> create 'Tanggal Posting Revisi' normalized
    if date_col is not None and date_col in df_work.columns:
        # try robust parsing
        def parse_date_safe(x):
            if pd.isna(x) or str(x).strip() == '':
                return pd.NaT
            # if already Timestamp
            if isinstance(x, (pd.Timestamp, datetime)):
                return pd.to_datetime(x)
            s = str(x).strip()
            # try common formats
            for fmt in ("%Y-%m-%d","%d-%m-%Y","%d/%m/%Y","%Y/%m/%d","%d %b %Y","%d %B %Y"):
                try:
                    return pd.to_datetime(s, format=fmt)
                except Exception:
                    pass
            # fallback to pandas parser
            try:
                return pd.to_datetime(s, dayfirst=True, errors='coerce')
            except Exception:
                return pd.NaT
        df_work['Tanggal Posting Revisi'] = df_work[date_col].apply(parse_date_safe)
    else:
        # no date column found -> create NaT
        df_work['Tanggal Posting Revisi'] = pd.NaT

    # 5) Determine Year: prefer explicit 'Tahun' column if present, else take from Tanggal Posting Revisi
    if 'Tahun' in df_work.columns:
        # coerce to int where possible, else infer from date
        def year_from_cell(x, fallback_dt):
            try:
                y = int(str(x).strip())
                if 1900 < y < 3000:
                    return int(y)
            except Exception:
                pass
            if not pd.isna(fallback_dt):
                return int(fallback_dt.year)
            return None
        df_work['Tahun'] = df_work.apply(lambda r: year_from_cell(r.get('Tahun', ''), r['Tanggal Posting Revisi']), axis=1)
    else:
        df_work['Tahun'] = df_work['Tanggal Posting Revisi'].apply(lambda d: int(d.year) if not pd.isna(d) else None)

    # 6) Normalize Kode Satker padding
    df_work['Kode Satker'] = df_work['Kode Satker'].apply(lambda x: normalize_kode_satker(x))

    # 7) For safety, keep original columns (but ensure date col parsed)
    # 8) Select latest revision per Kode Satker per Tahun (groupby)
    # Only keep rows where Tahun not None
    df_valid = df_work[df_work['Tahun'].notna()].copy()
    if df_valid.empty:
        # fallback: return empty df with standardized cols
        return df_work

    # If Tanggal Posting Revisi is all NaT, try to group by Kode Satker and take last occurrence
    if df_valid['Tanggal Posting Revisi'].isna().all():
        # take last occurrence per (Kode Satker, Tahun) keeping last by index
        df_valid = df_valid.sort_index()
        df_latest = df_valid.groupby(['Kode Satker','Tahun'], as_index=False).last()
    else:
        df_valid = df_valid.sort_values(by=['Tanggal Posting Revisi'])
        df_latest = df_valid.groupby(['Kode Satker','Tahun'], as_index=False).last()

    # Ensure result columns include core fields
    core_cols = ['Kode Satker','Tahun','Tanggal Posting Revisi']
    # bring core cols first then others
    other_cols = [c for c in df_latest.columns if c not in core_cols]
    df_latest = df_latest[core_cols + other_cols]

    # add source info
    if source_name:
        df_latest['_source_file'] = source_name

    return df_latest


# ============================================================
# 🔧 FUNGSI HELPER: Load Data DIPA dari GitHub
# ============================================================
def load_data_dipa_from_github():
    """
    Load semua file DIPA dari folder manapun di root repo yang mengandung nama 'dipa'.
    File valid: DIPA_2022.xlsx, DIPA_2023.xlsx, DIPA_2024.xlsx, DIPA_2025.xlsx, dll.
    """
    try:
        token = st.secrets.get("GITHUB_TOKEN")
        repo_name = st.secrets.get("GITHUB_REPO")

        if not token or not repo_name:
            st.warning("⚠️ GitHub credentials tidak ditemukan untuk load DIPA")
            return

        g = Github(auth=Auth.Token(token))
        repo = g.get_repo(repo_name)

        # FIX UTAMA: membaca root repo harus string kosong ""
        root_items = repo.get_contents("")

        # Cari folder yang mengandung kata 'dipa'
        dipa_folder = None
        for item in root_items:
            if item.type == "dir" and "dipa" in item.name.lower():
                dipa_folder = item.name  # contoh: "DATA_DIPA"
                break

        if not dipa_folder:
            st.warning("⚠️ Folder DIPA tidak ditemukan di GitHub.")
            return

        # Ambil isi folder DATA_DIPA
        contents = repo.get_contents(dipa_folder)
        if not isinstance(contents, list):
            contents = [contents]

        # Siapkan storage
        if "data_dipa_by_year" not in st.session_state:
            st.session_state.data_dipa_by_year = {}

        loaded_count = 0

        # Proses setiap file DIPA_xxxx.xlsx
        for content_file in contents:
            if content_file.type == "file" and content_file.name.lower().endswith(('.xlsx', '.xls')):
                filename = content_file.name

                # Extract tahun dari nama file
                year_match = re.search(r'dipa[_\-]?(\d{4})', filename.lower())
                if not year_match:
                    continue

                year = int(year_match.group(1))

                # Download file
                file_content = repo.get_contents(content_file.path)
                file_data = base64.b64decode(file_content.content)

                # Baca Excel
                df = pd.read_excel(io.BytesIO(file_data), dtype=str)

                # Normalisasi kode satker
                if "Kode Satker" in df.columns:
                    df["Kode Satker"] = df["Kode Satker"].apply(lambda x: normalize_kode_satker(str(x)))
                else:
                    df["Kode Satker"] = ""

                # ✅ Pastikan unik per satker (revisi terbaru)
                df = ensure_unique_dipa_per_satker(df)

                # Simpan
                st.session_state.data_dipa_by_year[year] = df
                loaded_count += 1

        if loaded_count > 0:
            years_loaded = sorted(st.session_state.data_dipa_by_year.keys())
            st.success(f"✅ Berhasil load {loaded_count} file DIPA: {', '.join(map(str, years_loaded))}")

    except Exception as e:
        st.error(f"❌ Error saat load data DIPA dari GitHub: {e}")


# Fungsi untuk memproses file Excel
def process_excel_file(uploaded_file, year):
    """
    Memproses file Excel IKPA sesuai struktur yang telah ditentukan
    """
    try:
        df_raw = pd.read_excel(uploaded_file, header=None)
        
        month = None   # tidak ambil bulan dari isi file karena tidak konsisten
        
        # 2️⃣ Ekstrak data (baris ke-5 dst)
        df_data = df_raw.iloc[4:].reset_index(drop=True)
        df_data.columns = range(len(df_data.columns))
        
        processed_rows = []
        i = 0
        while i < len(df_data):
            if i + 3 >= len(df_data):
                break
            
            nilai_row = df_data.iloc[i]
            bobot_row = df_data.iloc[i + 1]
            nilai_akhir_row = df_data.iloc[i + 2]
            nilai_aspek_row = df_data.iloc[i + 3]
            
            # Ekstrak kolom
            no = nilai_row[0]
            kode_kppn = str(nilai_row[1]).strip("'") if pd.notna(nilai_row[1]) else ""
            kode_ba = str(nilai_row[2]).strip("'") if pd.notna(nilai_row[2]) else ""
            kode_satker = str(nilai_row[3]).strip("'") if pd.notna(nilai_row[3]) else ""
            uraian_satker = nilai_row[4] if pd.notna(nilai_row[4]) else ""
            
            aspek_perencanaan = nilai_aspek_row[6] if pd.notna(nilai_aspek_row[6]) else 0
            aspek_pelaksanaan = nilai_aspek_row[8] if pd.notna(nilai_aspek_row[8]) else 0
            aspek_hasil = nilai_aspek_row[12] if pd.notna(nilai_aspek_row[12]) else 0
            
            revisi_dipa = nilai_row[6] if pd.notna(nilai_row[6]) else 0
            deviasi_hal3 = nilai_row[7] if pd.notna(nilai_row[7]) else 0
            penyerapan = nilai_row[8] if pd.notna(nilai_row[8]) else 0
            belanja_kontraktual = nilai_row[9] if pd.notna(nilai_row[9]) else 0
            penyelesaian_tagihan = nilai_row[10] if pd.notna(nilai_row[10]) else 0
            pengelolaan_up = nilai_row[11] if pd.notna(nilai_row[11]) else 0
            capaian_output = nilai_row[12] if pd.notna(nilai_row[12]) else 0
            
            nilai_total = nilai_row[13] if pd.notna(nilai_row[13]) else 0
            konversi_bobot = nilai_row[14] if pd.notna(nilai_row[14]) else 0
            dispensasi_spm = nilai_row[15] if pd.notna(nilai_row[15]) else 0
            nilai_akhir = nilai_row[16] if pd.notna(nilai_row[16]) else 0

            # Simpan bobot & nilai terbobot
            bobot_dict = {
                'Revisi DIPA': bobot_row[6], 'Deviasi Halaman III DIPA': bobot_row[7],
                'Penyerapan Anggaran': bobot_row[8], 'Belanja Kontraktual': bobot_row[9],
                'Penyelesaian Tagihan': bobot_row[10], 'Pengelolaan UP dan TUP': bobot_row[11],
                'Capaian Output': bobot_row[12]
            }
            nilai_terbobot_dict = {
                'Revisi DIPA': nilai_akhir_row[6], 'Deviasi Halaman III DIPA': nilai_akhir_row[7],
                'Penyerapan Anggaran': nilai_akhir_row[8], 'Belanja Kontraktual': nilai_akhir_row[9],
                'Penyelesaian Tagihan': nilai_akhir_row[10], 'Pengelolaan UP dan TUP': nilai_akhir_row[11],
                'Capaian Output': nilai_akhir_row[12]
            }

            row_data = {
                'No': no, 'Kode KPPN': kode_kppn, 'Kode BA': kode_ba, 'Kode Satker': kode_satker,
                'Uraian Satker': uraian_satker,
                'Kualitas Perencanaan Anggaran': aspek_perencanaan,
                'Kualitas Pelaksanaan Anggaran': aspek_pelaksanaan,
                'Kualitas Hasil Pelaksanaan Anggaran': aspek_hasil,
                'Revisi DIPA': revisi_dipa, 'Deviasi Halaman III DIPA': deviasi_hal3,
                'Penyerapan Anggaran': penyerapan, 'Belanja Kontraktual': belanja_kontraktual,
                'Penyelesaian Tagihan': penyelesaian_tagihan, 'Pengelolaan UP dan TUP': pengelolaan_up,
                'Capaian Output': capaian_output,
                'Nilai Total': nilai_total, 'Konversi Bobot': konversi_bobot,
                'Dispensasi SPM (Pengurang)': dispensasi_spm,
                'Nilai Akhir (Nilai Total/Konversi Bobot)': nilai_akhir,
                'Bulan': None,  
                'Tahun': year,
                'Bobot': bobot_dict, 'Nilai Terbobot': nilai_terbobot_dict
            }
            processed_rows.append(row_data)
            i += 4

        df_processed = pd.DataFrame(processed_rows)
        df_processed = df_processed.sort_values('Nilai Akhir (Nilai Total/Konversi Bobot)', ascending=False)
        df_processed['Peringkat'] = range(1, len(df_processed) + 1)

        # Apply reference short names (if available)
        df_processed = apply_reference_short_names(df_processed)
        df_processed = create_satker_column(df_processed)  # Use helper function
        df_processed['Source'] = 'Upload'

        return df_processed, month, year

    except Exception as e:
        st.error(f"Error memproses file: {str(e)}")
        return None, None, None


# Save any file (Excel/template) to your GitHub repo
def save_file_to_github(file_bytes, filename, folder="data"):
    token = st.secrets.get("GITHUB_TOKEN")
    repo_name = st.secrets.get("GITHUB_REPO")

    if not token or not repo_name:
        st.stop()
        st.error("❌ Gagal mengakses GitHub: GITHUB_TOKEN atau GITHUB_REPO tidak ditemukan di secrets.")
        return

    g = Github(auth=Auth.Token(token))
    repo = g.get_repo(repo_name)
    path = f"{folder}/{filename}"

    try:
        contents = repo.get_contents(path)
        repo.update_file(contents.path, f"Update {filename}", file_bytes, contents.sha)
        st.success(f"✅ File {filename} diperbarui di GitHub.")
    except Exception:
        repo.create_file(path, f"Upload {filename}", file_bytes)
        st.success(f"✅ File {filename} diunggah ke GitHub.")


# ============================
#  LOAD DATA IKPA DARI GITHUB
# ============================
def load_data_from_github():
    """
    Load all IKPA Excel files from GitHub /data folder.
    Filename format expected: IKPA_<BULAN>_<TAHUN>.xlsx
    """

    # ---------------------------
    # Validasi token GitHub
    # ---------------------------
    token = st.secrets.get("GITHUB_TOKEN")
    repo_name = st.secrets.get("GITHUB_REPO")

    if not token or not repo_name:
        st.error("❌ Gagal mengakses GitHub: GITHUB_TOKEN atau GITHUB_REPO tidak ditemukan.")
        st.stop()
        return

    g = Github(auth=Auth.Token(token))
    repo = g.get_repo(repo_name)

    # ---------------------------
    # Ambil list file di folder /data
    # ---------------------------
    try:
        contents = repo.get_contents("data")
    except Exception:
        st.info("📁 Folder 'data' belum ada di repository GitHub.")
        return

    if "data_storage" not in st.session_state:
        st.session_state.data_storage = {}
    else:
        st.session_state.data_storage.clear()

    # ================================
    # Helper: Hapus kolom Unnamed
    # ================================
    def clean_unnamed(df):
        return df.loc[:, ~df.columns.str.contains('^Unnamed')]

    # ================================
    # Proses setiap file Excel
    # ================================
    for file in contents:
        if not file.name.endswith(".xlsx"):
            continue

        # ---------------------------
        # Baca file Excel
        # ---------------------------
        decoded = base64.b64decode(file.content)
        df = pd.read_excel(io.BytesIO(decoded))
        df = clean_unnamed(df)

        # ---------------------------
        # PARSE NAMA FILE
        # IKPA_JANUARI_2025.xlsx
        # ---------------------------
        name = file.name.replace("IKPA_", "").replace(".xlsx", "")
        parts = name.split("_")

        if len(parts) < 2:
            continue

        year = parts[-1]
        month = "_".join(parts[:-1]).upper().replace(" ", "")

        # Standardisasi bulan (misal "PEBRUARI" → "FEBRUARI")
        MONTH_FIX = {
            "PEBRUARI": "FEBRUARI",
            "PEBRUARY": "FEBRUARI",
            "OKT": "OKTOBER",
        }
        month = MONTH_FIX.get(month, month)

        month_num = MONTH_ORDER.get(month.upper(), 0)

        # ---------------------------
        # Pastikan kolom Bulan & Tahun exist
        # ---------------------------
        df["Bulan"] = df.get("Bulan", month)
        df["Tahun"] = df.get("Tahun", year)

        df["Bulan"] = df["Bulan"].astype(str).str.upper()
        df["Tahun"] = df["Tahun"].astype(str)

        # ---------------------------
        # Normalisasi Kode Satker
        # ---------------------------
        if "Kode Satker" in df.columns:
            df["Kode Satker"] = df["Kode Satker"].astype(str).apply(normalize_kode_satker)
        else:
            df["Kode Satker"] = ""

        # ---------------------------
        # Reference Names (aman jika fungsi ada)
        # ---------------------------
        try:
            df = apply_reference_short_names(df)
        except:
            pass

        try:
            df = create_satker_column(df)
        except:
            pass

        # ---------------------------
        # Numeric columns
        # ---------------------------
        numeric_cols = [
            "Nilai Akhir (Nilai Total/Konversi Bobot)",
            "Nilai Total", "Konversi Bobot",
            "Revisi DIPA", "Deviasi Halaman III DIPA",
            "Penyerapan Anggaran",
            "Belanja Kontraktual", "Penyelesaian Tagihan",
            "Pengelolaan UP dan TUP", "Capaian Output"
        ]
        for col in numeric_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

        # ---------------------------
        # Date columns
        # ---------------------------
        if "Tanggal Posting Revisi" in df.columns:
            df["Tanggal Posting Revisi"] = pd.to_datetime(df["Tanggal Posting Revisi"], errors="coerce")

        # ---------------------------
        # Tambahan kolom helper
        # ---------------------------
        df["Source"] = "GitHub"
        df["Period"] = f"{month} {year}"
        df["Period_Sort"] = f"{int(year):04d}-{month_num:02d}"

        # ---------------------------
        # Generate ranking jika belum ada
        # ---------------------------
        if "Peringkat" not in df.columns and "Nilai Akhir (Nilai Total/Konversi Bobot)" in df.columns:
            df = df.sort_values("Nilai Akhir (Nilai Total/Konversi Bobot)", ascending=False)
            df["Peringkat"] = range(1, len(df) + 1)

        # ---------------------------
        # Simpan ke session_state
        # ---------------------------
        st.session_state.data_storage[(month, year)] = df

    st.success(f"✅ {len(st.session_state.data_storage)} file berhasil dimuat dari GitHub.")


# ============================
#  BACA TEMPLATE FILE
# ============================
def get_template_file():
    try:
        if Path(TEMPLATE_PATH).exists():
            with open(TEMPLATE_PATH, "rb") as f:
                return f.read()
        else:
            if "template_file" in st.session_state:
                return st.session_state.template_file
            return None
    except Exception as e:
        st.error(f"Error membaca template: {e}")
        return None

# Fungsi visualisasi podium/bintang
def create_ranking_chart(df, title, top=True, limit=10):
    """
    Membuat visualisasi ranking dengan bar chart horizontal yang menarik
    (Sekarang menggunakan kolom 'Satker' untuk label agar unik)
    """
    if top:
        df_sorted = df.nlargest(limit, 'Nilai Akhir (Nilai Total/Konversi Bobot)')
        color_scale = 'Greens'
        emoji = '🏆'
    else:
        df_sorted = df.nsmallest(limit, 'Nilai Akhir (Nilai Total/Konversi Bobot)')
        color_scale = 'Reds'
        emoji = '⚠️'
    
    fig = go.Figure()
    
    colors = px.colors.sequential.Greens if top else px.colors.sequential.Reds
    
    # use 'Satker' for y labels to keep them unique
    fig.add_trace(go.Bar(
    y=df_filtered['Satker'],
    x=df_filtered[column],
    orientation='h',
    marker=dict(
        color=df_filtered[column],
        colorscale='OrRd_r',
        showscale=True,
        cmin=min_val,
        cmax=max_val,
    ),
    text=df_filtered[column].round(2),
    textposition='outside',
    hovertemplate='<b>%{y}</b><br>Nilai: %{x:.2f}<extra></extra>'
))
    
    fig.update_layout(
        title=f"{emoji} {title}",
        xaxis_title="Nilai Akhir",
        yaxis_title="",
        height=max(400, limit * 40),
        yaxis={'categoryorder': 'total ascending' if not top else 'total descending'},
        showlegend=False
    )
    # ============================
    # Rotated labels 45° di bawah
    # ============================
    annotations = []
    y_positions = list(range(len(df_filtered)))

    for i, satker in enumerate(df_filtered['Satker']):
        annotations.append(dict(
        x=df_filtered[column].min() - 3,
        y=i,
        text=satker,
        xanchor="right",
        yanchor="middle",
        showarrow=False,
        textangle=45,
        font=dict(size=10),
    ))

    fig.update_layout(annotations=annotations)

    # Sembunyikan label Y-axis
    fig.update_yaxes(showticklabels=False)

    return fig

# ============================================================
# Improved Problem Chart (with sorting, sliders, and filters)
# ============================================================
def make_column_chart(data, title, color_scale, y_min, y_max, limit=10, show_yaxis=False):
    df_top = data.nlargest(limit, "Nilai Akhir (Nilai Total/Konversi Bobot)")
    fig = px.bar(
        df_top,
        x="Nilai Akhir (Nilai Total/Konversi Bobot)",
        y="Satker",
        orientation="h",
        color="Nilai Akhir (Nilai Total/Konversi Bobot)",
        color_continuous_scale=color_scale,
        title=title
    )

    fig.update_layout(
        xaxis_range=[y_min, y_max],
        yaxis_title="",
        xaxis_title="Nilai IKPA",
        height=500,
        margin=dict(l=10, r=10, t=40, b=20),
        coloraxis_showscale=False,
        showlegend=False
    )

    fig.update_traces(
        texttemplate="%{x:.2f}",  
        textposition="outside",
        hovertemplate="<b>%{y}</b><br>Nilai: %{x:.2f}<extra></extra>"
    )

    return fig

# ============================================================
# Problem Chart untuk Dashboard Internal
# ============================================================
def create_problem_chart(df, column, threshold, title, comparison='less', y_min=None, y_max=None, show_yaxis=True):

    if comparison == 'less':
        df_filtered = df[df[column] < threshold]
    elif comparison == 'greater':
        df_filtered = df[df[column] > threshold]
    else:
        df_filtered = df.copy()

    # Jika hasil filter kosong → Cegah error
    if df_filtered.empty:
        df_filtered = df.head(1)

    df_filtered = df_filtered.sort_values(by=column, ascending=False)

    # Ambil nilai range untuk colormap
    min_val = df_filtered[column].min()
    max_val = df_filtered[column].max()

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=df_filtered['Satker'],
        y=df_filtered[column],
        marker=dict(
            color=df_filtered[column],
            colorscale='OrRd_r',
            showscale=True,
            cmin=min_val,
            cmax=max_val,
        ),
        text=df_filtered[column].round(2),
        textposition='outside',
        textangle=0,
        textfont=dict(family="Arial Black", size=12), 
        hovertemplate='<b>%{x}</b><br>Nilai: %{y:.2f}<extra></extra>'
    ))

    # Garis target threshold (tidak berubah)
    fig.add_hline(
        y=threshold,
        line_dash="dash",
        line_color="red",
        annotation_text=f"Target: {threshold}",
        annotation_position="top right"
    )

    # Bold judul dan label axis
    fig.update_layout(
        xaxis=dict(
        tickangle=-45,
        tickmode='linear',
        tickfont=dict(family="Arial Black", size=10),
        automargin=True
    ),
    yaxis=dict(
        tickfont=dict(family="Arial Black", size=11)
        ),
        height=600,
        margin=dict(l=50, r=20, t=80, b=200),
        showlegend=False,
    )

    if not show_yaxis:
        fig.update_yaxes(showticklabels=False)

    return fig
# ===============================================
# Helper to apply reference short names (Simplified)
# ===============================================
def apply_reference_short_names(df):
    """
    Simple version: apply reference short names to dataframe.
    - Adds 'Uraian Satker-RINGKAS' (from reference 'Uraian Satker-SINGKAT' when available,
      otherwise falls back to original 'Uraian Satker').
    - Performs basic normalization on 'Kode Satker' before merging.
    - Minimal user messages (no Excel/CSV creation, no verbose debugging).
    """
    # Defensive copy
    df = df.copy()

    # Ensure period columns exist
    if 'Bulan' not in df.columns:
        df['Bulan'] = ''
    if 'Tahun' not in df.columns:
        df['Tahun'] = ''

    # If no reference in session, fallback silently to original names
    if 'reference_df' not in st.session_state or st.session_state.reference_df is None:
        if 'Uraian Satker-RINGKAS' not in df.columns:
            df['Uraian Satker-RINGKAS'] = df.get('Uraian Satker', '')
        # also keep a final fallback column for compatibility
        df['Uraian Satker Final'] = df.get('Uraian Satker', '')
        return df

    # Copy reference
    ref = st.session_state.reference_df.copy()

    # Normalize Kode Satker if column exists; else create empty codes to avoid crashes
    if 'Kode Satker' in df.columns:
        df['Kode Satker'] = df['Kode Satker'].apply(normalize_kode_satker)
    else:
        df['Kode Satker'] = ''

    if 'Kode Satker' in ref.columns:
        ref['Kode Satker'] = ref['Kode Satker'].apply(normalize_kode_satker)
    else:
        # If reference has no Kode Satker, cannot match — fallback
        if 'Uraian Satker-RINGKAS' not in df.columns:
            df['Uraian Satker-RINGKAS'] = df.get('Uraian Satker', '')
        df['Uraian Satker Final'] = df.get('Uraian Satker', '')
        return df

    # Ensure kode fields are strings and stripped
    df['Kode Satker'] = df['Kode Satker'].astype(str).str.strip()
    ref['Kode Satker'] = ref['Kode Satker'].astype(str).str.strip()

    # If the reference does not contain the expected short-name column, fallback
    if 'Uraian Satker-SINGKAT' not in ref.columns:
        if 'Uraian Satker-RINGKAS' not in df.columns:
            df['Uraian Satker-RINGKAS'] = df.get('Uraian Satker', '')
        df['Uraian Satker Final'] = df.get('Uraian Satker', '')
        return df

    # Perform the merge and create final short-name column; keep it simple and robust
    try:
        df_merged = df.merge(
            ref[['Kode Satker', 'Uraian Satker-SINGKAT']].rename(columns={'Uraian Satker-SINGKAT': 'Uraian Satker-RINGKAS'}),
            on='Kode Satker',
            how='left',
            indicator=False
        )

        # Create final name column using reference when available, otherwise fallback to original
        df_merged['Uraian Satker-RINGKAS'] = df_merged['Uraian Satker-RINGKAS'].fillna(
            df_merged.get('Uraian Satker', '')
        )

        # Keep a generic final field for backward compatibility
        df_merged['Uraian Satker Final'] = df_merged['Uraian Satker-RINGKAS']

        # Drop the reference short-name column in case it remains under other names
        df_merged = df_merged.drop(columns=['Uraian Satker-SINGKAT'], errors='ignore')

        return df_merged

    except Exception as e:
        # Minimal error notification and fallback
        st.error(f"❌ Gagal menerapkan nama singkat untuk periode {df.get('Bulan', [''])[0]} {df.get('Tahun', [''])[0]}: {e}")
        if 'Uraian Satker-RINGKAS' not in df.columns:
            df['Uraian Satker-RINGKAS'] = df.get('Uraian Satker', '')
        df['Uraian Satker Final'] = df.get('Uraian Satker', '')
        return df

# ===============================================
# UPDATED: Helper function to create Satker column consistently
# ===============================================
def create_satker_column(df):
    """
    Creates 'Satker' column consistently across all data sources.
    Should be called after apply_reference_short_names().
    """
    if 'Uraian Satker-RINGKAS' not in df.columns:
        # fallback to older field names
        if 'Uraian Satker Final' in df.columns:
            df['Uraian Satker-RINGKAS'] = df['Uraian Satker Final']
        else:
            df['Uraian Satker-RINGKAS'] = df.get('Uraian Satker', '')

    # Create Satker display using ringkas
    df['Satker'] = (
        df['Uraian Satker-RINGKAS'].astype(str) + 
        ' (' + df['Kode Satker'].astype(str) + ')'
    )
    # Keep backward compatible column
    df['Uraian Satker Final'] = df['Uraian Satker-RINGKAS']
    return df
    
# HALAMAN 1: DASHBOARD UTAMA (REVISED)
def page_dashboard():
    st.title("📊 Dashboard Utama IKPA Satker Mitra KPPN Baturaja")
    
    st.markdown("""
    <style>
    /* Warna tombol popover */
    div[data-testid="stPopover"] button {
        background-color: #FFF9E6 !important;
        border: 1px solid #E6C200 !important;
        color: #664400 !important;
    }
    div[data-testid="stPopover"] button:hover {
        background-color: #FFE4B5 !important;
        color: black !important;
    }
    button[data-testid="baseButton"][kind="popover"] {
        background-color: #FFF9E6 !important;
        border: 1px solid #E6C200 !important;
        color: #664400 !important;
    }
    button[data-testid="baseButton"][kind="popover"]:hover {
        background-color: #FFE4B5 !important;
    }
    </style>
    """, unsafe_allow_html=True)

    # protect against missing data_storage
    if not st.session_state.get('data_storage'):
        st.warning("⚠️ Belum ada data yang diunggah. Silakan unggah data melalui halaman Admin.")
        return

    # Dapatkan data terbaru
    try:
        all_periods = sorted(
            st.session_state.data_storage.keys(),
            key=lambda x: (int(x[1]), MONTH_ORDER.get(x[0].upper(), 0)),
            reverse=True
        )
    except Exception:
        st.warning("⚠️ Format periode pada data tidak sesuai. Periksa struktur data di session_state.data_storage.")
        return

    if not all_periods:
        st.warning("⚠️ Belum ada data yang tersedia.")
        return

    # ---------------------------
    # Ensure a selected_period and df exist BEFORE any branch uses df
    # ---------------------------
    if "selected_period" not in st.session_state:
        st.session_state.selected_period = all_periods[0]

    # safe fetch of df for the selected_period (always a tuple key like ('JANUARI','2025'))
    selected_period_key = st.session_state.get("selected_period", all_periods[0])
    df = st.session_state.data_storage.get(selected_period_key, None)

    if df is None:
        st.warning(f"⚠️ Data untuk periode {selected_period_key} tidak ditemukan. Periksa st.session_state.data_storage keys.")
        # show available keys to help debugging (optional - remove if sensitive)
        st.write("Periode yang tersedia:", list(st.session_state.data_storage.keys()))
        return

    # ensure main_tab state exists
    if "main_tab" not in st.session_state:
        st.session_state.main_tab = "🎯 Highlights"

    # ---------- persistent main tab ----------
    main_tab = st.radio(
        "Pilih Bagian Dashboard",
        ["🎯 Highlights", "📋 Data Detail Satker"],
        key="main_tab_choice",
        horizontal=True
    )
    st.session_state["main_tab"] = main_tab

    # -------------------------
    # HIGHLIGHTS
    # -------------------------
    if main_tab == "🎯 Highlights":
        st.markdown("## 🎯 Highlights Kinerja Satker")

        # Single-row layout for period + metrics
        col_period, col1, col2, col3, col4 = st.columns([1, 1, 1, 1, 1])

        with col_period:
            # update selected_period in session_state when changed here
            st.session_state.selected_period = st.selectbox(
                "Pilih Periode",
                options=all_periods,
                index=0,
                format_func=lambda x: f"{x[0].capitalize()} {x[1]}",
                key="select_period_main"
            )
            # refresh df variable to reflect selection immediately (keeps consistency)
            selected_period_key = st.session_state.selected_period
            df = st.session_state.data_storage.get(selected_period_key, df)

        # now df is guaranteed to be set (we checked earlier)
        avg_score = df['Nilai Akhir (Nilai Total/Konversi Bobot)'].mean()
        perfect_df = df[df['Nilai Akhir (Nilai Total/Konversi Bobot)'] == 100]
        below89_df = df[df['Nilai Akhir (Nilai Total/Konversi Bobot)'] < 89]
        
        # Pastikan kolom Satker tersedia
        def make_satker_col(dd):
            if 'Satker' in dd.columns:
                return dd
            uraian = dd.get('Uraian Satker-RINGKAS', dd.index.astype(str))
            kode = dd.get('Kode Satker', '')
            dd['Satker'] = uraian.astype(str) + " (" + kode.astype(str) + ")"
            return dd

        perfect_df = make_satker_col(perfect_df)
        below89_df = make_satker_col(below89_df)

        # Hitung jumlah
        jumlah_100 = len(perfect_df)
        jumlah_below = len(below89_df)

        with col1:
            st.metric("📋 Total Satker", len(df))
        with col2:
            st.metric("📈 Rata-rata Nilai", f"{avg_score:.2f}")
        
        with col3:
            st.metric("⭐ Nilai 100", jumlah_100)
            with st.popover("Lihat daftar satker"):
                if jumlah_100 == 0:
                    st.write("Tidak ada satker dengan nilai 100.")
                else:
                    display_df = perfect_df[['Satker']].reset_index(drop=True)
                    display_df.insert(0, 'No', range(1, len(display_df) + 1))
                    st.dataframe(display_df, use_container_width=True, hide_index=True, height=min(400, len(display_df) * 35 + 38))
        with col4:
            st.metric("⚠️ Nilai < 89 (Predikat Belum Baik)", jumlah_below)
            with st.popover("Lihat daftar satker"):
                if jumlah_below == 0:
                    st.write("Tidak ada satker dengan nilai < 89.")
                else:
                    display_df = below89_df[['Satker']].reset_index(drop=True)
                    display_df.insert(0, 'No', range(1, len(display_df) + 1))
                    st.dataframe(display_df, use_container_width=True, hide_index=True, height=min(400, len(display_df) * 35 + 38))

        # Chart controls
        st.markdown("###### Atur Skala Nilai (Sumbu Y)")
        col_min, col_max = st.columns(2)
        with col_min:
            y_min = st.slider(
                "Nilai Minimum (Y-Axis)",
                min_value=0,
                max_value=50,
                value=50,
                step=1,
                key="high_ymin"
            )
        with col_max:
            y_max = st.slider(
                "Nilai Maksimum (Y-Axis)",
                min_value=51,
                max_value=110,
                value=110,
                step=1,
                key="high_ymax"
            )

        # Data preparation for charts
        df_with_kontrak = df[df['Belanja Kontraktual'] != 0]
        df_without_kontrak = df[df['Belanja Kontraktual'] == 0]

        # 4 charts side-by-side
        col1, col2, col3, col4 = st.columns(4)

        with col1:
            st.markdown("##### 🏆 10 Satker Terbaik (Dengan Kontraktual)")
            if len(df_with_kontrak) > 0:
                top_with = df_with_kontrak.nlargest(10, 'Nilai Akhir (Nilai Total/Konversi Bobot)')
                fig1 = make_column_chart(top_with, "", "greens", y_min, y_max, show_yaxis=True)
                st.plotly_chart(fig1, use_container_width=True)
            else:
                st.info("Tidak ada data.")

        with col2:
            st.markdown("##### 🏆 10 Satker Terbaik (Tanpa Kontraktual)")
            if len(df_without_kontrak) > 0:
                top_without = df_without_kontrak.nlargest(10, 'Nilai Akhir (Nilai Total/Konversi Bobot)')
                fig2 = make_column_chart(top_without, "", "greens", y_min, y_max, show_yaxis=False)
                st.plotly_chart(fig2, use_container_width=True)
            else:
                st.info("Tidak ada data.")

        with col3:
            st.markdown("##### 📉 10 Satker Terendah (Dengan Kontraktual)")
            if len(df_with_kontrak) > 0:
                bottom_with = df_with_kontrak.nsmallest(10, 'Nilai Akhir (Nilai Total/Konversi Bobot)')
                fig3 = make_column_chart(bottom_with, "", "orrd_r", y_min, y_max, show_yaxis=False)
                st.plotly_chart(fig3, use_container_width=True)
            else:
                st.info("Tidak ada data.")

        with col4:
            st.markdown("##### 📉 10 Satker Terendah (Tanpa Kontraktual)")
            if len(df_without_kontrak) > 0:
                bottom_without = df_without_kontrak.nsmallest(10, 'Nilai Akhir (Nilai Total/Konversi Bobot)')
                fig4 = make_column_chart(bottom_without, "", "orrd_r", y_min, y_max, show_yaxis=False)
                st.plotly_chart(fig4, use_container_width=True)
            else:
                st.info("Tidak ada data.")

        st.markdown("---")

        # Satker dengan masalah (Deviasi Hal 3 DIPA)
        st.subheader("🚨 Satker yang Memerlukan Perhatian Khusus")
        st.markdown("###### Atur Skala Nilai (Sumbu Y)")
        col_min_dev, col_max_dev = st.columns(2)
        with col_min_dev:
            y_min_dev = st.slider(
                "Nilai Minimum (Y-Axis)",
                min_value=0,
                max_value=50,
                value=40,
                step=1,
                key="high_ymin_dev"
            )
        with col_max_dev:
            y_max_dev = st.slider(
                "Nilai Maksimum (Y-Axis)",
                min_value=51,
                max_value=110,
                value=110,
                step=1,
                key="high_ymax_dev"
            )

        fig_dev = create_problem_chart(
            df, 
            'Deviasi Halaman III DIPA', 
            90, 
            "Deviasi Hal 3 DIPA Belum Optimal (< 90)",
            'less',
            y_min=y_min_dev,
            y_max=y_max_dev,
            show_yaxis=True
        )
        if fig_dev:
            st.plotly_chart(fig_dev, use_container_width=True)
        else:
            st.success("✅ Semua satker sudah optimal untuk Deviasi Hal 3 DIPA")

    # -------------------------
    # DATA DETAIL SATKER
    # -------------------------
    else:
        st.subheader("📋 Tabel Detail Satker")

        # persistent sub-tab for Periodik / Detail Satker
        if "active_table_tab" not in st.session_state:
            st.session_state.active_table_tab = "📆 Periodik"

        sub_tab = st.radio(
            "Pilih Mode Tabel",
            ["📆 Periodik", "📋 Detail Satker"],
            key="sub_tab_choice",
            horizontal=True
        )
        st.session_state['active_table_tab'] = sub_tab

        # -------------------------
        # PERIODIK TABLE
        # -------------------------
        if sub_tab == "📆 Periodik":
            st.markdown("#### Periodik — ringkasan per bulan / triwulan / perbandingan")

            # Tentukan tahun yang tersedia
            years = set()
            for k, df_period in st.session_state.data_storage.items():
                years.update(df_period['Tahun'].astype(str).unique())
            years = sorted([int(y) for y in years if str(y).strip() != ''], reverse=True)

            if not years:
                st.info("Tidak ada data periodik untuk ditampilkan.")
                st.stop()

            default_year = years[0]
            selected_year = st.selectbox("Pilih Tahun", options=years, index=0, key='tab_periodik_year_select')

            # session state untuk period_type
            if "period_type" not in st.session_state:
                st.session_state.period_type = "quarterly"

            period_options = ["quarterly", "monthly", "compare"]
            try:
                period_index = period_options.index(st.session_state.period_type)
            except ValueError:
                period_index = 0
                st.session_state.period_type = "quarterly"

            # Radio button
            period_type = st.radio(
                "Jenis Periode",
                options=period_options,
                format_func=lambda x: {"quarterly": "Triwulan", "monthly": "Bulanan", "compare": "Perbandingan"}.get(x, x),
                horizontal=True,
                index=period_index,
                key="period_type_radio_v2"
            )
            st.session_state.period_type = period_type

            # Pilih indikator (satu untuk semua mode)
            indicator_options = [
                'Kualitas Perencanaan Anggaran', 'Kualitas Pelaksanaan Anggaran', 'Kualitas Hasil Pelaksanaan Anggaran',
                'Revisi DIPA', 'Deviasi Halaman III DIPA', 'Penyerapan Anggaran', 'Belanja Kontraktual',
                'Penyelesaian Tagihan', 'Pengelolaan UP dan TUP', 'Capaian Output', 'Dispensasi SPM (Pengurang)',
                'Nilai Akhir (Nilai Total/Konversi Bobot)'
            ]
            default_indicator = 'Deviasi Halaman III DIPA'
            selected_indicator = st.selectbox(
                "Pilih Indikator", 
                options=indicator_options, 
                index=indicator_options.index(default_indicator) if default_indicator in indicator_options else 0,
                key='tab_periodik_indicator_select'
            )
            
            # -------------------------
            # Monthly / Quarterly
            # -------------------------
            if period_type in ['monthly', 'quarterly']:
    
                # Gabungkan data berdasarkan tahun
                df_list = []
                for (mon, yr), df_period in st.session_state.data_storage.items():
                    if str(yr).strip() == str(selected_year).strip():
                        temp = df_period.copy()
                        temp["Bulan_raw"] = mon.upper()
                        temp["Bulan_upper"] = mon.upper()
                        df_list.append(temp)

                if not df_list:
                    st.info(f"Tidak ditemukan data untuk tahun {selected_year}.")
                    st.stop()

                df_year = pd.concat(df_list, ignore_index=True)
                
                # NORMALISASI NAMA BULAN
                MONTH_FIX = {
                    "JAN": "JANUARI", "JANUARY": "JANUARI",
                    "FEB": "FEBRUARI",
                    "MAR": "MARET", "MRT": "MARET",
                    "APR": "APRIL",
                    "AGT": "AGUSTUS", "AUG": "AGUSTUS",
                    "SEP": "SEPTEMBER", "SEPT": "SEPTEMBER",
                    "OKT": "OKTOBER", "OCT": "OKTOBER",
                    "DES": "DESEMBER", "DEC": "DESEMBER"
                }

                import re
                def normalize_month(b):
                    b = re.sub(r'[^A-Z]', '', str(b).upper())
                    return MONTH_FIX.get(b, b)

                df_year["Bulan_upper"] = df_year["Bulan_raw"].apply(normalize_month)

                months_available = sorted(
                    [m for m in df_year['Bulan_upper'].unique() if m],
                    key=lambda m: MONTH_ORDER.get(m, 999)
                )

                # =============================
                # 🔧 PERBAIKAN UTAMA: Pivot berdasarkan Kode Satker
                # =============================

                # 1. Buat kolom periode yang sesuai
                if period_type == 'monthly':
                    # gunakan BUKAN Bulan_upper, tetapi Bulan dari uploader
                    df_year['Period_Column'] = df_year['Bulan'].str.upper()

                else:  # quarterly
                    def map_to_quarter(month):
                        if month in ['MARET', 'MAR', 'MRT']:
                            return 'Tw I'
                        elif month == 'JUNI':
                            return 'Tw II'
                        elif month in ['SEPTEMBER', 'SEPT', 'SEP']:
                            return 'Tw III'
                        elif month == 'DESEMBER':
                            return 'Tw IV'
                        return None
                    
                    df_year['Period_Column'] = df_year['Bulan'].str.upper().apply(map_to_quarter)
                    df_year = df_year[df_year['Period_Column'].notna()]

                # 2. Ambil kolom yang diperlukan
                base_cols = ['Kode BA', 'Kode Satker', 'Uraian Satker-RINGKAS', 'Period_Column']
                df_pivot = df_year[base_cols + [selected_indicator]].copy()

                # 3. Groupby untuk menghindari duplikasi (ambil nilai terakhir per satker per periode)
                df_pivot = df_pivot.sort_values('Period_Column')
                df_pivot = df_pivot.groupby(
                    ['Kode BA', 'Kode Satker', 'Uraian Satker-RINGKAS', 'Period_Column'],
                    as_index=False
                ).last()

                # 4. Pivot tabel
                df_wide = df_pivot.pivot_table(
                    index=['Kode BA', 'Kode Satker', 'Uraian Satker-RINGKAS'],
                    columns='Period_Column',
                    values=selected_indicator,
                    aggfunc='last'  # ambil nilai terakhir jika ada duplikasi
                ).reset_index()

                # 5. Urutkan kolom periode
                if period_type == 'monthly':
                    ordered_periods = [m for m in months_available if m in df_wide.columns]
                else:
                    ordered_periods = [tw for tw in ['Tw I', 'Tw II', 'Tw III', 'Tw IV'] if tw in df_wide.columns]

                # 6. Susun ulang kolom
                final_cols = ['Kode BA', 'Kode Satker', 'Uraian Satker-RINGKAS'] + ordered_periods
                df_wide = df_wide[final_cols]

                # 7. Hitung peringkat berdasarkan periode terakhir
                if ordered_periods:
                    last_period = ordered_periods[-1]
                    df_wide['Latest_Value'] = df_wide[last_period]
                    df_wide['Peringkat'] = (
                        df_wide['Latest_Value']
                        .rank(ascending=False, method='dense')
                        .astype('Int64')
                    )
                else:
                    df_wide['Peringkat'] = None

                # 8. Urutkan berdasarkan peringkat
                df_wide = df_wide.sort_values('Peringkat', ascending=True)

                # 9. Susun kolom final untuk display
                display_cols = ['Peringkat', 'Kode BA', 'Kode Satker', 'Uraian Satker-RINGKAS'] + ordered_periods
                df_display = df_wide[display_cols].copy()

                # 10. Rename kolom periode untuk display yang lebih baik
                if period_type == 'monthly':
                    rename_dict = {m: m.capitalize() for m in ordered_periods}
                    df_display = df_display.rename(columns=rename_dict)
                    display_period_cols = [m.capitalize() for m in ordered_periods]
                else:
                    display_period_cols = ordered_periods

                # =============================
                # SEARCH & STYLING (sama seperti sebelumnya)
                # =============================
                search_query = st.text_input(
                    "🔎 Cari (Periodik) – ketik untuk filter di semua kolom",
                    value="",
                    key='tab_periodik_search'
                )

                if search_query:
                    q = str(search_query).strip().lower()
                    mask = df_display.apply(
                        lambda row: row.astype(str).str.lower().str.contains(q, na=False).any(),
                        axis=1
                    )
                    df_display_filtered = df_display[mask].copy()
                else:
                    df_display_filtered = df_display.copy()

                # Trend coloring
                def color_trend(row):
                    styles = []
                    vals = [row[c] for c in display_period_cols if pd.notna(row[c])]
                    if len(vals) >= 2:
                        if vals[-1] > vals[-2]:
                            color = 'background-color: #c6efce'
                        elif vals[-1] < vals[-2]:
                            color = 'background-color: #f8d7da'
                        else:
                            color = ''
                    else:
                        color = ''

                    for c in df_display_filtered.columns:
                        styles.append(color if (display_period_cols and c == display_period_cols[-1]) else '')
                    return styles

                def highlight_top(s):
                    if s.name == 'Peringkat':
                        return [
                            'background-color: gold' if (pd.to_numeric(v, errors='coerce') <= 3) else ''
                            for v in s
                        ]
                    return ['' for _ in s]

                styler = df_display_filtered.style.format(precision=2, na_rep='–')
                if display_period_cols:
                    styler = styler.apply(color_trend, axis=1)
                styler = styler.apply(highlight_top)

                st.dataframe(styler, use_container_width=True, height=600)


            elif period_type == "compare":
                st.markdown("### Perbandingan Antara Dua Tahun")

               # Gabungkan seluruh data
                all_data = []
                for (mon, yr), df in st.session_state.data_storage.items():
                    df2 = df.copy()
                    df2["Bulan_upper"] = df2["Bulan"].astype(str).str.upper().str.strip()
                    df2["Tahun"] = df2["Tahun"].astype(int)
                    all_data.append(df2)

                if not all_data:
                    st.warning("Belum ada data yang di-upload.")
                    st.stop()

                df_full = pd.concat(all_data, ignore_index=True)


                # Tahun yang valid
                available_years = sorted([y for y in df_full["Tahun"].unique() if 2022 <= y <= 2025])
                if len(available_years) < 2:
                    st.warning("Data tahun tidak cukup.")
                    st.stop()

                # Pilih Tahun A dan B
                colA, colB = st.columns(2)
                with colA:
                    year_a = st.selectbox("Tahun A (Awal)", available_years, index=0, key="tahunA_compare")
                with colB:
                    year_b = st.selectbox("Tahun B (Akhir)", available_years, index=1, key="tahunB_compare")

                if year_a == year_b:
                    st.info("Pilih dua tahun yang berbeda.")
                    st.stop()

                # Filter tahun
                df_a = df_full[df_full["Tahun"] == year_a]
                df_b = df_full[df_full["Tahun"] == year_b]

                def extract_tw(df_):
                    return {
                        "Tw I": df_[df_["Bulan_upper"] == "MARET"],
                        "Tw II": df_[df_["Bulan_upper"] == "JUNI"],
                        "Tw III": df_[df_["Bulan_upper"] == "SEPTEMBER"],
                        "Tw IV": df_[df_["Bulan_upper"] == "DESEMBER"],
                    }

                tw_a = extract_tw(df_a)
                tw_b = extract_tw(df_b)

                # Pilihan Satker
                satker_list = df_full[['Kode Satker', 'Uraian Satker-RINGKAS']].drop_duplicates()
                satker_options = ["SEMUA SATKER"] + satker_list['Kode Satker'].tolist()

                selected_satkers = st.multiselect(
                    "Pilih Satker",
                    satker_options,
                    format_func=lambda x: (
                        "SEMUA SATKER" if x == "SEMUA SATKER"
                        else satker_list[satker_list['Kode Satker'] == x]['Uraian Satker-RINGKAS'].values[0]
                    ),
                    default=["SEMUA SATKER"],
                    key="satker_compare"
                )

                if "SEMUA SATKER" in selected_satkers:
                    selected_satkers_final = satker_list['Kode Satker'].tolist()
                else:
                    selected_satkers_final = selected_satkers

                # Build tabel
                rows = []
                for _, m in satker_list.iterrows():
                    kode = m['Kode Satker']
                    if kode not in selected_satkers_final:
                        continue

                    nama = m['Uraian Satker-RINGKAS']
                    row = {"Kode Satker": kode, "Uraian Satker": nama}

                    latest_a = None
                    latest_b = None

                    for tw in ['Tw I', 'Tw II', 'Tw III', 'Tw IV']:
                        # Tahun A
                        valA = tw_a[tw][tw_a[tw]['Kode Satker'] == kode][selected_indicator].values
                        valA = valA[0] if len(valA) else None
                        row[f"{tw} {year_a}"] = valA
                        if valA is not None:
                            latest_a = valA

                        # Tahun B
                        valB = tw_b[tw][tw_b[tw]['Kode Satker'] == kode][selected_indicator].values
                        valB = valB[0] if len(valB) else None
                        row[f"{tw} {year_b}"] = valB
                        if valB is not None:
                            latest_b = valB

                    # Selisih
                    if latest_a is not None and latest_b is not None:
                        row[f"Δ Total ({year_b}-{year_a})"] = latest_b - latest_a
                    else:
                        row[f"Δ Total ({year_b}-{year_a})"] = None

                    rows.append(row)

                df_compare = pd.DataFrame(rows)

                # Styling warna
                def highlight_years(col_name):
                    if str(year_a) in col_name:
                        return 'background-color: #FFF8C6;'  # kuning muda
                    if str(year_b) in col_name:
                        return 'background-color: #DCEBFF;'  # biru muda
                    return ''

                df_style = df_compare.style.apply(
                    lambda row: [highlight_years(col) for col in df_compare.columns],
                    axis=1
                ).format(precision=2)

                st.markdown("### Hasil Perbandingan")
                st.dataframe(df_style, use_container_width=True, height=600)

        # -------------------------
        # DETAIL SATKER (legacy table)
        # -------------------------
        else:
            # ensure df available (use selected period if set)
            df = st.session_state.data_storage.get(st.session_state.get('selected_period', all_periods[0]), None)
            if df is None:
                st.info("Data untuk detail satker tidak tersedia untuk periode yang dipilih.")
                return

            col1, col2 = st.columns([2, 1])
            with col1:
                view_mode = st.radio(
                    "Tampilan",
                    options=['aspek', 'komponen'],
                    format_func=lambda x: 'Berdasarkan Aspek' if x == 'aspek' else 'Berdasarkan Komponen',
                    horizontal=True
                )
            with col2:
                st.write("")

            display_columns = ['Peringkat', 'Kode BA', 'Kode Satker', 'Uraian Satker-RINGKAS']
            if view_mode == 'aspek':
                display_columns += [
                    'Kualitas Perencanaan Anggaran',
                    'Kualitas Pelaksanaan Anggaran',
                    'Kualitas Hasil Pelaksanaan Anggaran'
                ]
                df_display = df[display_columns + ['Nilai Total',
                                                   'Dispensasi SPM (Pengurang)',
                                                   'Nilai Akhir (Nilai Total/Konversi Bobot)']].copy()
            else:
                component_cols = [
                    'Revisi DIPA', 'Deviasi Halaman III DIPA', 'Penyerapan Anggaran',
                    'Belanja Kontraktual', 'Penyelesaian Tagihan', 
                    'Pengelolaan UP dan TUP', 'Capaian Output'
                ]
                df_display = df[display_columns + ['Nilai Total',
                                                   'Dispensasi SPM (Pengurang)',
                                                   'Nilai Akhir (Nilai Total/Konversi Bobot)']].copy()
                for col in component_cols:
                    df_display[col] = df.get(col, 0)
                final_cols = display_columns + component_cols + ['Nilai Total',
                                                                 'Dispensasi SPM (Pengurang)',
                                                                 'Nilai Akhir (Nilai Total/Konversi Bobot)']
                df_display = df_display[final_cols]

            # Search widget & styling
            search_query = st.text_input("🔎 Cari (ketik untuk filter di semua kolom)", value="", help="Cari teks pada semua kolom (case-insensitive).", key='search_detail')
            if search_query:
                q = str(search_query).strip().lower()
                mask = df_display.apply(lambda row: row.astype(str).str.lower().str.contains(q, na=False).any(), axis=1)
                df_display_filtered = df_display[mask].copy()
            else:
                df_display_filtered = df_display.copy()

            def highlight_top(s):
                if s.name == 'Peringkat':
                    return ['background-color: gold' if (pd.to_numeric(v, errors='coerce') <= 3) else '' for v in s]
                return ['' for _ in s]

            st.dataframe(
                df_display_filtered.style.apply(highlight_top).format(precision=2),
                use_container_width=True,
                height=600
            )


# HALAMAN 2: DASHBOARD INTERNAL KPPN (Protected)
def page_trend():
    st.title("🏛️ Early Warning System Kinerja Keuangan Satker")

    # 🔒 Access restriction (same password as Admin page)
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False

    if not st.session_state.authenticated:
        st.warning("🔒 Halaman ini memerlukan autentikasi Admin untuk diakses.")
        password = st.text_input("Masukkan Password", type="password")
        if st.button("Login"):
            if password == "109KPPN":
                st.session_state.authenticated = True
                st.success("✅ Login berhasil! Silakan akses halaman ini.")
                st.rerun()
            else:
                st.error("❌ Password salah!")
        return
    
    if not st.session_state.data_storage:
        st.warning("⚠️ Belum ada data yang diunggah. Silakan unggah data melalui halaman Admin.")
        return
    
    # Gabungkan semua data
    all_data = []
    for period, df in st.session_state.data_storage.items():
        df_copy = df.copy()
        # ensure Period & Period_Sort exist
        df_copy['Period'] = f"{period[0]} {period[1]}"
        df_copy['Period_Sort'] = f"{period[1]}-{period[0]}"
        all_data.append(df_copy)
    
    if not all_data:
        st.warning("⚠️ Belum ada data historis yang tersedia.")
        return
    
    df_all = pd.concat(all_data, ignore_index=True)
      
    # Analisis tren dan Early Warning System
    # Gunakan data periode terkini
    latest_period = sorted(st.session_state.data_storage.keys(), key=lambda x: (int(x[1]), MONTH_ORDER.get(x[0].upper(), 0)), reverse=True)[0]
    df_latest = st.session_state.data_storage[latest_period]

    st.markdown("---")
    st.subheader("🚨 Satker yang Memerlukan Perhatian Khusus")

    # 🎚️ Pengaturan Sumbu Y
    st.markdown("###### Atur Skala Nilai (Sumbu Y)")
    col_min, col_max = st.columns(2)
    with col_min:
        y_min_int = st.slider(
            "Nilai Minimum (Y-Axis)",
            min_value=0,
            max_value=50,
            value=50,
            step=1,
            key="ymin_internal"
        )
    with col_max:
        y_max_int = st.slider(
            "Nilai Maksimum (Y-Axis)",
            min_value=51,
            max_value=110,
            value=110,
            step=1,
            key="ymax_internal"
        )

    # 📊 Highlights Kinerja Satker yang Perlu Perhatian Khusus
    col1, col2 = st.columns(2)

    with col1:
        fig_up = create_problem_chart(
            df_latest,
            'Pengelolaan UP dan TUP',
            100,
            "Pengelolaan UP dan TUP Belum Optimal (< 100)",
            'less',
            y_min=y_min_int,
            y_max=y_max_int,
            show_yaxis=True  # Left chart shows Y-axis
        )
        if fig_up:
            st.plotly_chart(fig_up, use_container_width=True)
        else:
            st.success("✅ Semua satker sudah optimal untuk Pengelolaan UP dan TUP")

    with col2:
        fig_output = create_problem_chart(
            df_latest,
            'Capaian Output',
            100,
            "Capaian Output Belum Optimal (< 100)",
            'less',
            y_min=y_min_int,
            y_max=y_max_int,
            show_yaxis=False  # Right chart hides Y-axis
        )
        if fig_output:
            st.plotly_chart(fig_output, use_container_width=True)
        else:
            st.success("✅ Semua satker sudah optimal untuk Capaian Output")
    
    warnings = []

    st.markdown("---")
# Analisis Tren
    st.subheader("📈 Analisis Tren")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        # 🔍 DETAILED ERROR CHECKING
        # Map month names to numbers
        df_all['Month_Num'] = df_all['Bulan'].str.strip().str.upper().map(MONTH_ORDER)
        
        # Check for unmapped months
        missing_months = df_all[df_all['Month_Num'].isna()]
        if len(missing_months) > 0:
            st.error("❌ **DITEMUKAN BULAN YANG TIDAK VALID:**")
            
            # Group by period to show which files have issues
            problem_periods = missing_months.groupby(['Bulan', 'Tahun']).size().reset_index(name='Count')
            
            for _, row in problem_periods.iterrows():
                st.warning(f"⚠️ Periode **{row['Bulan']} {row['Tahun']}** - Nama bulan '{row['Bulan']}' tidak dikenali (ditemukan di {row['Count']} baris)")
            
            st.info("""
            **Solusi:**
            1. Periksa file Excel untuk periode yang bermasalah
            2. Pastikan nama bulan sesuai format: JANUARI, FEBRUARI, MARET, dst (huruf besar)
            3. Upload ulang file yang bermasalah dari halaman Admin
            """)
            
            # Show expected month names
            with st.expander("📋 Lihat format bulan yang valid"):
                st.write("Format yang diterima:")
                st.code(", ".join(MONTH_ORDER.keys()))
            
            # Option to proceed with cleaned data
            if st.checkbox("⚠️ Abaikan data bermasalah dan lanjutkan"):
                df_all = df_all.dropna(subset=['Month_Num'])
                st.info(f"✅ Data dibersihkan. Sisa {len(df_all)} baris.")
            else:
                st.stop()
        
        # Check for invalid years
        invalid_years = df_all[df_all['Tahun'].isna()]
        if len(invalid_years) > 0:
            st.error("❌ **DITEMUKAN TAHUN YANG TIDAK VALID:**")
            
            problem_periods = invalid_years.groupby(['Bulan']).size().reset_index(name='Count')
            for _, row in problem_periods.iterrows():
                st.warning(f"⚠️ Bulan **{row['Bulan']}** - Tahun tidak valid (ditemukan di {row['Count']} baris)")
            
            st.stop()
        
        # Try to create Period_Sort with detailed error handling
        try:
            # Convert to int safely
            df_all['Tahun_Int'] = df_all['Tahun'].astype(int)
            df_all['Month_Num_Int'] = df_all['Month_Num'].astype(int)
            
            # Create Period_Sort
            df_all['Period_Sort'] = df_all.apply(
                lambda x: f"{x['Tahun_Int']:04d}-{x['Month_Num_Int']:02d}", 
                axis=1
            )
                        
        except Exception as e:
            st.error(f"❌ **ERROR saat membuat Period_Sort:** {str(e)}")
            
            # Show problematic rows
            st.write("**Baris yang bermasalah:**")
            problem_cols = ['Bulan', 'Tahun', 'Month_Num', 'Kode Satker', 'Uraian Satker']
            st.dataframe(df_all[problem_cols].head(20))
            
            st.stop()
        
        # Now create the selectbox
        available_periods = sorted(df_all['Period_Sort'].unique())
        start_period = st.selectbox(
            "Periode Awal",
            options=available_periods,
            index=0
        )
    
    with col2:
        end_period = st.selectbox(
            "Periode Akhir",
            options=available_periods,
            index=len(available_periods) - 1
        )
    
    # Filter berdasarkan periode
    df_filtered = df_all[
        (df_all['Period_Sort'] >= start_period) & 
        (df_all['Period_Sort'] <= end_period)
    ]
    
    with col3:
        # Pilihan metrik
        metric_options = {
            'Nilai Akhir (Nilai Total/Konversi Bobot)': 'Nilai Akhir (Nilai Total/Konversi Bobot)',
            'Kualitas Perencanaan Anggaran': 'Kualitas Perencanaan Anggaran',
            'Kualitas Pelaksanaan Anggaran': 'Kualitas Pelaksanaan Anggaran',
            'Kualitas Hasil Pelaksanaan Anggaran': 'Kualitas Hasil Pelaksanaan Anggaran',
            'Revisi DIPA': 'Revisi DIPA',
            'Deviasi Halaman III DIPA': 'Deviasi Halaman III DIPA',
            'Penyerapan Anggaran': 'Penyerapan Anggaran',
            'Belanja Kontraktual': 'Belanja Kontraktual',
            'Penyelesaian Tagihan': 'Penyelesaian Tagihan',
            'Pengelolaan UP dan TUP': 'Pengelolaan UP dan TUP',
            'Capaian Output': 'Capaian Output'
        }
        
        selected_metric = st.selectbox(
            "Metrik yang Ditampilkan",
            options=list(metric_options.keys()),
            index=0
        )
    
    # Pilih satker
    # All keys are (month_str, year_str). To sort by year then month, create sortable key:
    def period_sort_key(k):
        mon, yr = k
        # convert year to int if possible, month remain string but sorting will be stable for same year
        try:
            y = int(yr)
        except Exception as e:
            st.warning(f"⚠️ Tidak bisa convert tahun '{yr}' untuk periode {mon}: {e}")
            y = 0
        return (y, mon)

    try:
        latest_period = sorted(st.session_state.data_storage.keys(), key=period_sort_key, reverse=True)[0]
        latest_df = st.session_state.data_storage[latest_period].copy()
    except Exception as e:
        st.error(f"❌ Error mendapatkan periode terbaru: {e}")
        st.write("**Periode yang tersedia:**")
        st.write(list(st.session_state.data_storage.keys()))
        st.stop()
    
    # Make sure 'Kode Satker' exists and is a string
    if 'Kode Satker' in latest_df.columns:
        latest_df['Kode Satker'] = latest_df['Kode Satker'].astype(str)
    else:
        latest_df['Kode Satker'] = latest_df.index.astype(str)

    bottom_10_default = latest_df.nsmallest(10, 'Nilai Akhir (Nilai Total/Konversi Bobot)')['Kode Satker'].astype(str).tolist()
    
    # use the new 'Satker' column for selection (unique)
    all_satker = sorted(df_all['Satker'].unique())
    selected_satker = st.multiselect(
        "Pilih Satker",
        options=all_satker,
        default=[s for s in all_satker if any(str(code) in s for code in bottom_10_default)][:10]
    )
    
    if not selected_satker:
        st.warning("Silakan pilih minimal satu satker untuk melihat tren.")
        return
    
    # Filter berdasarkan satker (use 'Satker' to avoid duplicate names)
    df_plot = df_filtered[df_filtered['Satker'].isin(selected_satker)]
    
    # Buat line chart
    fig = go.Figure()
    
    try:
        for satker in selected_satker:
            df_satker = df_plot[df_plot['Satker'] == satker].sort_values('Period_Sort')

            # Ensure x-axis uses correct chronological month order
            categories = [f"{m} {y}" for y, m in sorted(
                {(int(x['Tahun']), x['Bulan'].upper()) for _, x in df_all.iterrows()},
                key=lambda t: (t[0], MONTH_ORDER.get(t[1], 0))
            )]
            
            fig.add_trace(go.Scatter(
                x=pd.Categorical(
                    df_satker['Period'],
                    categories=categories,
                    ordered=True
                ),
                y=df_satker[selected_metric],
                mode='lines+markers',
                name=satker,
                hovertemplate='<b>%{fullData.name}</b><br>Periode: %{x}<br>Nilai: %{y:.2f}<extra></extra>'
            ))
    except Exception as e:
        st.error(f"❌ Error membuat chart: {str(e)}")
        st.write("**Debug Info:**")
        st.write(f"Selected satker: {selected_satker}")
        st.write(f"df_plot shape: {df_plot.shape}")
        st.write(f"Unique periods in df_plot: {df_plot['Period'].unique()}")
        st.stop()
    
    fig.update_layout(
        title=f"Tren {selected_metric}",
        xaxis_title="Periode",
        yaxis_title="Nilai",
        height=600,
        hovermode='x unified',
        legend=dict(
            orientation="v",
            yanchor="top",
            y=1,
            xanchor="left",
            x=1.02
        )
    )
    
    st.plotly_chart(fig, use_container_width=True)

    # Early Warning Satker Tren Menurun
    warnings = []  # Initialize warnings list
    
    for satker in selected_satker:
        df_satker = df_plot[df_plot['Satker'] == satker].sort_values('Period_Sort')
        
        if len(df_satker) >= 2:
            values = df_satker[selected_metric].values
            
            # Cek tren menurun (2 periode terakhir)
            if len(values) >= 2:
                last_value = values[-1]
                prev_value = values[-2]
                
                if last_value < prev_value:
                    decrease = prev_value - last_value
                    warnings.append({
                        'Satker': satker,
                        'Metrik': selected_metric,
                        'Nilai Sebelumnya': prev_value,
                        'Nilai Terkini': last_value,
                        'Penurunan': decrease
                    })
    
    if warnings:
        st.warning(f"⚠️ Ditemukan {len(warnings)} satker dengan tren menurun!")
        
        for w in warnings:
            st.markdown(f"""
            **{w['Satker']}**  
            - Metrik: {w['Metrik']}
            - Nilai sebelumnya: {w['Nilai Sebelumnya']:.2f}
            - Nilai terkini: {w['Nilai Terkini']:.2f}
            - Penurunan: {w['Penurunan']:.2f} poin
            """)
            st.markdown("---")
    else:
        st.success("✅ Tidak ada satker dengan tren menurun pada periode yang dipilih!")
        
# ============================================================
# 🔐 HALAMAN 3: ADMIN 
# ============================================================
def page_admin():
    st.title("🔐 Halaman Administrasi")
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False

    if not st.session_state.authenticated:
        st.warning("🔒 Halaman ini memerlukan autentikasi")
        password = st.text_input("Masukkan Password", type="password")
        if st.button("Login"):
            if password == "109KPPN":
                st.session_state.authenticated = True
                st.success("✅ Login berhasil!")
                st.rerun()
            else:
                st.error("❌ Password salah!")
        return

    st.success("✅ Anda telah login sebagai Admin")

    # 🧩 Debug GitHub connection
    with st.expander("🧩 Debug GitHub Connection"):
        try:
            token = st.secrets["GITHUB_TOKEN"]
            repo_name = st.secrets["GITHUB_REPO"]
            g = Github(auth=Auth.Token(token))
            repo = g.get_repo(repo_name)
            st.success(f"Terhubung ke GitHub repo: {repo.full_name}")
        except Exception as e:
            st.error(f"❌ Gagal terhubung ke GitHub: {e}")

    if st.button("🚪 Logout"):
        st.session_state.authenticated = False
        st.rerun()

    st.markdown("---")
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📤 Upload Data",
        "🗑️ Hapus Data",
        "📥 Download Data",
        "📋 Download Template",
        "🕓 Riwayat Aktivitas"
    ])

    # ============================================================
    # TAB 1: UPLOAD DATA (IKPA, DIPA, Referensi)
    # ============================================================
    with tab1:

        st.subheader("📤 Upload Data Bulanan IKPA")

        upload_year = st.selectbox(
            "Pilih Tahun",
            list(range(2020, 2031)),
            index=list(range(2020, 2031)).index(datetime.now().year)
        )

        uploaded_files = st.file_uploader(
            "Pilih satu atau beberapa file Excel IKPA",
            type=['xlsx', 'xls'],
            accept_multiple_files=True
        )

        if uploaded_files:

            st.info("📄 File yang diupload:")
            for f in uploaded_files:
                st.write("•", f.name)

            if st.button("🔄 Proses Semua Data IKPA", type="primary"):

                with st.spinner("Memproses semua file..."):

                    # Mapping Bulan Lengkap
                    MONTH_FIX = {
                        "JAN": "JANUARI", "JANUARY": "JANUARI", "JANUARI": "JANUARI",
                        "FEB": "FEBRUARI", "FEBRUARY": "FEBRUARI", "FEBRUARI": "FEBRUARI",
                        "MAR": "MARET", "MRT": "MARET", "MARET": "MARET",
                        "APR": "APRIL", "APRIL": "APRIL",
                        "MEI": "MEI",
                        "JUN": "JUNI", "JUNE": "JUNI", "JUNI": "JUNI",
                        "JUL": "JULI", "JULY": "JULI", "JULI": "JULI",
                        "AGT": "AGUSTUS", "AGS": "AGUSTUS", "AUG": "AGUSTUS", "AGUSTUS": "AGUSTUS",
                        "SEP": "SEPTEMBER", "SEPT": "SEPTEMBER", "SEPTEMBER": "SEPTEMBER",
                        "OKT": "OKTOBER", "OCT": "OKTOBER", "OKTOBER": "OKTOBER",
                        "NOV": "NOVEMBER", "NOVEMBER": "NOVEMBER",
                        "DES": "DESEMBER", "DEC": "DESEMBER", "DESEMBER": "DESEMBER"
                    }

                    import re

                    # ==============================
                    # NORMALISASI KODE SATKER KUAT
                    # ==============================
                    def normalize_kode_satker(x):
                        if pd.isna(x): return ""
                        x = str(x)
                        x = re.sub(r"[^0-9]", "", x)  # keep numeric only
                        return x.zfill(6)

                    for uploaded_file in uploaded_files:
                        try:
                            # 1) Load file untuk deteksi bulan
                            df_temp = pd.read_excel(uploaded_file, header=None)
                            df_temp = df_temp.loc[:, ~df_temp.columns.astype(str).str.contains("Unnamed")]

                            raw_text = " ".join(df_temp.astype(str).fillna("").values.flatten()).upper()

                            # 2) Cari bulan otomatis dari seluruh isi file
                            final_month = "UNKNOWN"
                            for k, v in MONTH_FIX.items():
                                if k in raw_text:
                                    final_month = v
                                    break

                            # 3) Jika gagal, ambil dari nama file
                            if final_month == "UNKNOWN":
                                filename = uploaded_file.name.upper().replace(".XLSX", "").replace(".XLS", "")
                                parts = filename.split("_")
                                if len(parts) >= 2:
                                    clean_month = re.sub(r"[^A-Z]", "", parts[-2])
                                    final_month = MONTH_FIX.get(clean_month, clean_month)

                            final_month = final_month.upper()

                            # 4) Proses Excel
                            df_processed, _, _ = process_excel_file(uploaded_file, upload_year)
                            if df_processed is None:
                                st.warning(f"⚠️ Gagal memproses file: {uploaded_file.name}")
                                continue

                            df_processed["Bulan"] = final_month
                            df_processed["Tahun"] = int(upload_year)

                            # Hapus semua unnamed
                            df_processed = df_processed.loc[:, ~df_processed.columns.str.contains("Unnamed")]

                            # Normalisasi kode satker IKPA
                            if "Kode Satker" in df_processed.columns:
                                df_processed["Kode Satker"] = df_processed["Kode Satker"].apply(normalize_kode_satker)
                            else:
                                df_processed["Kode Satker"] = ""

                            # ============================================
                            # MERGE IKPA + DIPA (REVISI TERBARU)
                            # ============================================
                            df_final = df_processed.copy()

                            # Pastikan data DIPA ada
                            if "data_dipa_by_year" in st.session_state:
                                dipa_year = st.session_state.data_dipa_by_year.get(upload_year)

                                if dipa_year is not None and not dipa_year.empty:
                                    
                            # --- SIMPAN DATA DIPA KE SESSION STATE UNTUK DOWNLOAD ---
                                    if "dipa_storage" not in st.session_state:
                                        st.session_state.dipa_storage = {}

                                    # key download pakai tahun
                                    st.session_state.dipa_storage[str(upload_year)] = dipa_year.copy()

                                    # ============================================
                                    # 1. Hilangkan kolom Unnamed otomatis
                                    # ============================================
                                    dipa_year = dipa_year.loc[:, ~dipa_year.columns.str.contains("^Unnamed")]

                                    # ============================================
                                    # 2. Deteksi otomatis kolom Total Pagu
                                    # ============================================
                                    pagu_candidates = [
                                        "Pagu (Jumlah)", "Total Pagu", "Pagu",
                                        "Jumlah", "Total Anggaran"
                                    ]
                                    pagu_col = next((c for c in pagu_candidates if c in dipa_year.columns), None)

                                    if pagu_col:
                                        dipa_year = dipa_year.rename(columns={pagu_col: "Total Pagu"})
                                    else:
                                        dipa_year["Total Pagu"] = 0

                                    # ============================================
                                    # 3. Deteksi otomatis kolom Tanggal Posting Revisi
                                    # ============================================
                                    tgl_candidates = [
                                        "Tanggal Posting Revisi", "Tgl Posting Revisi",
                                        "Tanggal_Revisi", "Tanggal Posting"
                                    ]
                                    tgl_col = next((c for c in tgl_candidates if c in dipa_year.columns), None)

                                    if tgl_col:
                                        dipa_year = dipa_year.rename(columns={tgl_col: "Tanggal Posting Revisi"})
                                    else:
                                        dipa_year["Tanggal Posting Revisi"] = None

                                    # ============================================
                                    # 4. Normalisasi kode satker
                                    # ============================================
                                    if "Kode Satker" in dipa_year.columns:
                                        dipa_year["Kode Satker"] = dipa_year["Kode Satker"].apply(normalize_kode_satker)

                                    # ============================================
                                    # 5. MERGE ke IKPA
                                    # ============================================
                                    df_final = df_final.merge(
                                        dipa_year[["Kode Satker", "Total Pagu", "Tanggal Posting Revisi"]]
                                            .rename(columns={"Total Pagu": "Total Pagu DIPA"}),
                                        on="Kode Satker",
                                        how="left"
                                    )

                                    # ============================================
                                    # 6. Kategori Satker
                                    # ============================================
                                    if "Total Pagu DIPA" in df_final.columns:
                                        p70 = df_final["Total Pagu DIPA"].astype(float).quantile(0.70)
                                        p40 = df_final["Total Pagu DIPA"].astype(float).quantile(0.40)

                                        def kategori(pagu):
                                            pagu = float(pagu) if pagu not in [None, ""] else 0
                                            if pagu >= p70:
                                                return "Satker Besar"
                                            elif pagu >= p40:
                                                return "Satker Sedang"
                                            return "Satker Kecil"

                                        df_final["Jenis Satker"] = df_final["Total Pagu DIPA"].apply(kategori)

                            # ========================
                            # SIMPAN SESSION STATE
                            # ========================
                            key = (final_month, str(upload_year))
                            st.session_state.data_storage[key] = df_final.copy()

                            # ============================================
                            # SIMPAN KE GITHUB + DOWNLOAD (HASIL BERSIH)
                            # ============================================

                            KEEP_COLUMNS = [
                                "Kode KPPN", "Kode BA", "Kode Satker",
                                "Uraian Satker-RINGKAS",
                                "Kualitas Perencanaan Anggaran",
                                "Kualitas Pelaksanaan Anggaran",
                                "Kualitas Hasil Pelaksanaan Anggaran",
                                "Revisi DIPA",
                                "Deviasi Halaman III DIPA",
                                "Penyerapan Anggaran",
                                "Belanja Kontraktual",
                                "Penyelesaian Tagihan",
                                "Pengelolaan UP dan TUP",
                                "Capaian Output",
                                "Nilai Total",
                                "Nilai Akhir (Nilai Total/Konversi Bobot)",

                                # kolom hasil merge
                                "Total Pagu DIPA",
                                "Tanggal Posting Revisi",
                                "Jenis Satker",

                                "Bulan",
                                "Tahun"
                            ]

                            df_excel = df_final[[c for c in KEEP_COLUMNS if c in df_final.columns]]

                            excel_bytes = io.BytesIO()
                            with pd.ExcelWriter(excel_bytes, engine="openpyxl") as writer:
                                df_excel.to_excel(writer, index=False, sheet_name="Data IKPA")

                            excel_bytes.seek(0)
                            save_file_to_github(
                                excel_bytes.getvalue(),
                                f"IKPA_{final_month}_{upload_year}.xlsx",
                                folder="data"
                            )
                            
                            # Log
                            st.session_state.activity_log.append({
                                "Waktu": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                "Aksi": "Upload",
                                "Periode": f"{final_month} {upload_year}",
                                "Status": "Sukses"
                            })

                            st.success(f"✅ {uploaded_file.name} → {final_month} {upload_year} berhasil disimpan.")

                        # ====================================
                        # PERBAIKAN PENTING → Tutup try agar tidak error
                        # ====================================
                        except Exception as e:
                            st.error(f"❌ Error saat memproses data: {e}")


        # ============================================================
        # SUBMENU: UPLOAD DATA DIPA
        # ============================================================
        st.markdown("---")
        st.subheader("📤 Upload Data DIPA")

        upload_year_dipa = st.selectbox(
            "Pilih Tahun DIPA",
            list(range(2020, 2031)),
            index=list(range(2020, 2031)).index(datetime.now().year),
            key="year_dipa"
        )

        uploaded_dipa_file = st.file_uploader("Pilih file Excel DIPA", type=['xlsx', 'xls', 'csv'], key="dipa_upload")

        # ============================
        # Fungsi untuk membersihkan file DIPA
        # ============================
        def clean_dipa(df):

            # defensive copy
            df = df.copy()

            # 0. If dataframe has no columns (empty) -> return as is
            if df is None or df.shape[1] == 0:
                return df

            # 1. Remove Unnamed columns (case-insensitive)
            df = df.loc[:, ~df.columns.astype(str).str.lower().str.contains("unnamed")]

            # 2. Try to detect header row by scanning first few rows for keywords
            header_row = None
            max_scan = min(10, len(df))
            for i in range(max_scan):
                # join sample of row text
                try:
                    row_text = " ".join([str(x) for x in df.iloc[i].tolist()]).lower()
                except Exception:
                    row_text = ""
                if any(k in row_text for k in ["kode satker", "kode", "satker", "pagu", "no", "tanggal"]):
                    header_row = i
                    break

            # 3. If header row found, set columns; else try first row as header if it looks like strings
            if header_row is not None:
                df.columns = df.iloc[header_row].fillna('').astype(str)
                df = df[(header_row + 1):]
            else:
                # fallback: ensure columns are strings
                df.columns = df.columns.astype(str)

            # 4. Drop rows that are completely empty
            df = df.dropna(how="all")

            # 5. Normalize column names: strip, remove newlines, collapse multiple spaces, uppercase for matching
            clean_cols = (
                df.columns.astype(str)
                .str.strip()
                .str.replace('\n', ' ', regex=False)
                .str.replace('\r', ' ', regex=False)
                .str.replace('\s+', ' ', regex=True)
            )
            df.columns = clean_cols

            # 6. Secondary removal of 'Unnamed' columns if any remained
            df = df.loc[:, ~df.columns.astype(str).str.lower().str.contains("unnamed")]

            # 7. Standardize important column names (handle variants)
            rename_map = {}
            for c in df.columns:
                cu = c.strip().lower()
                if "kode" in cu and "satker" in cu:
                    rename_map[c] = "Kode Satker"
                elif "pagu" in cu or "total pagu" in cu or "jumlah pagu" in cu:
                    rename_map[c] = "Total Pagu"
                elif "tanggal" in cu and "revisi" in cu:
                    rename_map[c] = "Tanggal Posting Revisi"
                elif cu in ["tgl revisi", "tgl posting revisi"]:
                    rename_map[c] = "Tanggal Posting Revisi"
                # keep other columns as-is

            if rename_map:
                df = df.rename(columns=rename_map)

            # 8. Reset index
            df = df.reset_index(drop=True)

            return df

        if uploaded_dipa_file is not None:
            try:
                # Baca file untuk preview (pastikan seek ke awal)
                # Re-read file and clean
                uploaded_dipa_file.seek(0)
                filename_preview = getattr(uploaded_dipa_file, "name", "uploaded_dipa")

                if filename_preview.lower().endswith('.csv'):
                    df_read = pd.read_csv(uploaded_dipa_file, dtype=str, encoding='utf-8', engine='python')
                else:
                    # HEADER ASLI ADA DI BARIS KE-3 → row index 2
                    df_read = pd.read_excel(uploaded_dipa_file, dtype=str, header=2)

                # Bersihkan tabel sebelum diproses
                df_read = clean_dipa(df_read)

                # Preview tahun yang terdeteksi dari data (prefer Tanggal Posting Revisi)
                if 'Tanggal Posting Revisi' in df_temp_dipa.columns and not df_temp_dipa['Tanggal Posting Revisi'].dropna().empty:
                    try:
                        sample_date = pd.to_datetime(df_temp_dipa['Tanggal Posting Revisi'].dropna().iloc[0], errors='coerce')
                        if pd.isna(sample_date):
                            year_preview = upload_year_dipa
                        else:
                            year_preview = sample_date.year
                    except Exception:
                        year_preview = upload_year_dipa
                else:
                    # fallback: try to detect year from any column that contains 4-digit year
                    year_preview = upload_year_dipa
                    for col in df_temp_dipa.columns:
                        sample_vals = df_temp_dipa[col].dropna().astype(str).head(10).tolist()
                        for v in sample_vals:
                            m = re.search(r'(\b20\d{2}\b)', v)
                            if m:
                                year_preview = int(m.group(1))
                                break
                        if year_preview != upload_year_dipa:
                            break

                period_key_preview = str(year_preview)
                uploaded_dipa_file.seek(0)

                # Cek apakah data tahun ini sudah ada
                if "data_dipa_by_year" not in st.session_state:
                    st.session_state.data_dipa_by_year = {}

                if int(period_key_preview) in st.session_state.data_dipa_by_year:
                    st.warning(f"⚠️ Data DIPA untuk tahun **{year_preview}** sudah ada.")
                    confirm_replace_dipa = st.checkbox(
                        "✅ Ganti data yang sudah ada.",
                        key=f"confirm_replace_dipa_{year_preview}"
                    )
                else:
                    confirm_replace_dipa = True
                    st.info(f"📝 Akan mengunggah data baru untuk tahun: **{year_preview}**")

            except Exception as e:
                st.error(f"❌ Gagal membaca preview file: {e}")
                confirm_replace_dipa = False

            if st.button("🔄 Proses Data DIPA", type="primary", disabled=not confirm_replace_dipa):
                with st.spinner("Memproses data DIPA..."):
                    try:
                        # Re-read file and clean
                        uploaded_dipa_file.seek(0)
                        filename_preview = getattr(uploaded_dipa_file, "name", "uploaded_dipa")
                        if filename_preview.lower().endswith('.csv'):
                            df_read = pd.read_csv(uploaded_dipa_file, dtype=str, encoding='utf-8', engine='python')
                        else:
                            df_read = pd.read_excel(uploaded_dipa_file, dtype=str)

                        # Clean the DIPA table BEFORE processing
                        df_read = clean_dipa(df_read)

                        # Process DIPA (your existing robust processor)
                        dfp = process_dipa_dataframe(df_read, source_name=filename_preview)

                        if dfp is None or dfp.empty:
                            st.error("❌ Gagal memproses file DIPA.")
                            st.stop()

                        # Normalize Kode Satker in processed DIPA
                        if 'Kode Satker' in dfp.columns:
                            dfp['Kode Satker'] = dfp['Kode Satker'].apply(normalize_kode_satker)
                        else:
                            dfp['Kode Satker'] = ''

                        # Ensure 'Tanggal Posting Revisi' datetime
                        if 'Tanggal Posting Revisi' in dfp.columns:
                            dfp['Tanggal Posting Revisi'] = pd.to_datetime(dfp['Tanggal Posting Revisi'], errors='coerce')

                        # Normalize/rename Total Pagu variations in dfp (if present)
                        for c in list(dfp.columns):
                            if 'pagu' in c.lower() and 'total' not in c.lower():
                                # rename loosely to 'Total Pagu' to be safe
                                dfp = dfp.rename(columns={c: 'Total Pagu'})

                        # Group by year and save
                        years = sorted(dfp['Tahun'].dropna().unique().astype(int).tolist())

                        for yr in years:
                            df_year = dfp[dfp['Tahun'] == yr].copy().reset_index(drop=True)

                            # Ensure Kode Satker normalized in existing dataset too
                            existing = st.session_state.data_dipa_by_year.get(int(yr))
                            if existing is not None and not existing.empty:
                                existing = existing.copy()
                                if 'Kode Satker' in existing.columns:
                                    existing['Kode Satker'] = existing['Kode Satker'].apply(normalize_kode_satker)
                                else:
                                    existing['Kode Satker'] = ''

                                # Normalize tanggal in existing as well
                                if 'Tanggal Posting Revisi' in existing.columns:
                                    existing['Tanggal Posting Revisi'] = pd.to_datetime(existing['Tanggal Posting Revisi'], errors='coerce')

                                combined = pd.concat([existing, df_year], ignore_index=True, sort=False)

                                # If Tanggal Posting Revisi present, sort by (Kode Satker, Tanggal Posting Revisi) then take last per Kode Satker
                                if 'Tanggal Posting Revisi' in combined.columns:
                                    combined = combined.sort_values(by=['Kode Satker', 'Tanggal Posting Revisi'])
                                    # take last row per Kode Satker
                                    combined_latest = combined.groupby('Kode Satker', as_index=False).tail(1).reset_index(drop=True)
                                else:
                                    # fallback: last occurrence per Kode Satker
                                    combined_latest = combined.groupby('Kode Satker', as_index=False).last().reset_index(drop=True)

                                st.session_state.data_dipa_by_year[int(yr)] = combined_latest
                            else:
                                # Ensure df_year has normalized kode and date types
                                if 'Kode Satker' in df_year.columns:
                                    df_year['Kode Satker'] = df_year['Kode Satker'].apply(normalize_kode_satker)
                                else:
                                    df_year['Kode Satker'] = ''
                                if 'Tanggal Posting Revisi' in df_year.columns:
                                    df_year['Tanggal Posting Revisi'] = pd.to_datetime(df_year['Tanggal Posting Revisi'], errors='coerce')

                                st.session_state.data_dipa_by_year[int(yr)] = df_year.reset_index(drop=True)

                            # ✅ Save to GitHub with standard names
                            filename_dipa = f"DIPA_{yr}.xlsx"
                            excel_bytes_dipa = io.BytesIO()
                            with pd.ExcelWriter(excel_bytes_dipa, engine='openpyxl') as writer:
                                st.session_state.data_dipa_by_year[int(yr)].to_excel(
                                    writer, index=False, sheet_name=f'DIPA_{yr}',
                                    startrow=0, startcol=0
                                )

                                # Format header (if worksheet exists)
                                try:
                                    workbook = writer.book
                                    worksheet = writer.sheets[f'DIPA_{yr}']
                                    # apply header formatting to first row (row 1 is header)
                                    for cell in worksheet[1]:
                                        try:
                                            cell.font = Font(bold=True, color="FFFFFF")
                                            cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
                                            cell.alignment = Alignment(horizontal="center", vertical="center")
                                        except Exception:
                                            pass
                                except Exception:
                                    pass

                            excel_bytes_dipa.seek(0)
                            save_file_to_github(excel_bytes_dipa.getvalue(), filename_dipa, folder="data_dipa")

                        st.success(f"✅ Data DIPA tahun {', '.join(map(str, years))} berhasil disimpan.")
                        st.snow()

                        st.session_state.activity_log.append({
                            "Waktu": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "Aksi": "Upload DIPA",
                            "Periode": ", ".join([str(y) for y in years]),
                            "Status": "✅ Sukses"
                        })

                    except Exception as e:
                        st.error(f"❌ Gagal menyimpan ke GitHub: {e}")

        # ============================================================
        # SUBMENU: Upload Data Referensi
        # ============================================================
        st.markdown("---")
        st.subheader("📚 Upload / Perbarui Data Referensi Satker & K/L")
        st.info("""
        - File referensi ini berisi kolom: **Kode BA, K/L, Kode Satker, Uraian Satker-SINGKAT, Uraian Satker-LENGKAP**  
        - Saat diupload, sistem akan **menggabungkan** dengan data lama:  
        🔹 Jika `Kode Satker` sudah ada → baris lama akan **diganti**  
        🔹 Jika `Kode Satker` belum ada → akan **ditambahkan baru**
        """)

        uploaded_ref = st.file_uploader(
            "📤 Pilih File Data Referensi Satker & K/L",
            type=['xlsx', 'xls'],
            key="ref_upload"
        )

        if uploaded_ref is not None:
            try:
                new_ref = pd.read_excel(uploaded_ref)
                new_ref.columns = [c.strip() for c in new_ref.columns]

                required = ['Kode BA', 'K/L', 'Kode Satker', 'Uraian Satker-SINGKAT', 'Uraian Satker-LENGKAP']
                if not all(col in new_ref.columns for col in required):
                    st.error("❌ Kolom wajib tidak lengkap dalam file referensi.")
                    st.stop()

                new_ref['Kode Satker'] = new_ref['Kode Satker'].apply(normalize_kode_satker)

                # Gabungkan atau buat baru
                if 'reference_df' in st.session_state:
                    old_ref = st.session_state.reference_df.copy()

                    # 🔹 Normalize old reference too (critical!)
                    if 'Kode Satker' in old_ref.columns:
                        old_ref['Kode Satker'] = old_ref['Kode Satker'].apply(normalize_kode_satker)

                    # 🔹 Combine and deduplicate
                    merged = pd.concat([old_ref, new_ref], ignore_index=True)
                    merged = merged.drop_duplicates(subset=['Kode Satker'], keep='last')

                    # 🔹 Optional: enforce consistent string stripping
                    merged['Kode Satker'] = merged['Kode Satker'].astype(str).str.strip()

                    st.session_state.reference_df = merged
                    st.success(f"✅ Data Referensi diperbarui ({len(merged)} total baris).")
                else:
                    st.session_state.reference_df = new_ref
                    st.success(f"✅ Data Referensi baru dimuat ({len(new_ref)} baris).")

                st.dataframe(st.session_state.reference_df.tail(10), use_container_width=True)

                # ✅ Save merged reference data permanently to GitHub
                try:
                    excel_bytes_ref = io.BytesIO()
                    with pd.ExcelWriter(excel_bytes_ref, engine='openpyxl') as writer:
                        st.session_state.reference_df.to_excel(
                            writer, index=False, sheet_name='Data Referensi',
                            startrow=0, startcol=0  # ✅ PERBAIKAN
                        )
                        
                        # Format header
                        workbook = writer.book
                        worksheet = writer.sheets['Data Referensi']
                        
                        for cell in worksheet[1]:
                            cell.font = Font(bold=True, color="FFFFFF")
                            cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
                            cell.alignment = Alignment(horizontal="center", vertical="center")
                    
                    excel_bytes_ref.seek(0)

                    save_file_to_github(
                        excel_bytes_ref.getvalue(),
                        "Template_Data_Referensi.xlsx",
                        folder="templates"
                    )
                    st.success("💾 Data Referensi berhasil disimpan ke GitHub (templates/Template_Data_Referensi.xlsx).")
                except Exception as e:
                    st.error(f"❌ Gagal menyimpan Data Referensi ke GitHub: {e}")

            except Exception as e:
                st.error(f"❌ Gagal memproses Data Referensi: {e}")

    # ============================================================
    # TAB 2: HAPUS DATA
    # ============================================================
    with tab2:
        # Submenu Hapus Data IKPA
        st.subheader("🗑️ Hapus Data Bulanan IKPA")
        if not st.session_state.data_storage:
            st.info("ℹ️ Belum ada data IKPA tersimpan.")
        else:
            available_periods = sorted(st.session_state.data_storage.keys(), reverse=True)
            period_to_delete = st.selectbox(
                "Pilih periode yang akan dihapus",
                options=available_periods,
                format_func=lambda x: f"{x[0].capitalize()} {x[1]}"
            )
            month, year = period_to_delete
            filename = f"data/IKPA_{month}_{year}.xlsx"

            confirm_delete = st.checkbox(
                f"⚠️ Hapus data {month} {year} dari sistem dan GitHub.",
                key=f"confirm_delete_{month}_{year}"
            )

            if st.button("🗑️ Hapus Data IKPA Ini", type="primary") and confirm_delete:
                try:
                    del st.session_state.data_storage[period_to_delete]
                    token = st.secrets.get("GITHUB_TOKEN")
                    repo_name = st.secrets.get("GITHUB_REPO")
                    g = Github(auth=Auth.Token(token))
                    repo = g.get_repo(repo_name)
                    contents = repo.get_contents(f"data/IKPA_{month}_{year}.xlsx")
                    repo.delete_file(contents.path, f"Delete {filename}", contents.sha)
                    st.success(f"✅ Data {month} {year} dihapus dari sistem & GitHub.")
                    st.snow()
                    st.session_state.activity_log.append({
                        "Waktu": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "Aksi": "Hapus IKPA",
                        "Periode": f"{month} {year}",
                        "Status": "✅ Sukses"
                    })
                except Exception as e:
                    st.error(f"❌ Gagal menghapus data: {e}")

        # Submenu Hapus Data DIPA
        st.markdown("---")
        st.subheader("🗑️ Hapus Data DIPA")
        if not st.session_state.get("data_dipa_by_year"):
            st.info("ℹ️ Belum ada data DIPA tersimpan.")
        else:
            available_years = sorted(st.session_state.data_dipa_by_year.keys(), reverse=True)
            year_to_delete = st.selectbox(
                "Pilih tahun DIPA yang akan dihapus",
                options=available_years,
                format_func=lambda x: f"Tahun {x}",
                key="delete_dipa_year"
            )
            filename_dipa = f"data_dipa/DIPA_{year_to_delete}.xlsx"

            confirm_delete_dipa = st.checkbox(
                f"⚠️ Hapus data DIPA tahun {year_to_delete} dari sistem dan GitHub.",
                key=f"confirm_delete_dipa_{year_to_delete}"
            )

            if st.button("🗑️ Hapus Data DIPA Ini", type="primary", key="btn_delete_dipa") and confirm_delete_dipa:
                try:
                    del st.session_state.data_dipa_by_year[year_to_delete]
                    token = st.secrets.get("GITHUB_TOKEN")
                    repo_name = st.secrets.get("GITHUB_REPO")
                    g = Github(auth=Auth.Token(token))
                    repo = g.get_repo(repo_name)
                    contents = repo.get_contents(filename_dipa)
                    repo.delete_file(contents.path, f"Delete {filename_dipa}", contents.sha)
                    st.success(f"✅ Data DIPA tahun {year_to_delete} dihapus dari sistem & GitHub.")
                    st.snow()
                    st.session_state.activity_log.append({
                        "Waktu": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "Aksi": "Hapus DIPA",
                        "Periode": f"Tahun {year_to_delete}",
                        "Status": "✅ Sukses"
                    })
                except Exception as e:
                    st.error(f"❌ Gagal menghapus data DIPA: {e}")

    # ============================================================
    # TAB 3: DOWNLOAD DATA (DIPERBAIKI LENGKAP)
    # ============================================================
    with tab3:
        # =========================================================
        # Submenu Download Data IKPA (DIPERBAIKI)
        # =========================================================
        st.subheader("📥 Download Data IKPA")
        
        if not st.session_state.data_storage:
            st.info("ℹ️ Belum ada data IKPA.")
        else:
            available_periods = sorted(
                st.session_state.data_storage.keys(), 
                reverse=True
            )
            
            period_to_download = st.selectbox(
                "Pilih periode untuk download",
                options=available_periods,
                format_func=lambda x: f"{x[0].capitalize()} {x[1]}"
            )
            
            # Ambil data IKPA
            df_download = st.session_state.data_storage[period_to_download].copy()
            month, year = period_to_download
            
            # ✅ PERBAIKAN: Merge dengan DIPA yang sudah unik
            if 'data_dipa_by_year' in st.session_state:
                dipa_year = st.session_state.data_dipa_by_year.get(int(year))
                
                if dipa_year is not None and not dipa_year.empty:
                    # 🔹 Pastikan DIPA unik per Kode Satker
                    dipa_unique = ensure_unique_dipa_per_satker(dipa_year)
                    
                    # 🔹 Normalize Kode Satker di kedua dataframe
                    if 'Kode Satker' in df_download.columns:
                        df_download['Kode Satker'] = df_download['Kode Satker'].apply(
                            normalize_kode_satker
                        )
                    
                    if 'Kode Satker' in dipa_unique.columns:
                        dipa_unique['Kode Satker'] = dipa_unique['Kode Satker'].apply(
                            normalize_kode_satker
                        )
                    
                    # 🔹 Deteksi kolom Total Pagu di DIPA
                    pagu_col = None
                    pagu_candidates = [
                        "Pagu (Jumlah)", "Total Pagu", "Pagu",
                        "Jumlah", "Total Anggaran"
                    ]
                    for col in pagu_candidates:
                        if col in dipa_unique.columns:
                            pagu_col = col
                            break
                    
                    if pagu_col:
                        dipa_unique = dipa_unique.rename(columns={pagu_col: "Total Pagu"})
                    else:
                        dipa_unique["Total Pagu"] = 0
                    
                    # 🔹 Merge IKPA dengan DIPA
                    cols_to_merge = ['Kode Satker', 'Total Pagu']
                    
                    if 'Tanggal Posting Revisi' in dipa_unique.columns:
                        cols_to_merge.append('Tanggal Posting Revisi')
                    
                    df_download = df_download.merge(
                        dipa_unique[cols_to_merge].rename(
                            columns={'Total Pagu': 'Total Pagu DIPA'}
                        ),
                        on='Kode Satker',
                        how='left'
                    )
                    
                    # 🔹 Tambahkan Kolom Jenis Satker
                    df_download = add_jenis_satker_column(df_download)
                    
                    st.success(
                        f"✅ Data DIPA tahun {year} berhasil digabung "
                        f"({len(dipa_unique)} satker unik)"
                    )
                else:
                    st.info(f"ℹ️ Data DIPA untuk tahun {year} tidak tersedia.")
                    df_download['Total Pagu DIPA'] = None
                    df_download['Tanggal Posting Revisi'] = None
                    df_download['Jenis Satker'] = 'Tidak Ada Data Pagu'
            else:
                st.info("ℹ️ Data DIPA belum dimuat ke sistem.")
                df_download['Total Pagu DIPA'] = None
                df_download['Tanggal Posting Revisi'] = None
                df_download['Jenis Satker'] = 'Tidak Ada Data Pagu'
            
            # =========================================================
            # Generate Excel dengan formatting
            # =========================================================
            output = io.BytesIO()
            
            # Kolom yang akan di-export (sesuaikan urutan)
            export_columns = [
                'Peringkat', 'Kode KPPN', 'Kode BA', 'Kode Satker',
                'Uraian Satker-RINGKAS',
                'Total Pagu DIPA',
                'Jenis Satker',
                'Tanggal Posting Revisi',
                'Kualitas Perencanaan Anggaran',
                'Kualitas Pelaksanaan Anggaran',
                'Kualitas Hasil Pelaksanaan Anggaran',
                'Revisi DIPA', 'Deviasi Halaman III DIPA',
                'Penyerapan Anggaran', 'Belanja Kontraktual',
                'Penyelesaian Tagihan', 'Pengelolaan UP dan TUP',
                'Capaian Output',
                'Nilai Total', 'Konversi Bobot',
                'Dispensasi SPM (Pengurang)',
                'Nilai Akhir (Nilai Total/Konversi Bobot)',
                'Bulan', 'Tahun'
            ]
            
            # Filter kolom yang ada
            df_excel = df_download[[col for col in export_columns if col in df_download.columns]]
            
            # Drop kolom internal jika ada
            df_excel = df_excel.drop(
                ['Bobot', 'Nilai Terbobot', 'Source', 'Period', 'Period_Sort'], 
                axis=1, 
                errors='ignore'
            )
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_excel.to_excel(
                    writer, 
                    index=False, 
                    sheet_name='Data IKPA', 
                    startrow=0, 
                    startcol=0
                )
                
                # Format header
                workbook = writer.book
                worksheet = writer.sheets['Data IKPA']
                
                # Style header row
                header_font = Font(bold=True, color="FFFFFF", size=11)
                header_fill = PatternFill(
                    start_color="366092", 
                    end_color="366092", 
                    fill_type="solid"
                )
                header_alignment = Alignment(
                    horizontal="center", 
                    vertical="center",
                    wrap_text=True
                )
                
                for cell in worksheet[1]:
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = header_alignment
                
                # Auto-adjust column width
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width
            
            output.seek(0)
            
            # Download button
            st.download_button(
                label="📥 Download Excel IKPA (dengan Data DIPA & Jenis Satker)",
                data=output,
                file_name=f"IKPA_{period_to_download[0]}_{period_to_download[1]}_lengkap.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            # Preview data
            with st.expander("👁️ Preview Data yang Akan Di-Download"):
                st.dataframe(
                    df_excel.head(20),
                    use_container_width=True
                )
                
                # Statistik Jenis Satker
                if 'Jenis Satker' in df_excel.columns:
                    st.markdown("##### 📊 Distribusi Jenis Satker")
                    jenis_counts = df_excel['Jenis Satker'].value_counts()
                    
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric(
                            "Satker Besar", 
                            jenis_counts.get('Satker Besar', 0)
                        )
                    with col2:
                        st.metric(
                            "Satker Sedang", 
                            jenis_counts.get('Satker Sedang', 0)
                        )
                    with col3:
                        st.metric(
                            "Satker Kecil", 
                            jenis_counts.get('Satker Kecil', 0)
                        )
                    with col4:
                        st.metric(
                            "Tanpa Data Pagu", 
                            jenis_counts.get('Tidak Ada Data Pagu', 0)
                        )
        
        # =========================================================
        # Submenu Download Data DIPA
        # =========================================================
        st.markdown("---")
        st.subheader("📥 Download Data DIPA")
        
        if not st.session_state.get("data_dipa_by_year"):
            st.info("ℹ️ Belum ada data DIPA.")
        else:
            available_years_download = sorted(
                st.session_state.data_dipa_by_year.keys(), 
                reverse=True
            )
            
            year_to_download = st.selectbox(
                "Pilih tahun DIPA untuk download",
                options=available_years_download,
                format_func=lambda x: f"Tahun {x}",
                key="download_dipa_year"
            )
            
            # ✅ Pastikan DIPA unik sebelum download
            df_download_dipa = st.session_state.data_dipa_by_year[year_to_download].copy()
            df_download_dipa = ensure_unique_dipa_per_satker(df_download_dipa)
            
            output_dipa = io.BytesIO()
            
            with pd.ExcelWriter(output_dipa, engine='openpyxl') as writer:
                df_download_dipa.to_excel(
                    writer, 
                    index=False, 
                    sheet_name=f'DIPA_{year_to_download}',
                    startrow=0, 
                    startcol=0
                )
                
                # Format header
                workbook = writer.book
                worksheet = writer.sheets[f'DIPA_{year_to_download}']
                
                for cell in worksheet[1]:
                    cell.font = Font(bold=True, color="FFFFFF")
                    cell.fill = PatternFill(
                        start_color="366092", 
                        end_color="366092", 
                        fill_type="solid"
                    )
                    cell.alignment = Alignment(
                        horizontal="center", 
                        vertical="center"
                    )
            
            output_dipa.seek(0)
            
            st.download_button(
                label=f"📥 Download Excel DIPA {year_to_download} (Unik per Satker)",
                data=output_dipa,
                file_name=f"DIPA_{year_to_download}_unik.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_download_dipa"
            )
            
            st.info(f"ℹ️ Total satker unik: {len(df_download_dipa)}")

        # Download Data Satker Tidak Terdaftar
        st.markdown("---")
        st.subheader("📥 Download Data Satker yang Belum Terdaftar di Tabel Referensi")
        
        if st.button("📥 Generate & Download Laporan"):
            st.info("ℹ️ Fitur ini menggunakan data dari session state untuk performa optimal.")
            
    # ============================================================
    # TAB 4: DOWNLOAD TEMPLATE
    # ============================================================
    with tab4:
        st.subheader("📋 Download Template")
        st.markdown("### 📘 Template IKPA")
        try:
            token = st.secrets["GITHUB_TOKEN"]
            repo_name = st.secrets["GITHUB_REPO"]
            g = Github(auth=Auth.Token(token))
            repo = g.get_repo(repo_name)
            file_content = repo.get_contents("templates/Template_IKPA.xlsx")
            template_data = base64.b64decode(file_content.content)
        except Exception:
            template_data = get_template_file()

        if template_data:
            st.download_button(
                label="📥 Download Template IKPA",
                data=template_data,
                file_name="Template_IKPA.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        st.markdown("---")
        st.markdown("### 📗 Template Data Referensi Satker & K/L")

        # 🧩 Use latest reference data for template content
        if 'reference_df' in st.session_state and not st.session_state.reference_df.empty:
            template_ref = st.session_state.reference_df.copy()
        else:
            # fallback: try load from GitHub
            try:
                token = st.secrets["GITHUB_TOKEN"]
                repo_name = st.secrets["GITHUB_REPO"]
                g = Github(auth=Auth.Token(token))
                repo = g.get_repo(repo_name)
                ref_content = repo.get_contents("templates/Template_Data_Referensi.xlsx")
                ref_data = base64.b64decode(ref_content.content)
                template_ref = pd.read_excel(io.BytesIO(ref_data))
            except Exception:
                template_ref = pd.DataFrame({
                    'No': [],
                    'Kode BA': [],
                    'K/L': [],
                    'Kode Satker': [],
                    'Uraian Satker-SINGKAT': [],
                    'Uraian Satker-LENGKAP': []
                })

        output_ref = io.BytesIO()
        with pd.ExcelWriter(output_ref, engine='openpyxl') as writer:
            # ✅ PERBAIKAN: Mulai dari A1
            template_ref.to_excel(writer, index=False, sheet_name='Data Referensi',
                                  startrow=0, startcol=0)
            
            # Format header
            workbook = writer.book
            worksheet = writer.sheets['Data Referensi']
            
            for cell in worksheet[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
                cell.alignment = Alignment(horizontal="center", vertical="center")
        
        output_ref.seek(0)

        st.download_button(
            label="📥 Download Template Data Referensi",
            data=output_ref,
            file_name="Template_Data_Referensi.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    # ============================================================
    # TAB 5: LOG AKTIVITAS
    # ============================================================
    with tab5:
        st.subheader("📖 Log Aktivitas GitHub")
        if not st.session_state.activity_log:
            st.info("Belum ada aktivitas.")
        else:
            df_log = pd.DataFrame(st.session_state.activity_log)
            st.dataframe(df_log[::-1].reset_index(drop=True), use_container_width=True)
            if st.button("🧹 Bersihkan Log"):
                st.session_state.activity_log = []
                st.success("🧹 Log dibersihkan.")

# ===============================
# 🔹 MAIN APP
# ===============================
def main():
    # ============================================================
    # 🧩 Auto-load Reference Data from GitHub FIRST
    # ============================================================
    if 'reference_df' not in st.session_state:
        try:
            token = st.secrets["GITHUB_TOKEN"]
            repo_name = st.secrets["GITHUB_REPO"]
            g = Github(auth=Auth.Token(token))
            repo = g.get_repo(repo_name)
            ref_path = "templates/Template_Data_Referensi.xlsx"
            ref_file = repo.get_contents(ref_path)
            ref_data = base64.b64decode(ref_file.content)

            ref_df = pd.read_excel(io.BytesIO(ref_data))
            short_col = 'Uraian Satker-SINGKAT'
            ref_df.columns = [c.strip() for c in ref_df.columns]  # normalize header whitespace
            st.session_state.reference_df = ref_df

            if short_col not in ref_df.columns:
                # Build simple diagnostic workbook with reference columns + example head
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    pd.DataFrame({"Reference Columns": list(ref_df.columns)}).to_excel(writer, sheet_name='Reference_Columns', index=False)
                    ref_df.head(200).to_excel(writer, sheet_name='Reference_Sample', index=False)
                    # (optional) include a note sheet
                    pd.DataFrame({"Issue": [f"Missing expected column: {short_col}"]}).to_excel(writer, sheet_name='Issue', index=False)
                excel_data = output.getvalue()

                st.error(f"❌ Data Referensi dimuat tetapi kolom '{short_col}' tidak ada. Lihat file diagnostik.")
                st.download_button(
                    label="📥 Download Diagnostic Reference File",
                    data=excel_data,
                    file_name=f"diagnostic_reference_columns_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            ref_df['Kode Satker'] = ref_df['Kode Satker'].astype(str)
            st.session_state.reference_df = ref_df
            st.info(f"📚 Data Referensi dimuat otomatis ({len(ref_df)} baris).")
        except Exception as e:
            pass 

    # ============================================================
    # ✅ Then load data from GitHub (files can now be merged cleanly)
    # ============================================================
    if not st.session_state.get("data_storage"):
        with st.spinner("🔄 Memuat data dari GitHub..."):
            try:
                load_data_from_github()
            except Exception as e:
                st.error(f"⚠️ Gagal memuat data dari GitHub: {e}")
    
    
    if 'data_dipa_by_year' not in st.session_state:
        with st.spinner("🔄 Memuat data DIPA dari GitHub..."):
            try:
                load_data_dipa_from_github()
                if 'data_dipa_by_year' in st.session_state:
                    st.info(f"📥 Data DIPA dimuat untuk tahun: {', '.join(map(str, st.session_state.data_dipa_by_year.keys()))}")
            except Exception as e:
                st.warning(f"⚠️ Gagal memuat data DIPA otomatis: {e}")


    # ===============================
    # 🔹 Sidebar Navigation 
    # ===============================
    st.sidebar.title("🧭 Navigasi")
    st.sidebar.markdown("---")

    # Inisialisasi page sekali saja
    if "page" not in st.session_state:
        st.session_state.page = "📊 Dashboard Utama"

    # Pastikan page aman (fallback jika terjadi glitch)
    st.session_state.page = st.session_state.get("page", "📊 Dashboard Utama")

    # Radio navigation (Streamlit akan otomatis update session_state["page"])
    selected_page = st.sidebar.radio(
        "Pilih Halaman",
        options=[
            "📊 Dashboard Utama",
            "📈 Dashboard Internal",
            "🔐 Admin"
        ],
        key="page"   # gunakan key yg sama
    )

    st.sidebar.markdown("---")
    st.sidebar.info("""
    **Dashboard IKPA**  
    Indikator Kinerja Pelaksanaan Anggaran  
    KPPN Baturaja

    📧 Support: ameer.noor@kemenkeu.go.id
    """)


    # ===============================
    # 🔹 Routing Halaman
    # ===============================
    if st.session_state.page == "📊 Dashboard Utama":
        page_dashboard()

    elif st.session_state.page == "📈 Dashboard Internal":
        page_trend()

    elif st.session_state.page == "🔐 Admin":
        page_admin()

# ===============================
# 🔹 ENTRY POINT
# ===============================
if __name__ == "__main__":
    main()
