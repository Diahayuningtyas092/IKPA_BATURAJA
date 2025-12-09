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
    if pd.isna(k):
        return ''
    s = str(k).strip()
    digits = re.findall(r'\d+', s)
    if not digits:
        return ''
    kod = digits[0].zfill(width)
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

# ============================================================
# 🔧 FUNGSI HELPER: Load Data DIPA dari GitHub
# ============================================================
def load_DATA_DIPA_from_github():
    """
    Load semua file DIPA dari folder DATA_DIPA di GitHub.
    """
    try:
        token = st.secrets.get("GITHUB_TOKEN")
        repo_name = st.secrets.get("GITHUB_REPO")

        if not token or not repo_name:
            st.warning("⚠️ GitHub credentials tidak ditemukan")
            return

        g = Github(auth=Auth.Token(token))
        repo = g.get_repo(repo_name)


        try:
            contents = repo.get_contents("DATA_DIPA")
        except Exception as e:
            st.info(f"📁 Folder DATA_DIPA belum ada atau kosong")
            return

        if not isinstance(contents, list):
            contents = [contents]

        if "DATA_DIPA_by_year" not in st.session_state:
            st.session_state.DATA_DIPA_by_year = {}

        loaded_count = 0

        for content_file in contents:
            if content_file.type == "file" and content_file.name.lower().endswith(('.xlsx', '.xls')):
                filename = content_file.name
                year_match = re.search(r'(\d{4})', filename)
                
                if not year_match:
                    continue

                year = int(year_match.group(1))

                try:
                    file_content = repo.get_contents(content_file.path)
                    file_data = base64.b64decode(file_content.content)
                    df = pd.read_excel(io.BytesIO(file_data))

                    # Urutan kolom yang benar
                    desired_columns = [
                        "Kode Satker", "Satker", "Tahun", "Tanggal Posting Revisi",
                        "Total Pagu", "Jenis Satker", "NO", "Kementerian",
                        "Kode Status History", "Jenis Revisi", "Revisi ke-",
                        "No Dipa", "Tanggal Dipa", "Owner", "Digital Stamp"
                    ]
                    
                    available_cols = [col for col in desired_columns if col in df.columns]
                    if available_cols:
                        df = df[available_cols]
                    
                    if "Kode Satker" in df.columns:
                        df["Kode Satker"] = df["Kode Satker"].astype(str).apply(normalize_kode_satker)
                    
                    if "Total Pagu" in df.columns:
                        df["Total Pagu"] = pd.to_numeric(df["Total Pagu"], errors="coerce").fillna(0).astype(int)
                    
                    if "Revisi ke-" in df.columns:
                        df["Revisi ke-"] = pd.to_numeric(df["Revisi ke-"], errors="coerce").fillna(0).astype(int)
                    
                    if "Tahun" in df.columns:
                        df["Tahun"] = pd.to_numeric(df["Tahun"], errors="coerce").fillna(year).astype(int)
                    
                    if "Tanggal Posting Revisi" in df.columns:
                        df["Tanggal Posting Revisi"] = pd.to_datetime(df["Tanggal Posting Revisi"], errors="coerce")

                    st.session_state.DATA_DIPA_by_year[year] = df
                    loaded_count += 1
                    
                except Exception as e:
                    st.warning(f"⚠️ Gagal load {filename}: {e}")
                    continue

        if loaded_count > 0:
            years_loaded = sorted(st.session_state.DATA_DIPA_by_year.keys())
            st.success(f"✅ Load {loaded_count} file DIPA: {', '.join(map(str, years_loaded))}")

    except Exception as e:
        st.error(f"❌ Error load DIPA: {e}")


# Fungsi untuk memproses file Excel
def process_excel_file(uploaded_file, year):
    """
    Memproses file Excel IKPA sesuai struktur yang telah ditentukan
    """
    try:
        df_raw = pd.read_excel(uploaded_file, header=None)
        df_raw = df_raw.loc[:, ~df_raw.columns.str.contains('^Unnamed')]

        
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
# ======================================================================================
# PROCESS IKPA
# ======================================================================================
# ======================================================================================
# DETECT DIPA HEADER (ROBUST VERSION)
# ======================================================================================
def detect_dipa_header(uploaded_file):
    """
    Auto-detect header row dalam file DIPA mentah.
    Returns: DataFrame dengan header yang sudah benar
    """
    try:
        uploaded_file.seek(0)
        
        # Baca 20 baris pertama untuk preview
        preview = pd.read_excel(uploaded_file, header=None, nrows=20, dtype=str)
        
        # Keywords yang PASTI ada di header DIPA
        header_keywords = [
            "satker", "kode", "pagu", "jumlah", "dipa", 
            "tanggal", "revisi", "no", "status"
        ]
        
        header_row = None
        max_matches = 0
        
        # Cari baris dengan keyword terbanyak
        for i in range(len(preview)):
            row_text = " ".join(preview.iloc[i].fillna("").astype(str).str.lower())
            matches = sum(1 for kw in header_keywords if kw in row_text)
            
            if matches > max_matches:
                max_matches = matches
                header_row = i
        
        # Jika tidak ada yang cocok, gunakan baris 0
        if header_row is None or max_matches < 3:
            st.warning("⚠️ Header otomatis tidak terdeteksi, menggunakan baris pertama")
            header_row = 0
        else:
            st.info(f"✅ Header terdeteksi di baris {header_row + 1} (keyword match: {max_matches})")
        
        # Baca ulang dengan header yang benar
        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, header=header_row, dtype=str)
        
        # Bersihkan nama kolom
        df.columns = (
            df.columns.astype(str)
            .str.replace("\n", " ", regex=False)
            .str.replace("\r", " ", regex=False)
            .str.replace("\s+", " ", regex=True)
            .str.strip()
            .str.upper()  # Normalize to uppercase
        )
        
        # Hapus baris kosong
        df = df.dropna(how='all').reset_index(drop=True)
        
        # Debug: tampilkan preview
        st.write("**Preview 5 baris pertama setelah deteksi header:**")
        st.dataframe(df.head(5))
        
        return df
        
    except Exception as e:
        st.error(f"❌ Error deteksi header: {e}")
        uploaded_file.seek(0)
        return pd.read_excel(uploaded_file, dtype=str)


# ======================================================================================
# CLEAN DIPA (SUPER ROBUST VERSION)
# ======================================================================================
def clean_dipa(df_raw):
    """
    Membersihkan file DIPA mentah dan mengembalikan format standar.
    """
    
    df = df_raw.copy()
    
    # Hapus kolom Unnamed
    df = df.loc[:, ~df.columns.astype(str).str.contains("Unnamed", case=False)]
    
    # Normalize column names untuk matching yang lebih mudah
    df.columns = df.columns.astype(str).str.upper().str.strip()
    
    st.write("**Kolom yang terdeteksi di file DIPA:**")
    st.write(list(df.columns))
    
    # ====== HELPER: Flexible column finder ======
    def find_col(keywords_list):
        """Find column by multiple possible names"""
        for col in df.columns:
            col_clean = str(col).upper().replace(" ", "").replace("_", "")
            for kw in keywords_list:
                kw_clean = kw.upper().replace(" ", "").replace("_", "")
                if kw_clean in col_clean:
                    return col
        return None
    
    # ====== 1. KODE SATKER & NAMA SATKER ======
    satker_col = find_col([
        "SATKER", "KODESATKER", "KODE SATKER", "KODE SATUAN KERJA",
        "SATUAN KERJA", "SATKER/KPPN"
    ])
    
    if satker_col is None:
        st.error("❌ Kolom Satker tidak ditemukan!")
        st.write("Kolom available:", list(df.columns))
        raise ValueError("Kolom Satker tidak ditemukan")
    
    st.success(f"✅ Kolom Satker ditemukan: {satker_col}")
    
    # Extract kode 6 digit
    df["Kode Satker"] = (
        df[satker_col].astype(str)
        .str.extract(r"(\d{6})", expand=False)
        .fillna("")
    )
    
    # Extract nama satker (remove kode di depan)
    df["Satker"] = (
        df[satker_col].astype(str)
        .str.replace(r"^\d{6}\s*-?\s*", "", regex=True)
        .str.strip()
    )
    
    # Jika nama kosong, gunakan format default
    mask_empty = (df["Satker"] == "") | (df["Satker"].isna())
    df.loc[mask_empty, "Satker"] = df.loc[mask_empty, "Kode Satker"] + " - SATKER"
    
    # ====== 2. TOTAL PAGU ======
    pagu_col = find_col([
        "PAGU", "JUMLAH", "TOTAL PAGU", "TOTALPAGU", "NILAI PAGU",
        "PAGU DIPA", "JUMLAH DIPA"
    ])
    
    if pagu_col:
        st.success(f"✅ Kolom Pagu ditemukan: {pagu_col}")
        df["Total Pagu"] = pd.to_numeric(df[pagu_col], errors="coerce").fillna(0).astype(int)
    else:
        st.warning("⚠️ Kolom Pagu tidak ditemukan, menggunakan 0")
        df["Total Pagu"] = 0
    
    # ====== 3. TANGGAL POSTING REVISI ======
    tgl_col = find_col([
        "TANGGAL POSTING", "TGL POSTING", "TANGGALPOSTING",
        "TANGGAL REVISI", "TGL REVISI", "TANGGAL", "DATE"
    ])
    
    if tgl_col:
        st.success(f"✅ Kolom Tanggal ditemukan: {tgl_col}")
        df["Tanggal Posting Revisi"] = pd.to_datetime(df[tgl_col], errors="coerce")
    else:
        st.warning("⚠️ Kolom Tanggal tidak ditemukan")
        df["Tanggal Posting Revisi"] = pd.NaT
    
    # Extract Tahun
    df["Tahun"] = df["Tanggal Posting Revisi"].dt.year
    df["Tahun"] = df["Tahun"].fillna(datetime.now().year).astype(int)
    
    # ====== 4. KOLOM LAINNYA (OPTIONAL) ======
    
    # NO
    no_col = find_col(["NO", "NOMOR", "NO."])
    df["NO"] = df[no_col].astype(str).str.strip() if no_col else ""
    
    # KEMENTERIAN
    kementerian_col = find_col(["KEMENTERIAN", "KEMENTRIAN", "KL", "K/L", "BA"])
    df["Kementerian"] = df[kementerian_col].astype(str).str.strip() if kementerian_col else ""
    
    # KODE STATUS HISTORY
    status_col = find_col(["STATUS HISTORY", "KODE STATUS", "KODESTATUS", "STATUS"])
    df["Kode Status History"] = df[status_col].astype(str).str.strip() if status_col else ""
    
    # JENIS REVISI
    jenis_col = find_col(["JENIS REVISI", "JENISREVISI", "TIPE REVISI"])
    df["Jenis Revisi"] = df[jenis_col].astype(str).str.strip() if jenis_col else ""
    
    # REVISI KE-
    revisi_col = find_col(["REVISI KE", "REVISIKE", "REVISI"])
    if revisi_col:
        df["Revisi ke-"] = pd.to_numeric(df[revisi_col], errors="coerce").fillna(0).astype(int)
    else:
        df["Revisi ke-"] = 0
    
    # NO DIPA
    nodipa_col = find_col(["NO DIPA", "NODIPA", "NOMOR DIPA", "NO. DIPA"])
    df["No Dipa"] = df[nodipa_col].astype(str).str.strip() if nodipa_col else ""
    
    # TANGGAL DIPA
    tgldipa_col = find_col(["TANGGAL DIPA", "TGL DIPA", "TGLDIPA"])
    if tgldipa_col:
        df["Tanggal Dipa"] = pd.to_datetime(df[tgldipa_col], errors="coerce").dt.strftime("%d-%m-%Y")
    else:
        df["Tanggal Dipa"] = ""
    
    # OWNER
    owner_col = find_col(["OWNER", "PEMILIK"])
    df["Owner"] = df[owner_col].astype(str).str.strip() if owner_col else ""
    
    # DIGITAL STAMP
    stamp_col = find_col(["DIGITAL STAMP", "DIGITALSTAMP", "STAMP", "TTD DIGITAL"])
    df["Digital Stamp"] = df[stamp_col].astype(str).str.strip() if stamp_col else ""
    
    # ====== FINAL: Susun kolom sesuai urutan ======
    final_columns = [
        "Kode Satker", "Satker", "Tahun", "Tanggal Posting Revisi", "Total Pagu",
        "NO", "Kementerian", "Kode Status History", "Jenis Revisi", "Revisi ke-",
        "No Dipa", "Tanggal Dipa", "Owner", "Digital Stamp"
    ]
    
    df_clean = df[final_columns].copy()
    
    # Ambil revisi terakhir per satker per tahun
    df_clean = df_clean.sort_values(["Kode Satker", "Tahun", "Tanggal Posting Revisi"])
    df_clean = df_clean.groupby(["Kode Satker", "Tahun"], as_index=False).tail(1)
    
    # Filter hanya kode satker yang valid (6 digit)
    df_clean = df_clean[df_clean["Kode Satker"].str.len() == 6]
    
    st.write(f"**Hasil cleaning: {len(df_clean)} baris satker valid**")
    
    return df_clean


# ======================================================================================
# ASSIGN JENIS SATKER
# ======================================================================================
def assign_jenis_satker(df):
    """Klasifikasi satker berdasarkan Total Pagu"""
    
    if df.empty or "Total Pagu" not in df.columns:
        df["Jenis Satker"] = "Satker Kecil"
        return df
    
    q70 = df["Total Pagu"].quantile(0.70)
    q40 = df["Total Pagu"].quantile(0.40)
    
    def classify(pagu):
        if pagu >= q70: return "Satker Besar"
        elif pagu >= q40: return "Satker Sedang"
        else: return "Satker Kecil"
    
    df["Jenis Satker"] = df["Total Pagu"].apply(classify)
    
    # Reorder: Jenis Satker setelah Total Pagu
    cols = list(df.columns)
    if "Jenis Satker" in cols and "Total Pagu" in cols:
        cols.remove("Jenis Satker")
        pagu_idx = cols.index("Total Pagu")
        cols.insert(pagu_idx + 1, "Jenis Satker")
        df = df[cols]
    
    return df


# ======================================================================================
# PROCESS UPLOADED DIPA (MAIN FUNCTION)
# ======================================================================================
def process_uploaded_dipa(uploaded_file, save_file_to_github):
    """Process DIPA file dengan error handling lengkap"""
    
    try:
        st.info("🔄 Memulai proses DIPA...")
        
        # 1️⃣ Deteksi header
        with st.spinner("Mendeteksi header..."):
            raw = detect_dipa_header(uploaded_file)
        
        if raw.empty:
            return None, None, "❌ File kosong setelah deteksi header"
        
        # 2️⃣ Clean data
        with st.spinner("Membersihkan data..."):
            clean = clean_dipa(raw)
        
        if clean.empty:
            return None, None, "❌ Tidak ada data valid setelah cleaning"
        
        # 3️⃣ Merge dengan referensi
        if "reference_df" in st.session_state and not st.session_state.reference_df.empty:
            with st.spinner("Menggabungkan dengan data referensi..."):
                ref = st.session_state.reference_df.copy()
                ref["Kode Satker"] = ref["Kode Satker"].apply(normalize_kode_satker)
                clean["Kode Satker"] = clean["Kode Satker"].apply(normalize_kode_satker)
                
                clean = clean.merge(
                    ref[["Kode BA", "K/L", "Kode Satker"]],
                    on="Kode Satker",
                    how="left"
                )
                
                if "Kementerian" in clean.columns and "K/L" in clean.columns:
                    clean["Kementerian"] = clean["Kementerian"].fillna(clean["K/L"])
        
        # 4️⃣ Klasifikasi jenis satker
        with st.spinner("Mengklasifikasi jenis satker..."):
            clean = assign_jenis_satker(clean)
        
        # 5️⃣ Extract tahun
        tahun_dipa = int(clean["Tahun"].mode()[0]) if "Tahun" in clean.columns else datetime.now().year
        
        # 6️⃣ Save per tahun
        if "DATA_DIPA_by_year" not in st.session_state:
            st.session_state.DATA_DIPAa_by_year = {}
        
        for yr in clean["Tahun"].unique():
            yr = int(yr)
            df_year = clean[clean["Tahun"] == yr].copy()
            
            st.session_state.DATA_DIPA_by_year[yr] = df_year
            
            # Save to GitHub
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine="openpyxl") as writer:
                df_year.to_excel(writer, index=False, sheet_name=f"DIPA_{yr}")
                
                # Format header
                try:
                    wb = writer.book
                    ws = writer.sheets[f"DIPA_{yr}"]
                    
                    for cell in ws[1]:
                        cell.font = Font(bold=True, color="FFFFFF")
                        cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                except:
                    pass
            
            out.seek(0)
            save_file_to_github(out.getvalue(), f"DIPA_{yr}.xlsx", "DATA_DIPA")
            
        return clean, tahun_dipa, "✅ Sukses"
        
    except Exception as e:
        error_msg = f"❌ Error: {str(e)}"
        st.error(error_msg)
        import traceback
        st.code(traceback.format_exc())
        return None, None, error_msg
    
# ------------------------------------------------------------
# PAGE ADMIN
# ------------------------------------------------------------
def page_admin():
    st.title("🔐 Halaman Administrasi")

    # ============================================================
    # 🔑 LOGIN ADMIN
    # ============================================================
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False

    if not st.session_state.authenticated:
        st.warning("🔒 Halaman ini memerlukan autentikasi Admin")
        password = st.text_input("Masukkan Password Admin", type="password")
        if st.button("Login"):
            if password == "109KPPN":
                st.session_state.authenticated = True
                st.success("✔ Login berhasil")
                st.rerun()
            else:
                st.error("❌ Password salah")
        return

    st.success("✔ Anda login sebagai Admin")

    st.markdown("---")

    # ============================================================
    # 📌 TAB MENU
    # ============================================================
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

                    # Mapping bulan lengkap
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

                    for uploaded_file in uploaded_files:
                        try:
                            # ==========================
                            # 1️⃣ DETEKSI BULAN AMAN
                            # ==========================
                            uploaded_file.seek(0)
                            df_temp = pd.read_excel(uploaded_file, header=None, dtype=str)
                            raw_text = " ".join(df_temp.fillna("").astype(str).values.flatten()).upper()

                            detected_month = None
                            for k, v in MONTH_FIX.items():
                                if k in raw_text:
                                    detected_month = v
                                    break

                            # fallback → dari nama file
                            if detected_month is None:
                                filename = uploaded_file.name.upper()
                                clean = re.sub(r"[^A-Z_]", "", filename)
                                parts = clean.split("_")
                                for p in parts:
                                    if p in MONTH_FIX:
                                        detected_month = MONTH_FIX[p]
                                        break

                            if detected_month is None:
                                detected_month = "UNKNOWN"

                            # ==========================
                            # 2️⃣ PROSES FILE IKPA
                            # ==========================
                            uploaded_file.seek(0)
                            df_processed, _, _ = process_excel_file(uploaded_file, upload_year)

                            if df_processed is None:
                                st.warning(f"⚠️ Gagal memproses file: {uploaded_file.name}")
                                continue

                            df_processed["Bulan"] = detected_month
                            df_processed["Tahun"] = int(upload_year)

                            # ==========================
                            # 3️⃣ NORMALISASI KODE SATKER
                            # ==========================
                            if "Kode Satker" in df_processed.columns:
                                df_processed["Kode Satker"] = df_processed["Kode Satker"].astype(str).apply(normalize_kode_satker)
                            else:
                                df_processed["Kode Satker"] = ""

                            # ==========================
                            # 4️⃣ MERGE IKPA + DIPA
                            # ==========================
                            df_final = df_processed.copy()

                            dipa_year = None
                            if "DATA_DIPA_by_year" in st.session_state:
                                dipa_year = st.session_state.DATA_DIPA_by_year.get(int(upload_year))

                            if dipa_year is not None and not dipa_year.empty:

                                dipa_year = dipa_year.copy()
                                dipa_year["Kode Satker"] = dipa_year["Kode Satker"].apply(normalize_kode_satker)

                                required_cols = ["Kode Satker", "Total Pagu", "Tanggal Posting Revisi", "Jenis Satker"]
                                dipa_year = dipa_year[[c for c in required_cols if c in dipa_year.columns]]

                                df_final = df_final.merge(
                                    dipa_year.rename(columns={"Total Pagu": "Total Pagu DIPA"}),
                                    on="Kode Satker",
                                    how="left"
                                )

                            else:
                                df_final["Jenis Satker"] = pd.NA
                                df_final["Total Pagu DIPA"] = pd.NA
                                df_final["Tanggal Posting Revisi"] = pd.NA

                            # ==========================
                            # 5️⃣ SIMPAN SESSION STATE
                            # ==========================
                            key = (detected_month, str(upload_year))
                            st.session_state.data_storage[key] = df_final.copy()

                            # ==========================
                            # 6️⃣ EXPORT KE GITHUB
                            # ==========================
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
                                "Total Pagu DIPA",
                                "Tanggal Posting Revisi",
                                "Jenis Satker",
                                "Bulan",
                                "Tahun"
                            ]

                            df_export = df_final[[c for c in KEEP_COLUMNS if c in df_final.columns]]

                            excel_bytes = io.BytesIO()
                            with pd.ExcelWriter(excel_bytes, engine="openpyxl") as writer:
                                df_export.to_excel(writer, index=False, sheet_name="Data IKPA")
                            excel_bytes.seek(0)

                            save_file_to_github(
                                excel_bytes.getvalue(),
                                f"IKPA_{detected_month}_{upload_year}.xlsx",
                                folder="data"
                            )

                            # LOG
                            st.session_state.activity_log.append({
                                "Waktu": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                "Aksi": "Upload",
                                "Periode": f"{detected_month} {upload_year}",
                                "Status": "Sukses"
                            })

                            st.success(f"✅ {uploaded_file.name} → {detected_month} {upload_year} berhasil disimpan.")

                        except Exception as e:
                            st.error(f"❌ Error saat memproses data: {e}")

        # ============================================================
        # SUBMENU: UPLOAD DATA DIPA
        # ============================================================
        st.markdown("---")
        st.subheader("📤 Upload Data DIPA")

        uploaded_dipa_file = st.file_uploader(
            "Pilih file Excel DIPA (mentah dari SAS/SMART/Kemenkeu)",
            type=['xlsx', 'xls'],
            key="upload_dipa"
        )

        # Tombol proses DIPA
        if uploaded_dipa_file is not None:
            if st.button("🔄 Proses Data DIPA", type="primary"):
                with st.spinner("Memproses data DIPA..."):

                    try:
                        # 1️⃣ Proses file raw DIPA → dibersihkan → revisi terbaru
                        df_clean, tahun_dipa, status_msg = process_uploaded_dipa(uploaded_dipa_file, save_file_to_github)

                        if df_clean is None:
                            st.error(f"❌ Gagal memproses DIPA: {status_msg}")
                            st.stop()

                        # 2️⃣ Pastikan kolom Kode Satker distandardkan
                        df_clean["Kode Satker"] = df_clean["Kode Satker"].astype(str).apply(normalize_kode_satker)

                        # 3️⃣ Simpan ke session_state per tahun
                        if "DATA_DIPA_by_year" not in st.session_state:
                            st.session_state.DATA_DIPA_by_year = {}

                        st.session_state.DATA_DIPA_by_year[int(tahun_dipa)] = df_clean.copy()

                        # 4️⃣ Simpan ke GitHub dalam folder `DATA_DIPA`
                        excel_bytes = io.BytesIO()
                        with pd.ExcelWriter(excel_bytes, engine='openpyxl') as writer:
                            df_clean.to_excel(writer, index=False, sheet_name=f"DIPA_{tahun_dipa}")

                        excel_bytes.seek(0)

                        save_file_to_github(
                            excel_bytes.getvalue(),
                            f"DIPA_CLEAN_{tahun_dipa}.xlsx",
                            folder="DATA_DIPA"
                        )

                        # 5️⃣ Catat log
                        st.session_state.activity_log.append({
                            "Waktu": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "Aksi": "Upload DIPA",
                            "Periode": f"Tahun {tahun_dipa}",
                            "Status": "Sukses"
                        })

                        # 6️⃣ Tampilkan hasil preview
                        st.success(f"✅ Data DIPA tahun {tahun_dipa} berhasil diproses & disimpan.")
                        st.dataframe(df_clean.head(10), use_container_width=True)

                    except Exception as e:
                        st.error(f"❌ Terjadi error saat memproses file DIPA: {e}")

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
        if not st.session_state.get("DATA_DIPA_by_year"):
            st.info("ℹ️ Belum ada data DIPA tersimpan.")
        else:
            available_years = sorted(st.session_state.DATA_DIPA_by_year.keys(), reverse=True)
            year_to_delete = st.selectbox(
                "Pilih tahun DIPA yang akan dihapus",
                options=available_years,
                format_func=lambda x: f"Tahun {x}",
                key="delete_dipa_year"
            )
            filename_dipa = f"DATA_DIPA/DIPA_{year_to_delete}.xlsx"

            confirm_delete_dipa = st.checkbox(
                f"⚠️ Hapus data DIPA tahun {year_to_delete} dari sistem dan GitHub.",
                key=f"confirm_delete_dipa_{year_to_delete}"
            )

            if st.button("🗑️ Hapus Data DIPA Ini", type="primary", key="btn_delete_dipa") and confirm_delete_dipa:
                try:
                    del st.session_state.DATA_DIPA_by_year[year_to_delete]
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
    # TAB 3: DOWNLOAD DATA
    # ============================================================
    with tab3:
        # DOWNLOAD DATA IKPA
        st.subheader("📥 Download Data IKPA")

        # Jika belum ada data IKPA sama sekali
        if "data_storage" not in st.session_state or not st.session_state.data_storage:
            st.info("ℹ️ Belum ada data IKPA.")
        else:
            # Tampilkan periode yang tersedia
            available_periods = sorted(st.session_state.data_storage.keys(), reverse=True)

            period_to_download = st.selectbox(
                "Pilih periode untuk download",
                options=available_periods,
                format_func=lambda x: f"{x[0].capitalize()} {x[1]}"
            )

            df_download = st.session_state.data_storage[period_to_download].copy()

            # Siapkan file Excel untuk di-download
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_excel = df_download.drop(
                    ['Bobot', 'Nilai Terbobot'], axis=1, errors='ignore'
                )
                df_excel.to_excel(writer, index=False, sheet_name='Data IKPA')

                # Format header agar cantik
                try:
                    workbook = writer.book
                    worksheet = writer.sheets['Data IKPA']
                    for cell in worksheet[1]:
                        cell.font = Font(bold=True, color="FFFFFF")
                        cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                except Exception:
                    pass

            output.seek(0)

            st.download_button(
                label="📥 Download Excel IKPA",
                data=output,
                file_name=f"IKPA_{period_to_download[0]}_{period_to_download[1]}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
 
        # Submenu Download Data DIPA
        st.markdown("### 📥 Download Data DIPA")

        if not st.session_state.get("DATA_DIPA_by_year"):
            st.info("ℹ️ Belum ada data DIPA.")
        else:
            available_years_download = sorted(
                st.session_state.DATA_DIPA_by_year.keys(), 
                reverse=True
            )
            year_to_download = st.selectbox(
                "Pilih tahun DIPA untuk download",
                options=available_years_download,
                format_func=lambda x: f"Tahun {x}",
                key="download_dipa_year"
            )

            # Ambil data DIPA yang sudah bersih
            df_download_dipa = st.session_state.DATA_DIPA_by_year[year_to_download].copy()

            # ✅ Pastikan urutan kolom sesuai
            desired_columns = [
                "Kode Satker",
                "Satker",
                "Tahun",
                "Tanggal Posting Revisi",
                "Total Pagu",
                "Jenis Satker",
                "NO",
                "Kementerian",
                "Kode Status History",
                "Jenis Revisi",
                "Revisi ke-",
                "No Dipa",
                "Tanggal Dipa",
                "Owner",
                "Digital Stamp"
            ]
            
            # Ambil hanya kolom yang ada
            available_cols = [col for col in desired_columns if col in df_download_dipa.columns]
            df_download_dipa = df_download_dipa[available_cols]

            # Preview
            with st.expander("👁️ Preview Data (5 baris pertama)"):
                st.dataframe(df_download_dipa.head(5), use_container_width=True)

            # Export to Excel
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
                try:
                    workbook = writer.book
                    worksheet = writer.sheets[f'DIPA_{year_to_download}']

                    for cell in worksheet[1]:
                        cell.font = Font(bold=True, color="FFFFFF")
                        cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                except:
                    pass

            output_dipa.seek(0)
            
            st.download_button(
                label="📥 Download Excel DIPA",
                data=output_dipa,
                file_name=f"DIPA_{year_to_download}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_download_dipa"
            )

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
    
    
    if 'DATA_DIPA_by_year' not in st.session_state:
        with st.spinner("🔄 Memuat data DIPA dari GitHub..."):
            try:
                load_DATA_DIPA_from_github()
                if 'DATA_DIPA_by_year' in st.session_state:
                    st.info(f"📥 Data DIPA dimuat untuk tahun: {', '.join(map(str, st.session_state.DATA_DIPA_by_year.keys()))}")
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
