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
    "JANUARI": 1, "FEBRUARI": 2, "PEBRUARI": 2, "MARET": 3, "APRIL": 4, "MEI": 5, "JUNI": 6,
    "JULI": 7, "AGUSTUS": 8, "SEPTEMBER": 9, "OKTOBER": 10, 
    "NOVEMBER": 11, "NOPEMBER": 11, "DESEMBER": 12
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

#fungsi untuk menormalisasi kode satker
def normalize_kode_satker(kode: str) -> str:
    """Normalize Kode Satker to always 6 digits, keep leading zeros, add apostrophe if needed."""
    if pd.isna(kode): return ''
    kode_str = str(kode).strip().lstrip("'")
    kode_str = ''.join(ch for ch in kode_str if ch.isdigit())
    if len(kode_str) < 6:
        kode_str = kode_str.zfill(6)
    elif len(kode_str) > 6:
        kode_str = kode_str[-6:]
    return f"'{kode_str}" if kode_str.startswith("0") else kode_str

# Fungsi untuk memproses file Excel
def process_excel_file(uploaded_file, year):
    """
    Memproses file Excel IKPA sesuai struktur yang telah ditentukan
    """
    try:
        df_raw = pd.read_excel(uploaded_file, header=None)
        
        # 1️⃣ Ekstrak bulan dari baris ke-2 (index 1)
        month_text = str(df_raw.iloc[1, 0])
        month = month_text.split(":")[-1].strip() if ":" in month_text else "UNKNOWN"
        
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
                'Bulan': month, 'Tahun': year,
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

# Load all uploaded data from GitHub (run on startup)
def load_data_from_github():
    """
    Load all IKPA Excel files from the GitHub 'data' folder into session state.
    Each file name must follow 'IKPA_<Bulan>_<Tahun>.xlsx'.
    Normalizes 'Kode Satker' using the global helper before merging.
    """
    token = st.secrets.get("GITHUB_TOKEN")
    repo_name = st.secrets.get("GITHUB_REPO")

    # Ensure credentials exist
    if not token or not repo_name:
        st.error("❌ Gagal mengakses GitHub: GITHUB_TOKEN atau GITHUB_REPO tidak ditemukan di secrets.")
        st.stop()
        return

    g = Github(auth=Auth.Token(token))
    repo = g.get_repo(repo_name)

    try:
        contents = repo.get_contents("data")
    except Exception:
        st.info("📁 Folder 'data' belum ada di repository GitHub.")
        return

    st.session_state.data_storage = {}

    # 🔄 Process all Excel files from GitHub
    for file in contents:
        if not file.name.endswith(".xlsx"):
            continue

        decoded = base64.b64decode(file.content)
        df = pd.read_excel(io.BytesIO(decoded))

        # Parse period (expects filenames like IKPA_JANUARI_2025.xlsx)
        parts = file.name.replace("IKPA_", "").replace(".xlsx", "").split("_")
        if len(parts) != 2:
            continue
        month, year = parts

        # Ensure period fields exist before normalization
        df['Bulan'] = df.get('Bulan', month)
        df['Tahun'] = df.get('Tahun', year)

        # 🧩 Normalize Kode Satker (using global helper)
        if 'Kode Satker' in df.columns:
            df['Kode Satker'] = df['Kode Satker'].apply(normalize_kode_satker)
        else:
            df['Kode Satker'] = ''

        # Apply references & cleanup
        df = apply_reference_short_names(df)
        df = create_satker_column(df)

        # Ensure numeric columns are numeric
        numeric_cols = [
            'Nilai Akhir (Nilai Total/Konversi Bobot)', 'Nilai Total', 'Konversi Bobot',
            'Revisi DIPA', 'Deviasi Halaman III DIPA', 'Penyerapan Anggaran',
            'Belanja Kontraktual', 'Penyelesaian Tagihan',
            'Pengelolaan UP dan TUP', 'Capaian Output'
        ]
        for col in numeric_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        # Add helper columns for sorting and source tracking
        df['Source'] = 'GitHub'
        df['Period'] = f"{month} {year}"
        month_num = MONTH_ORDER.get(month.upper(), 0)
        df['Period_Sort'] = f"{int(year):04d}-{month_num:02d}"

        # Generate ranking if missing
        if 'Peringkat' not in df.columns and 'Nilai Akhir (Nilai Total/Konversi Bobot)' in df.columns:
            df = df.sort_values('Nilai Akhir (Nilai Total/Konversi Bobot)', ascending=False)
            df['Peringkat'] = range(1, len(df) + 1)

        # Store to session state
        st.session_state.data_storage[(str(month), str(year))] = df

    st.success(f"✅ {len(st.session_state.data_storage)} file berhasil dimuat dari GitHub.")

# Fungsi untuk membaca template Excel yang sudah ada
def get_template_file():
    """
    Membaca file template Excel yang sudah ada
    """
    try:
        # Cek apakah file template ada di path default
        if Path(TEMPLATE_PATH).exists():
            with open(TEMPLATE_PATH, 'rb') as f:
                return f.read()
        else:
            # Jika tidak ada, gunakan template dari session state (jika di-upload admin)
            if 'template_file' in st.session_state:
                return st.session_state.template_file
            else:
                return None
    except Exception as e:
        st.error(f"Error membaca template: {str(e)}")
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
/* Target utama */
div[data-testid="stPopover"] button {
    background-color: #FFF9E6 !important;
    border: 1px solid #E6C200 !important;
    color: #664400 !important;
}

/* Hover */
div[data-testid="stPopover"] button:hover {
    background-color: #FFE4B5 !important;
    color: black !important;
}

/* Tambahan: jika Streamlit mengubah struktur DOM */
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

    
    if not st.session_state.data_storage:
        st.warning("⚠️ Belum ada data yang diunggah. Silakan unggah data melalui halaman Admin.")
        return
    
    # Dapatkan data terbaru
    all_periods = sorted(
        st.session_state.data_storage.keys(),
        key=lambda x: (int(x[1]), MONTH_ORDER.get(x[0].upper(), 0)),
        reverse=True
    )
    
    if not all_periods:
        st.warning("⚠️ Belum ada data yang tersedia.")
        return

    if "main_tab" not in st.session_state:
        st.session_state.main_tab = "highlights"

    main_tab = st.radio(
        "Pilih Bagian Dashboard",
        ["🎯 Highlights", "📋 Data Detail Satker"],
        key="main_tab_choice",
        horizontal=True
    )
    st.session_state["main_tab"] = main_tab

    # -------------------------
    # HIGHLIGHTS (sebelumnya with tab_highlights)
    # -------------------------
    if main_tab == "🎯 Highlights":
        st.markdown("## 🎯 Highlights Kinerja Satker")

        # Single-row layout for period + metrics
        col_period, col1, col2, col3, col4 = st.columns([1, 1, 1, 1, 1])

        with col_period:
            selected_period = st.selectbox(
                "Pilih Periode",
                options=all_periods,
                index=0,
                format_func=lambda x: f"{x[0].capitalize()} {x[1]}"
            )

        df = st.session_state.data_storage[selected_period]

        # Hitung data
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
            avg_score = df['Nilai Akhir (Nilai Total/Konversi Bobot)'].mean()
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
    # DATA DETAIL SATKER (sebelumnya with tab_table)
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
            st.markdown("#### Periodik — ringkasan per bulan / triwulan")

            # determine available years from all data in session
            years = set()
            for k, df_period in st.session_state.data_storage.items():
                years.update(df_period['Tahun'].astype(str).unique())
            years = sorted([int(y) for y in years if str(y).strip() != ''], reverse=True)
            if not years:
                st.info("Tidak ada data periodik untuk ditampilkan.")
            else:
                default_year = years[0]
                selected_year = st.selectbox("Pilih Tahun", options=years, index=0, key='tab_periodik_year_select')
                
                # session state untuk period type
                if "period_type" not in st.session_state:
                    st.session_state.period_type = "quarterly"

                period_options = ["quarterly", "monthly"]
                try:
                    period_index = period_options.index(st.session_state.period_type)
                except ValueError:
                    period_index = 0
                    st.session_state.period_type = "quarterly"

                # Radio button tanpa callback (Streamlit akan simpan ke session_state via key)
                period_type = st.radio(
                    "Jenis Periode",
                    options=period_options,
                    format_func=lambda x: "Triwulan" if x == "quarterly" else "Bulanan",
                    horizontal=True,
                    index=period_index,
                    key="period_type_radio_v2"
                )
                # Update session state directly
                st.session_state.period_type = period_type

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

            # build dataframe for the selected year across all available periods
            dfs = []
            for (mon, yr), df_period in st.session_state.data_storage.items():
                try:
                    if int(yr) == int(selected_year):
                        dfs.append(df_period.copy())
                except Exception:
                    continue
            if not dfs:
                st.info(f"Tidak ditemukan data untuk tahun {selected_year}.")
            else:
                df_year = pd.concat(dfs, ignore_index=True)

                # normalize month names and get available months sorted
                df_year['Bulan_raw'] = df_year['Bulan'].astype(str).fillna('').str.strip()

                month_aliases = {
                    'PEBRUARI': 'FEBRUARI', 'PEBRUARY': 'FEBRUARI', 'NOPEMBER': 'NOVEMBER',
                    'NOVEMBER ': 'NOVEMBER', 'SEPT': 'SEPTEMBER', 'SEP': 'SEPTEMBER',
                    'MAR': 'MARET', 'MRT': 'MARET'
                }
                canonical_display = {k.upper(): k.capitalize() for k in MONTH_ORDER.keys()}

                def normalize_month_text(txt):
                    t = str(txt).strip().upper()
                    t = re.sub(r'[^A-Z]', '', t)
                    if t in month_aliases:
                        return month_aliases[t]
                    if t in MONTH_ORDER:
                        return t
                    for mm in MONTH_ORDER.keys():
                        if mm.startswith(t) or mm.startswith(t[:3]):
                            return mm
                    return t

                df_year['Bulan_upper'] = df_year['Bulan_raw'].apply(normalize_month_text)

                months_available = sorted(
                    [m for m in df_year['Bulan_upper'].unique() if m and m in MONTH_ORDER],
                    key=lambda m: MONTH_ORDER.get(m, 0)
                )

                # decide period columns (monthly or quarterly)
                if period_type == 'monthly':
                    months_sorted = months_available
                    display_month_names_ordered = [canonical_display.get(m, m.capitalize()) for m in months_sorted]
                else:
                    quarter_map = {
                        'Tw I': 'MARET',
                        'Tw II': 'JUNI',
                        'Tw III': 'SEPTEMBER',
                        'Tw IV': 'DESEMBER'
                    }
                    quarter_order = []
                    for tw, end_month in quarter_map.items():
                        if end_month in months_available:
                            quarter_order.append(tw)

                # Build records for pivoting
                records = []
                for _, row in df_year.iterrows():
                    rec = {
                        'Kode BA': row.get('Kode BA', ''),
                        'Kode Satker': row.get('Kode Satker', ''),
                        'Uraian Satker-RINGKAS': row.get('Uraian Satker-RINGKAS', row.get('Uraian Satker Final', row.get('Uraian Satker','')))
                    }
                    month_up = row.get('Bulan_upper', '')
                    if period_type == 'monthly':
                        if month_up in MONTH_ORDER:
                            rec[month_up] = row.get(selected_indicator, np.nan)
                    else:
                        if month_up == 'MARET':
                            rec['Tw I'] = row.get(selected_indicator, np.nan)
                        elif month_up == 'JUNI':
                            rec['Tw II'] = row.get(selected_indicator, np.nan)
                        elif month_up == 'SEPTEMBER':
                            rec['Tw III'] = row.get(selected_indicator, np.nan)
                        elif month_up == 'DESEMBER':
                            rec['Tw IV'] = row.get(selected_indicator, np.nan)
                    records.append(rec)

                df_rec = pd.DataFrame(records)
                if df_rec.empty:
                    st.info("Tidak ada data detail untuk indikator/periode yang dipilih.")
                else:
                    # aggregate by Kode Satker (take last non-null)
                    agg_dict = {}
                    possible_period_cols = [c for c in df_rec.columns if c not in ['Kode BA','Kode Satker','Uraian Satker-RINGKAS']]
                    for c in possible_period_cols:
                        def last_non_null(x):
                            s = x.dropna()
                            return float(s.iloc[-1]) if len(s) > 0 else np.nan
                        agg_dict[c] = last_non_null

                    df_agg = df_rec.groupby(['Kode BA','Kode Satker','Uraian Satker-RINGKAS']).agg(agg_dict).reset_index()

                    # rename raw canonical month columns to display names in order
                    display_period_cols = []
                    if period_type == 'monthly':
                        raw_cols_upper = {c.upper(): c for c in df_agg.columns}
                        for m in months_sorted:
                            if m in raw_cols_upper:
                                raw_col = raw_cols_upper[m]
                                display_name = canonical_display.get(m, m.capitalize())
                                if raw_col != display_name:
                                    df_agg.rename(columns={raw_col: display_name}, inplace=True)
                                display_period_cols.append(display_name)
                    else:
                        for tw in ['Tw I','Tw II','Tw III','Tw IV']:
                            if tw in df_agg.columns:
                                display_period_cols.append(tw)

                    # drop all-NaN period columns
                    display_period_cols = [c for c in display_period_cols if not df_agg[c].isna().all()]

                    if display_period_cols:
                        last_col = display_period_cols[-1]
                        df_agg['Latest_Value'] = df_agg[last_col]
                    else:
                        df_agg['Latest_Value'] = np.nan

                    df_agg['Peringkat'] = df_agg['Latest_Value'].rank(ascending=False, method='dense').astype('Int64')
                    df_agg_sorted = df_agg.sort_values(by=['Peringkat'], ascending=False)

                    final_cols = ['Peringkat','Kode BA','Kode Satker','Uraian Satker-RINGKAS'] + display_period_cols
                    for c in final_cols:
                        if c not in df_agg_sorted.columns:
                            df_agg_sorted[c] = np.nan

                    df_display = df_agg_sorted[final_cols].copy()

                    # SEARCH widget for Periodik
                    search_query = st.text_input("🔎 Cari (Periodik) – ketik untuk filter di semua kolom", value="", key='tab_periodik_search')
                    if search_query:
                        q = str(search_query).strip().lower()
                        mask = df_display.apply(lambda row: row.astype(str).str.lower().str.contains(q, na=False).any(), axis=1)
                        df_display_filtered = df_display[mask].copy()
                    else:
                        df_display_filtered = df_display.copy()

                    # Trend cell coloring & styling
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
                        for i, c in enumerate(df_display_filtered.columns):
                            if display_period_cols and c == display_period_cols[-1]:
                                styles.append(color)
                            else:
                                styles.append('')
                        return styles

                    def highlight_top(s):
                        if s.name == 'Peringkat':
                            return ['background-color: gold' if (pd.to_numeric(v, errors='coerce') <= 3) else '' for v in s]
                        return ['' for _ in s]

                    styler = df_display_filtered.style.format(precision=2)
                    if display_period_cols:
                        styler = styler.apply(lambda r: color_trend(r), axis=1)
                    styler = styler.apply(highlight_top)
                    st.dataframe(styler, use_container_width=True, height=600)

        # -------------------------
        # DETAIL SATKER (legacy table)
        # -------------------------
        else:
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
# 🔐 HALAMAN 3: ADMIN (Revised with integrated Reference Upload)
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
    # TAB 1: UPLOAD DATA (including Reference Upload)
    # ============================================================
    with tab1:
        # Submenu Upload Data Bulanan
        st.subheader("📤 Upload Data Bulanan IKPA")

        upload_year = st.selectbox(
            "Pilih Tahun",
            list(range(2020, 2031)),
            index=list(range(2020, 2031)).index(datetime.now().year)
        )

        uploaded_file = st.file_uploader("Pilih file Excel IKPA", type=['xlsx', 'xls'])

        if uploaded_file is not None:
            try:
                df_temp = pd.read_excel(uploaded_file, header=None)
                month_text = str(df_temp.iloc[1, 0])
                month_preview = month_text.split(":")[-1].strip() if ":" in month_text else "UNKNOWN"
                period_key_preview = (str(month_preview), str(upload_year))
                uploaded_file.seek(0)

                if period_key_preview in st.session_state.data_storage:
                    st.warning(f"⚠️ Data untuk **{month_preview} {upload_year}** sudah ada.")
                    confirm_replace = st.checkbox(
                        "✅ Ganti data yang sudah ada.",
                        key=f"confirm_replace_{month_preview}_{upload_year}"
                    )
                else:
                    confirm_replace = True
                    st.info(f"📝 Akan mengunggah data baru untuk periode: **{month_preview} {upload_year}**")

            except Exception as e:
                st.error(f"❌ Gagal membaca preview file: {e}")
                confirm_replace = False

            if st.button("🔄 Proses Data IKPA", type="primary", disabled=not confirm_replace):
                with st.spinner("Memproses data..."):
                    df_processed, month, year = process_excel_file(uploaded_file, upload_year)
                    if df_processed is None:
                        st.error("❌ Gagal memproses file.")
                        st.stop()

                    # 🧩 Normalize Kode Satker before saving or matching
                    if 'Kode Satker' in df_processed.columns:
                        df_processed['Kode Satker'] = df_processed['Kode Satker'].apply(normalize_kode_satker)
                    else:
                        df_processed['Kode Satker'] = ''

                    period_key = (str(month), str(year))
                    filename = f"IKPA_{month}_{year}.xlsx"

                    try:
                        df_processed['Kode Satker'] = df_processed['Kode Satker'].astype(str)
                        st.session_state.data_storage[period_key] = df_processed

                        excel_bytes = io.BytesIO()
                        with pd.ExcelWriter(excel_bytes, engine='openpyxl') as writer:
                            df_excel = df_processed.drop(['Bobot', 'Nilai Terbobot'], axis=1, errors='ignore')
                            df_excel.to_excel(writer, index=False, sheet_name='Data IKPA')
                        excel_bytes.seek(0)

                        save_file_to_github(excel_bytes.getvalue(), filename, folder="data")

                        st.success(f"✅ Data {month} {year} berhasil disimpan.")
                        st.snow()

                        st.session_state.activity_log.append({
                            "Waktu": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "Aksi": "Upload",
                            "Periode": f"{month} {year}",
                            "Status": "✅ Sukses"
                        })
                    except Exception as e:
                        st.error(f"❌ Gagal menyimpan ke GitHub: {e}")


        # Sub Menu Upload Data Referensi
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

                # 🧩 Save merged reference data permanently to GitHub
                try:
                    excel_bytes_ref = io.BytesIO()
                    with pd.ExcelWriter(excel_bytes_ref, engine='openpyxl') as writer:
                        st.session_state.reference_df.to_excel(writer, index=False, sheet_name='Data Referensi')
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
        st.subheader("🗑️ Hapus Data Bulanan")
        if not st.session_state.data_storage:
            st.info("ℹ️ Belum ada data tersimpan.")
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

            if st.button("🗑️ Hapus Data Ini", type="primary") and confirm_delete:
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
                except Exception as e:
                    st.error(f"❌ Gagal menghapus data: {e}")

    # ============================================================
    # TAB 3: DOWNLOAD DATA
    # ============================================================
    with tab3:
        # Submenu Download Data IKPA
        st.subheader("📥 Download Data IKPA")
        if not st.session_state.data_storage:
            st.info("ℹ️ Belum ada data.")
        else:
            available_periods = sorted(st.session_state.data_storage.keys(), reverse=True)
            period_to_download = st.selectbox(
                "Pilih periode untuk download",
                options=available_periods,
                format_func=lambda x: f"{x[0].capitalize()} {x[1]}"
            )
            df_download = st.session_state.data_storage[period_to_download]
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_excel = df_download.drop(['Bobot', 'Nilai Terbobot'], axis=1, errors='ignore')
                df_excel.to_excel(writer, index=False, sheet_name='Data IKPA')
            output.seek(0)
            st.download_button(
                label="📥 Download Excel",
                data=output,
                file_name=f"IKPA_{period_to_download[0]}_{period_to_download[1]}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # --- Submenu: Download Data Satker yang Belum Terdaftar di Tabel Referensi ---
        st.markdown("---")
        st.subheader("📥 Download Data Satker yang Belum Terdaftar di Tabel Referensi")

        # --- Helpers for GitHub fallback (minimal) ---
        GITHUB_REPO = st.secrets.get("GITHUB_REPO") if hasattr(st, "secrets") else None
        GITHUB_TOKEN = st.secrets.get("GITHUB_TOKEN") if hasattr(st, "secrets") else None

        def github_api_list_dir(repo, path, token=None):
            api = f"https://api.github.com/repos/{repo}/contents/{path}"
            headers = {"Authorization": f"token {token}"} if token else {}
            r = requests.get(api, headers=headers, timeout=30)
            r.raise_for_status()
            return r.json()

        def github_download_file_bytes(repo, path, token=None):
            api = f"https://api.github.com/repos/{repo}/contents/{path}"
            headers = {"Authorization": f"token {token}"} if token else {}
            r = requests.get(api, headers=headers, timeout=30)
            r.raise_for_status()
            j = r.json()
            if j.get("encoding") == "base64" and "content" in j:
                return base64.b64decode(j["content"])
            # fallback to raw URL
            raw = f"https://raw.githubusercontent.com/{repo}/HEAD/{path}"
            r2 = requests.get(raw, headers=headers, timeout=30)
            r2.raise_for_status()
            return r2.content

        def infer_tahun_bulan_from_filename(filename):
            # try to get year (20xx)
            year_match = re.search(r'(20\d{2})', filename)
            year = year_match.group(1) if year_match else ''
            # try month name
            found_month = ''
            for m in range(1, 13):
                short = calendar.month_name[m][:3].lower()
                if short in filename.lower():
                    found_month = calendar.month_name[m]
                    break
            return year, found_month

        # Action
        if st.button("📥 Generate & Download Rows dengan Uraian Satker-SINGKAT Kosong (Verbose)"):
            log_lines = []
            log_placeholder = st.empty()
            progress = st.progress(0)
            processed = 0
            found_list = []

            # Choose source: session state first
            data_storage = st.session_state.get('data_storage')
            use_session = bool(data_storage)

            if use_session:
                log_lines.append("📦 Menggunakan data dari st.session_state['data_storage']...")
                files_iter = []
                # data_storage keys expected as (month, year) -> df
                # convert to list of tuples (label, df)
                for key, df in data_storage.items():
                    # create a readable file label
                    if isinstance(key, (list, tuple)) and len(key) == 2:
                        month, year = key
                        label = f"{month}_{year}"
                    else:
                        label = str(key)
                    files_iter.append((label, df))
            else:
                # fallback to GitHub listing
                if not GITHUB_REPO:
                    st.error("❌ Tidak ada st.session_state['data_storage'] dan st.secrets['GITHUB_REPO'] tidak diset. Tidak bisa mengambil file.")
                    st.stop()
                try:
                    contents = github_api_list_dir(GITHUB_REPO, "data", GITHUB_TOKEN)
                except Exception as e:
                    st.error(f"❌ Gagal mengambil daftar file dari repo (folder `data`): {e}")
                    st.stop()
                # keep only xlsx/xls/csv
                file_items = [it for it in contents if it.get('type') == 'file' and it.get('name','').lower().endswith(('.xlsx','.xls','.csv'))]
                if not file_items:
                    st.warning("⚠️ Tidak ditemukan file .xlsx/.xls/.csv di folder `data` pada repo.")
                    st.stop()
                files_iter = []
                for it in file_items:
                    name = it.get('name')
                    path = it.get('path')
                    # we'll download bytes during processing
                    files_iter.append((name, path))

            total_files = len(files_iter)
            log_lines.append(f"🔎 Ditemukan {total_files} file untuk diproses.")
            log_placeholder.markdown("\n".join(log_lines))

            # iterate files
            for idx, (label, source) in enumerate(files_iter, start=1):
                processed += 1
                progress.progress(int((processed-1)/total_files * 100))

                # human-friendly filename for logs
                filename = label if isinstance(label, str) else str(label)
                log_lines.append(f"\n• Memproses file: **{filename}** ...")
                log_placeholder.markdown("\n".join(log_lines))

                # obtain dataframe
                df = None
                read_error = None
                if use_session:
                    df = source  # in session mode, source is already the DataFrame
                    # if the session stored some metadata with filename, optionally use that:
                    # but we proceed with label as filename.
                else:
                    # download bytes from repo path
                    try:
                        bytes_data = github_download_file_bytes(GITHUB_REPO, source, GITHUB_TOKEN)
                    except Exception as e:
                        read_error = f"Gagal mengunduh {filename}: {e}"
                        log_lines.append(f"  ❌ {read_error}")
                        log_placeholder.markdown("\n".join(log_lines))
                        continue

                    # read into dataframe (try .xlsx -> .csv)
                    try:
                        if filename.lower().endswith(('.xlsx', '.xls')):
                            df = pd.read_excel(io.BytesIO(bytes_data), dtype=str)
                        else:
                            # csv: try default then fallback to latin-1 if error
                            try:
                                df = pd.read_csv(io.BytesIO(bytes_data), dtype=str)
                            except Exception:
                                df = pd.read_csv(io.BytesIO(bytes_data), dtype=str, encoding='latin-1')
                    except Exception as e:
                        read_error = f"Gagal membaca {filename} ke DataFrame: {e}"
                        log_lines.append(f"  ❌ {read_error}")
                        log_placeholder.markdown("\n".join(log_lines))
                        continue

                # if not a DataFrame, skip
                if not isinstance(df, pd.DataFrame):
                    log_lines.append(f"  ❌ File {filename} tidak menghasilkan DataFrame. Dilewati.")
                    log_placeholder.markdown("\n".join(log_lines))
                    continue

                # show detected columns for debugging
                cols = df.columns.tolist()
                log_lines.append(f"  ℹ️ Kolom: {', '.join(cols)}")
                log_placeholder.markdown("\n".join(log_lines))

                # Ensure column exists: 'Uraian Satker-SINGKAT'
                uraian_col = None
                if 'Uraian Satker-SINGKAT' in df.columns:
                    uraian_col = 'Uraian Satker-SINGKAT'
                else:
                    # attempt to find best candidate
                    candidates = [c for c in df.columns if 'uraian' in c.lower() and ('satker' in c.lower() or 'singkat' in c.lower() or 'nama' in c.lower())]
                    if candidates:
                        uraian_col = candidates[0]
                        log_lines.append(f"  ℹ️ Menganggap kolom '{uraian_col}' sebagai 'Uraian Satker-SINGKAT'.")
                        log_placeholder.markdown("\n".join(log_lines))
                    else:
                        # no appropriate column — still continue (we can treat as all empty)
                        log_lines.append(f"  ⚠️ Tidak menemukan kolom 'Uraian Satker-SINGKAT' atau kandidat yang jelas pada {filename}. Menganggap semua uraian singkat sebagai kosong.")
                        log_placeholder.markdown("\n".join(log_lines))
                        uraian_col = None

                # Detect a full description 'Uraian Satker' column (prefer column with 'uraian' but not 'singkat')
                uraian_full_col = None
                full_candidates = [c for c in df.columns if 'uraian' in c.lower() and 'singkat' not in c.lower()]
                if full_candidates:
                    # pick first candidate that is not the same as uraian_col (if possible)
                    for c in full_candidates:
                        if uraian_col and c.lower() == uraian_col.lower():
                            continue
                        uraian_full_col = c
                        break
                    if uraian_full_col is None and full_candidates:
                        uraian_full_col = full_candidates[0]
                    if uraian_full_col:
                        log_lines.append(f"  ℹ️ Menganggap kolom '{uraian_full_col}' sebagai 'Uraian Satker' (deskripsi penuh).")
                        log_placeholder.markdown("\n".join(log_lines))
                else:
                    # no full uraian column found; we will create an empty 'Uraian Satker' later
                    log_lines.append(f"  ℹ️ Tidak menemukan kolom 'Uraian Satker' terpisah; akan membuat kolom kosong jika diperlukan.")
                    log_placeholder.markdown("\n".join(log_lines))

                # Prepare key columns: Kode Satker, Tahun, Bulan
                # Try common names for Kode Satker
                kode_candidates = [c for c in df.columns if 'kode' in c.lower() and 'satker' in c.lower()]
                if kode_candidates:
                    kode_col = kode_candidates[0]
                    if kode_col != 'Kode Satker':
                        log_lines.append(f"  ℹ️ Menganggap kolom '{kode_col}' sebagai 'Kode Satker'.")
                        log_placeholder.markdown("\n".join(log_lines))
                else:
                    # fallback to any column with numeric-like values and short names
                    kode_col = None
                    for c in df.columns:
                        if c.lower().strip() in ('kode','kd','kdsatker','kodesatker'):
                            kode_col = c
                            break
                    if kode_col is None:
                        # as last resort, look for first column that seems numeric-ish
                        for c in df.columns:
                            sample = df[c].dropna().astype(str).head(5).tolist()
                            if sample and all(re.fullmatch(r'\d+', s.strip()) for s in sample):
                                kode_col = c
                                break
                    if kode_col:
                        log_lines.append(f"  ℹ️ Menggunakan kolom '{kode_col}' sebagai 'Kode Satker' (fallback).")
                        log_placeholder.markdown("\n".join(log_lines))
                    else:
                        log_lines.append(f"  ⚠️ Tidak menemukan kolom 'Kode Satker' pada {filename}. File dilewati.")
                        log_placeholder.markdown("\n".join(log_lines))
                        continue

                # Normalize kode and uraian columns
                df_copy = df.copy()
                df_copy[kode_col] = df_copy[kode_col].astype(str).fillna('').str.strip()
                if uraian_col:
                    df_copy[uraian_col] = df_copy[uraian_col].astype(str).fillna('').str.strip()
                else:
                    # create empty column to mark empty uraian singkat
                    uraian_col = 'Uraian Satker-SINGKAT'
                    df_copy[uraian_col] = ''

                if uraian_full_col:
                    df_copy[uraian_full_col] = df_copy[uraian_full_col].astype(str).fillna('').str.strip()
                else:
                    # create empty full uraian column
                    uraian_full_col = 'Uraian Satker'
                    df_copy[uraian_full_col] = ''

                # Ensure Tahun/Bulan exist or infer from filename
                if 'Tahun' not in df_copy.columns:
                    year, _ = infer_tahun_bulan_from_filename(filename)
                    df_copy['Tahun'] = df_copy.get('Tahun', year)
                if 'Bulan' not in df_copy.columns:
                    _, month_name = infer_tahun_bulan_from_filename(filename)
                    df_copy['Bulan'] = df_copy.get('Bulan', month_name if month_name else '')

                # Identify rows where Uraian Satker-SINGKAT is empty (after strip)
                mask_empty_uraian = df_copy[uraian_col].astype(str).str.strip().replace('nan','').isin(['', None])
                empty_count = int(mask_empty_uraian.sum())

                if empty_count == 0:
                    log_lines.append(f"  ✅ Tidak ada baris dengan Uraian Satker-SINGKAT kosong pada {filename}.")
                    log_placeholder.markdown("\n".join(log_lines))
                else:
                    # select columns and attach source filename
                    # include both Uraian Satker-SINGKAT and Uraian Satker (full)
                    subset = df_copy.loc[mask_empty_uraian, ['Tahun', 'Bulan', kode_col, uraian_col, uraian_full_col]].copy()
                    subset = subset.rename(columns={
                        kode_col: 'Kode Satker',
                        uraian_col: 'Uraian Satker-SINGKAT',
                        uraian_full_col: 'Uraian Satker'
                    })
                    subset['Source File'] = filename
                    found_list.append(subset)
                    unique_codes = int(subset['Kode Satker'].nunique())
                    log_lines.append(f"  ⚠️ Menemukan **{empty_count}** baris dengan Uraian Satker-SINGKAT kosong pada {filename} (unik {unique_codes} kode).")
                    log_placeholder.markdown("\n".join(log_lines))

            # finalize
            progress.progress(100)
            log_lines.append("\n🔚 Pemeriksaan selesai.")
            log_placeholder.markdown("\n".join(log_lines))

            if not found_list:
                st.success("✅ Tidak ditemukan baris dengan Uraian Satker-SINGKAT kosong di semua file yang diproses.")
                st.session_state['last_empty_uraian_report'] = None
            else:
                all_empty = pd.concat(found_list, ignore_index=True).drop_duplicates()
                # ensure columns order
                cols_wanted = ['Tahun', 'Bulan', 'Kode Satker', 'Uraian Satker-SINGKAT', 'Uraian Satker', 'Source File']
                for c in cols_wanted:
                    if c not in all_empty.columns:
                        all_empty[c] = ''
                all_empty = all_empty[cols_wanted]

                total_rows = len(all_empty)
                total_unique = int(all_empty['Kode Satker'].nunique())
                st.warning(f"⚠️ Total {total_rows} baris ditemukan dengan Uraian Satker-SINGKAT kosong ({total_unique} satker unik).")

                # Build Excel to download
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    all_empty.to_excel(writer, sheet_name='Uraian_Kosong', index=False)
                output.seek(0)
                excel_bytes = output.getvalue()
                st.session_state['last_empty_uraian_report'] = excel_bytes

                fname = f"uraian_satker_singkat_kosong_{pd.Timestamp.now():%Y%m%d_%H%M%S}.xlsx"
                st.download_button(
                    label="📥 Download Laporan Uraian Satker-SINGKAT Kosong",
                    data=excel_bytes,
                    file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            # final flush of logs
            log_placeholder.markdown("\n".join(log_lines))

    # ============================================================
    # TAB 4: DOWNLOAD TEMPLATE (including Reference Template)
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
            template_ref.to_excel(writer, index=False, sheet_name='Data Referensi')
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
