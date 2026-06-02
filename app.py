import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime

# =============================================================================
# KONFIGURASI HALAMAN
# =============================================================================

st.set_page_config(
    page_title="GL Cleaner Pro",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =============================================================================
# KONSTANTA & KEYWORD MAP
# =============================================================================

# Peta keyword untuk setiap field kolom
# Urutan dalam list menentukan prioritas (lebih awal = lebih diprioritaskan)
KEYWORD_MAP = {
    'date'      : ['tanggal', 'tgl', 'date', 'posting date', 'trans date'],
    'source_no' : ['no. sumber', 'no sumber', 'nosumber', 'no. bukti',
                   'no bukti', 'source no', 'referensi', 'ref', 'voucher'],
    'desc'      : ['keterangan', 'uraian', 'narasi', 'deskripsi',
                   'description', 'memo', 'remark'],
    'debit'     : ['debit', 'dr', 'debet'],
    'credit'    : ['kredit', 'credit', 'cr'],
    'balance'   : ['saldo', 'balance', 'saldo akhir', 'running balance',
                   'saldo berjalan'],
}

# Skor minimum agar baris dianggap sebagai header tabel
MIN_HEADER_SCORE = 3

# Batas baris pencarian header (jangan scan seluruh file)
MAX_HEADER_SCAN_ROWS = 80

# Fallback hardcoded (format Accurate lama)
FALLBACK_COL_MAP = {
    'date': 2, 'source_no': 8, 'desc': 12,
    'debit': 19, 'credit': 21, 'balance': 23
}


# =============================================================================
# LAYER 1: DETEKSI KOLOM DINAMIS
# =============================================================================

def score_header_row(row_values: list) -> int:
    """
    Beri skor seberapa mungkin baris ini adalah header tabel GL.
    Semakin banyak keyword cocok, semakin tinggi skor.
    """
    score = 0
    all_keywords = [kw for keywords in KEYWORD_MAP.values() for kw in keywords]
    for cell in row_values:
        cell_clean = str(cell).lower().strip()
        if not cell_clean or cell_clean == 'nan':
            continue
        for kw in all_keywords:
            if kw in cell_clean:
                score += 1
                break  # satu cell maksimal 1 poin
    return score


def detect_columns(header_row: list) -> dict:
    """
    Petakan kolom berdasarkan keyword matching (partial/substring).
    Ambil kolom pertama yang cocok untuk setiap field.

    Args:
        header_row: list string dari satu baris header (belum diubah case-nya)

    Returns:
        dict col_map, contoh: {'date': 2, 'debit': 5, ...}
    """
    col_map = {}
    for col_idx, cell in enumerate(header_row):
        cell_clean = str(cell).lower().strip()
        if not cell_clean or cell_clean == 'nan':
            continue
        for field, keywords in KEYWORD_MAP.items():
            if field not in col_map:  # ambil match pertama saja
                if any(kw in cell_clean for kw in keywords):
                    col_map[field] = col_idx
    return col_map


def find_header_row(df_raw: pd.DataFrame) -> tuple[int | None, dict, str]:
    """
    Cari baris header terbaik menggunakan sistem skor.
    Scan maksimal MAX_HEADER_SCAN_ROWS baris pertama.

    Returns:
        (header_row_idx, col_map, detection_method)
        detection_method: 'dynamic' | 'fallback'
    """
    best_row_idx   = None
    best_row_score = 0

    scan_limit = min(MAX_HEADER_SCAN_ROWS, len(df_raw))

    for idx in range(scan_limit):
        row        = df_raw.iloc[idx]
        row_str    = [str(x) for x in row.values]
        score      = score_header_row(row_str)

        if score > best_row_score:
            best_row_score = score
            best_row_idx   = idx

    if best_row_score >= MIN_HEADER_SCORE and best_row_idx is not None:
        header_row   = df_raw.iloc[best_row_idx].tolist()
        col_map      = detect_columns(header_row)
        return best_row_idx, col_map, 'dynamic'
    else:
        return None, FALLBACK_COL_MAP.copy(), 'fallback'


# =============================================================================
# LAYER 2: TRANSFORMASI DATA
# =============================================================================

def clean_number(value) -> float:
    """
    Bersihkan format angka akuntansi → float standar.
    Mendukung:
      - Format Indonesia : 1.500.000,75
      - Format US        : 1,500,000.75
      - Ribuan tanpa desimal: 500.000 → 500000
      - Tanda akuntansi  : (Dr), (Cr), tanda kurung negatif
    """
    if pd.isna(value):
        return 0.0

    val = str(value).strip()

    # Hapus label akuntansi
    for token in ['(Dr)', '(Cr)', 'Dr', 'Cr']:
        val = val.replace(token, '')

    # Tangani tanda kurung sebagai negatif: (500) → -500
    is_negative = val.startswith('(') and val.endswith(')')
    val = val.replace('(', '').replace(')', '').strip()

    if not val:
        return 0.0

    last_comma = val.rfind(',')
    last_dot   = val.rfind('.')

    try:
        if last_comma > last_dot:
            # Format Indonesia: 1.500.000,75
            result = float(val.replace('.', '').replace(',', '.'))

        elif last_comma == -1 and last_dot != -1:
            after_dot  = val[last_dot + 1:]
            dots_count = val.count('.')
            # Heuristik ribuan Indonesia: 500.000 atau 1.500.000
            if len(after_dot) == 3 and after_dot.isdigit() and dots_count >= 1:
                result = float(val.replace('.', ''))
            else:
                result = float(val)

        else:
            # Format US: 1,500,000.75 atau tanpa separator
            result = float(val.replace(',', ''))

    except ValueError:
        return 0.0

    return -result if is_negative else result


def format_date(date_str) -> str:
    """
    Normalisasi berbagai format tanggal → DD/MM/YYYY.
    Mendukung:
      - DD/MM/YYYY, DD-MM-YYYY, DD.MM.YYYY
      - YYYY-MM-DD (ISO)
      - 1 Jan 2025, 01 Januari 2025, 10-nop-2025
      - Excel date serial numbers
      - pandas.Timestamp
    """
    if pd.isna(date_str):
        return ""

    if isinstance(date_str, pd.Timestamp):
        return date_str.strftime('%d/%m/%Y')

    if not isinstance(date_str, str):
        try:
            return pd.to_datetime(date_str).strftime('%d/%m/%Y')
        except Exception:
            return str(date_str)

    date_str = date_str.strip()
    if not date_str:
        return ""

    MONTHS = {
        # Inggris singkat
        'jan': '01', 'feb': '02', 'mar': '03', 'apr': '04',
        'may': '05', 'jun': '06', 'jul': '07', 'aug': '08',
        'sep': '09', 'oct': '10', 'nov': '11', 'dec': '12',
        # Indonesia singkat
        'mei': '05', 'agu': '08', 'okt': '10', 'nop': '11', 'des': '12',
        # Indonesia panjang
        'januari': '01', 'februari': '02', 'maret': '03', 'april': '04',
        'juni': '06', 'juli': '07', 'agustus': '08',
        'september': '09', 'oktober': '10', 'november': '11', 'desember': '12',
        # Inggris panjang
        'january': '01', 'february': '02', 'march': '03',
        'june': '06', 'july': '07', 'august': '08',
        'october': '10', 'december': '12',
    }

    # FORMAT 1: Teks bulan — "1 Jan 2025", "10-nop-2025", "01 Januari 2025"
    match = re.search(r'(\d{1,2})[\s\-\/\.]+([a-zA-Z]+)[\s\-\/\.]+(\d{4})', date_str)
    if match:
        day   = match.group(1).zfill(2)
        month = MONTHS.get(match.group(2).lower(), '')
        year  = match.group(3)
        if month:
            return f"{day}/{month}/{year}"

    # FORMAT 2: DD-MM-YYYY atau DD/MM/YYYY
    match = re.match(r'^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$', date_str)
    if match:
        return f"{match.group(1).zfill(2)}/{match.group(2).zfill(2)}/{match.group(3)}"

    # FORMAT 3: YYYY-MM-DD (ISO)
    match = re.match(r'^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$', date_str)
    if match:
        return f"{match.group(3).zfill(2)}/{match.group(2).zfill(2)}/{match.group(1)}"

    # FORMAT 4: DD.MM.YYYY (Eropa)
    match = re.match(r'^(\d{1,2})\.(\d{1,2})\.(\d{4})$', date_str)
    if match:
        return f"{match.group(1).zfill(2)}/{match.group(2).zfill(2)}/{match.group(3)}"

    # FALLBACK: pandas auto-detect
    try:
        return pd.to_datetime(date_str).strftime('%d/%m/%Y')
    except Exception:
        return date_str


def get_safe_cell(row, col_idx, default=""):
    """Ambil nilai cell dengan aman (guard index out of bounds)."""
    try:
        if col_idx is None or col_idx >= len(row):
            return default
        val = row.iloc[col_idx]
        return val if pd.notna(val) else default
    except Exception:
        return default


def clean_source_no(val) -> str:
    """Bersihkan nomor sumber dari trailing .0 (artefak Excel)."""
    if pd.isna(val):
        return "-"
    s = str(val).strip()
    return s[:-2] if s.endswith('.0') else s


# =============================================================================
# LAYER 3: PARSER UTAMA
# =============================================================================

@st.cache_data(show_spinner=False)
def parse_ledger(uploaded_file) -> tuple[pd.DataFrame, dict]:
    """
    Parse file GL dari berbagai format (xlsx, xls, csv).

    Returns:
        (df_result, meta)
        meta berisi info deteksi: header_row_idx, col_map, detection_method
    """
    filename = uploaded_file.name.lower()
    meta     = {}

    # --- BACA FILE ---
    try:
        if filename.endswith('.csv'):
            df_raw = pd.read_csv(uploaded_file, header=None, dtype=str)
        elif filename.endswith('.xls'):
            df_raw = pd.read_excel(uploaded_file, header=None, dtype=str, engine='xlrd')
        elif filename.endswith('.xlsx'):
            df_raw = pd.read_excel(uploaded_file, header=None, dtype=str, engine='openpyxl')
        else:
            st.error("Format tidak didukung. Gunakan .csv, .xls, atau .xlsx.")
            return pd.DataFrame(), meta
    except Exception as e:
        st.error(f"Gagal membaca file: {e}")
        return pd.DataFrame(), meta

    # --- DETEKSI HEADER (LAYER 1) ---
    header_row_idx, col_map, detection_method = find_header_row(df_raw)

    meta['header_row_idx']   = header_row_idx
    meta['col_map']          = col_map
    meta['detection_method'] = detection_method
    meta['total_raw_rows']   = len(df_raw)

    # --- PROSES BARIS ---
    processed_rows          = []
    current_account_name    = None
    current_account_type    = None

    start_idx = (header_row_idx + 1) if header_row_idx is not None else 0

    for idx in range(start_idx, len(df_raw)):
        row = df_raw.iloc[idx]

        # ── DETEKSI HEADER AKUN ──────────────────────────────────────────────
        # Ciri: kolom 0 kosong, kolom 1 ada isi (nama akun di Accurate)
        col0_empty = pd.isna(row.iloc[0]) or str(row.iloc[0]).strip() == ''
        col1_filled = len(row) > 1 and pd.notna(row.iloc[1]) and str(row.iloc[1]).strip() != ''

        if col0_empty and col1_filled:
            # Cari nama akun: ambil nilai non-kosong pertama mulai kolom 2
            account_name = None
            account_type = "Umum"
            name_col_idx = None

            for c in range(2, min(12, len(row))):
                val = str(row.iloc[c]).strip()
                if val and val.lower() != 'nan':
                    account_name = val
                    name_col_idx = c
                    break

            if not account_name:
                continue

            # Cari tipe akun di sebelah kanan nama akun
            if name_col_idx is not None:
                for c in range(name_col_idx + 1, min(25, len(row))):
                    val = str(row.iloc[c]).strip()
                    if val and val.lower() != 'nan':
                        account_type = val
                        break

            current_account_name = account_name
            current_account_type = account_type

            # ── SALDO AWAL ───────────────────────────────────────────────────
            date_raw      = get_safe_cell(row, col_map.get('date'), None)
            opening_date  = format_date(date_raw) if date_raw else "01/01/2025"

            # Coba ambil saldo dari kolom balance, fallback ke kolom paling kanan
            bal_raw = get_safe_cell(row, col_map.get('balance'), None)
            if bal_raw is None or str(bal_raw).strip() in ('', 'nan'):
                for c in range(len(row) - 1, 10, -1):
                    val = str(row.iloc[c]).strip()
                    if val and val.lower() != 'nan':
                        bal_raw = val
                        break

            opening_balance = clean_number(bal_raw)

            processed_rows.append({
                "Tanggal"   : opening_date,
                "Nama Akun" : current_account_name,
                "Tipe Akun" : current_account_type,
                "No. Sumber": "-",
                "Keterangan": "Saldo Awal",
                "Debit"     : 0.0,
                "Kredit"    : 0.0,
                "Saldo"     : opening_balance,
            })

        # ── DETEKSI BARIS TRANSAKSI ──────────────────────────────────────────
        elif current_account_name:
            date_val = get_safe_cell(row, col_map.get('date'), "")
            date_str = str(date_val).strip()

            # Skip baris kosong atau baris summary (tidak ada tanggal valid)
            if not date_str or date_str.lower() in ('nan', 'tanggal', 'date', ''):
                continue

            processed_rows.append({
                "Tanggal"   : format_date(date_val),
                "Nama Akun" : current_account_name,
                "Tipe Akun" : current_account_type,
                "No. Sumber": clean_source_no(get_safe_cell(row, col_map.get('source_no'), "-")),
                "Keterangan": str(get_safe_cell(row, col_map.get('desc'), "")),
                "Debit"     : clean_number(get_safe_cell(row, col_map.get('debit'), 0.0)),
                "Kredit"    : clean_number(get_safe_cell(row, col_map.get('credit'), 0.0)),
                "Saldo"     : clean_number(get_safe_cell(row, col_map.get('balance'), 0.0)),
            })

    if not processed_rows:
        return pd.DataFrame(), meta

    df = pd.DataFrame(processed_rows)
    df["Keterangan"] = df["Keterangan"].fillna("").astype(str)
    return df, meta


# =============================================================================
# LAYER 4: OUTPUT
# =============================================================================

def to_excel(df: pd.DataFrame) -> bytes:
    """Ekspor DataFrame ke bytes Excel dengan format akuntansi profesional."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='General Ledger')
        wb  = writer.book
        ws  = writer.sheets['General Ledger']

        # Format
        fmt_header = wb.add_format({
            'bold': True, 'bg_color': '#2E75B6', 'font_color': 'white',
            'border': 1, 'align': 'center', 'valign': 'vcenter',
            'text_wrap': True,
        })
        fmt_money  = wb.add_format({'num_format': '#,##0.00'})
        fmt_date   = wb.add_format({'align': 'center'})
        fmt_center = wb.add_format({'align': 'center'})
        fmt_alt    = wb.add_format({'bg_color': '#EBF3FB'})
        fmt_alt_money = wb.add_format({'bg_color': '#EBF3FB', 'num_format': '#,##0.00'})

        # Lebar kolom
        ws.set_column('A:A', 13, fmt_date)     # Tanggal
        ws.set_column('B:B', 35)               # Nama Akun
        ws.set_column('C:C', 22)               # Tipe Akun
        ws.set_column('D:D', 16, fmt_center)   # No. Sumber
        ws.set_column('E:E', 55)               # Keterangan
        ws.set_column('F:H', 20, fmt_money)    # Debit, Kredit, Saldo

        # Tulis ulang header dengan format
        for col_num, col_name in enumerate(df.columns):
            ws.write(0, col_num, col_name, fmt_header)
            ws.set_row(0, 22)

        # Alternating row color
        money_cols = {
            df.columns.get_loc('Debit'),
            df.columns.get_loc('Kredit'),
            df.columns.get_loc('Saldo'),
        }
        for row_num in range(1, len(df) + 1):
            is_alt = row_num % 2 == 0
            for col_num in range(len(df.columns)):
                val = df.iloc[row_num - 1, col_num]
                if col_num in money_cols:
                    ws.write_number(row_num, col_num, float(val) if val else 0.0,
                                    fmt_alt_money if is_alt else fmt_money)
                else:
                    ws.write(row_num, col_num, val,
                             fmt_alt if is_alt else None)

        ws.freeze_panes(1, 0)
        ws.autofilter(0, 0, len(df), len(df.columns) - 1)

    return output.getvalue()


# =============================================================================
# UI STREAMLIT
# =============================================================================

# --- HEADER ---
st.title("📊 GL Cleaner Pro")
st.caption("Bersihkan dan standarisasi General Ledger dari Accurate / software akuntansi lainnya secara otomatis.")

# --- SIDEBAR ---
with st.sidebar:
    st.header("⚙️ Pengaturan")

    fallback_year = st.text_input(
        "Tahun Fallback Saldo Awal", value="2025",
        help="Digunakan jika tanggal saldo awal tidak terdeteksi di file"
    )

    st.divider()
    st.markdown("**Format File yang Didukung**")
    st.info("`.xlsx` · `.xls` · `.csv`")

    st.divider()
    st.markdown("**Kolom Output**")
    st.markdown("""
    - Tanggal  
    - Nama Akun  
    - Tipe Akun  
    - No. Sumber  
    - Keterangan  
    - Debit  
    - Kredit  
    - Saldo  
    """)

    st.divider()
    st.markdown("**Tentang Deteksi Kolom**")
    st.markdown("""
    Sistem secara otomatis mencari baris header di 80 baris pertama 
    menggunakan sistem **skor keyword**. Semakin banyak kata kunci 
    akuntansi yang cocok, baris itu dipilih sebagai header.
    """)

# --- UPLOAD ---
uploaded_file = st.file_uploader(
    "Upload File General Ledger",
    type=["xlsx", "xls", "csv"],
    help="File Excel atau CSV hasil export dari Accurate / sistem lain."
)

if uploaded_file:
    with st.spinner("Memproses dan mendeteksi struktur file..."):
        result = parse_ledger(uploaded_file)

    # Handle return value (df, meta)
    if isinstance(result, tuple):
        df, meta = result
    else:
        df, meta = result, {}

    if df is None or df.empty:
        st.error("❌ Tidak ada data yang berhasil diproses. Periksa format file Anda.")
        st.stop()

    # --- INFO DETEKSI ---
    detection_method = meta.get('detection_method', 'unknown')
    col_map          = meta.get('col_map', {})
    header_row_idx   = meta.get('header_row_idx')

    if detection_method == 'dynamic':
        st.success(
            f"✅ Header terdeteksi otomatis di **baris {header_row_idx + 1}** "
            f"dengan **{len(col_map)} kolom** dipetakan."
        )
    else:
        st.warning(
            "⚠️ Header tidak terdeteksi otomatis. "
            "Menggunakan **mode kompatibilitas** (format kolom Accurate lama)."
        )

    # Detail mapping kolom (expandable)
    with st.expander("🔍 Detail Pemetaan Kolom yang Terdeteksi", expanded=False):
        if col_map:
            mapping_data = [
                {"Field": field, "Index Kolom": idx,
                 "Nama Kolom di File": f"Kolom {idx}"}
                for field, idx in col_map.items()
            ]
            st.dataframe(pd.DataFrame(mapping_data), use_container_width=True, hide_index=True)
        else:
            st.info("Tidak ada info pemetaan tersedia.")

    st.divider()

    # --- METRIK RINGKAS ---
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Baris",  f"{len(df):,}")
    c2.metric("Jumlah Akun",  f"{df['Nama Akun'].nunique():,}")
    c3.metric("Total Debit",  f"Rp {df['Debit'].sum():,.0f}")
    c4.metric("Total Kredit", f"Rp {df['Kredit'].sum():,.0f}")

    st.divider()

    # --- FILTER ---
    with st.expander("🔍 Filter Data", expanded=False):
        fc1, fc2, fc3 = st.columns(3)

        akun_list     = ["Semua"] + sorted(df["Nama Akun"].dropna().unique().tolist())
        selected_akun = fc1.selectbox("Nama Akun", akun_list)

        tipe_list     = ["Semua"] + sorted(df["Tipe Akun"].dropna().unique().tolist())
        selected_tipe = fc2.selectbox("Tipe Akun", tipe_list)

        ket_filter    = fc3.text_input("Keterangan mengandung...")

    df_filtered = df.copy()
    if selected_akun != "Semua":
        df_filtered = df_filtered[df_filtered["Nama Akun"] == selected_akun]
    if selected_tipe != "Semua":
        df_filtered = df_filtered[df_filtered["Tipe Akun"] == selected_tipe]
    if ket_filter.strip():
        df_filtered = df_filtered[
            df_filtered["Keterangan"].str.contains(ket_filter.strip(), case=False, na=False)
        ]

    # --- TABEL PREVIEW ---
    st.subheader(f"📋 Preview Data  —  {len(df_filtered):,} baris ditampilkan")
    st.dataframe(
        df_filtered,
        use_container_width=True,
        height=460,
        column_config={
            "Debit" : st.column_config.NumberColumn("Debit",  format="Rp %.2f"),
            "Kredit": st.column_config.NumberColumn("Kredit", format="Rp %.2f"),
            "Saldo" : st.column_config.NumberColumn("Saldo",  format="Rp %.2f"),
        }
    )

    st.divider()

    # --- DOWNLOAD ---
    st.subheader("💾 Download Hasil")
    dc1, dc2 = st.columns(2)
    base_name = uploaded_file.name.rsplit('.', 1)[0]

    with dc1:
        excel_bytes = to_excel(df_filtered)
        st.download_button(
            label           = "📥 Download Excel (.xlsx)",
            data            = excel_bytes,
            file_name       = f"{base_name}_cleaned.xlsx",
            mime            = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with dc2:
        csv_bytes = df_filtered.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label           = "📄 Download CSV (.csv)",
            data            = csv_bytes,
            file_name       = f"{base_name}_cleaned.csv",
            mime            = "text/csv",
            use_container_width=True,
        )

else:
    st.info("👆 Upload file General Ledger di atas untuk memulai.")

    st.markdown("""
    #### Cara Penggunaan
    1. Export General Ledger dari Accurate ke format **Excel** atau **CSV**
    2. Upload file menggunakan tombol di atas
    3. Sistem otomatis mendeteksi kolom dan memproses data
    4. Download hasilnya dalam format **Excel** atau **CSV**

    #### Format Tanggal yang Didukung
    | Format | Contoh |
    |---|---|
    | DD/MM/YYYY | `25/12/2024` |
    | DD-MM-YYYY | `25-12-2024` |
    | YYYY-MM-DD | `2024-12-25` |
    | DD.MM.YYYY | `25.12.2024` |
    | Bulan teks pendek | `25 Des 2024`, `25-nop-2024` |
    | Bulan teks panjang | `25 Desember 2024` |
    | Excel serial number | otomatis |
    """)