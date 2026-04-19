import streamlit as st
import pandas as pd
import io
import re

# Konfigurasi Halaman
st.set_page_config(
    page_title="General Ledger Cleaner",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- FUNGSI UTAMA (LOGIKA BISNIS) ---

def clean_number(value):
    """
    Membersihkan format angka akuntansi menjadi float standar.
    Otomatis mendeteksi format Indonesia (1.000,00) atau US (1,000.00).
    """
    if pd.isna(value):
        return 0.0

    # 1. Bersihkan teks (Hapus (Dr), (Cr), tanda kurung, spasi)
    val = str(value).replace('(Dr)', '').replace('(Cr)', '').replace('Dr', '').replace('Cr', '')
    val = val.replace('(', '').replace(')', '').strip()

    if not val:
        return 0.0

    # 2. Deteksi Format Berdasarkan Separator Terakhir
    last_comma = val.rfind(',')
    last_dot   = val.rfind('.')

    try:
        # KASUS A: Format Indonesia — desimal pakai koma (1.000,00 atau 100,00)
        if last_comma > last_dot:
            clean_val = val.replace('.', '').replace(',', '.')
            return float(clean_val)

        # KASUS B: Hanya ada titik, tanpa koma
        elif last_comma == -1 and last_dot != -1:
            # Heuristik: jika tepat 3 digit setelah titik terakhir → ribuan Indonesia
            # Contoh: "500.000" → 500000, "1.500.000" → 1500000
            # Jika bukan 3 digit (misal "500.75") → desimal US biasa
            after_dot = val[last_dot + 1:]
            dots_in_val = val.count('.')
            if len(after_dot) == 3 and after_dot.isdigit():
                # Pola ribuan Indonesia: hapus semua titik
                return float(val.replace('.', ''))
            else:
                return float(val)

        # KASUS C: Format US/Inggris — desimal pakai titik, ribuan pakai koma (1,000.00)
        else:
            clean_val = val.replace(',', '')
            return float(clean_val)

    except ValueError:
        return 0.0


def format_date(date_str):
    """Mengubah format tanggal menjadi DD/MM/YYYY."""
    if not isinstance(date_str, str):
        if pd.isna(date_str):
            return ""
        try:
            return date_str.strftime('%d/%m/%Y')
        except Exception:
            return str(date_str)

    months = {
        'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04',
        'Mei': '05', 'Jun': '06', 'Jul': '07', 'Agu': '08',
        'Sep': '09', 'Okt': '10', 'Nov': '11', 'Des': '12',
        'Agustus': '08', 'Januari': '01', 'Februari': '02',
        'Maret': '03', 'April': '04', 'Juni': '06', 'Juli': '07',
        'September': '09', 'Oktober': '10', 'November': '11', 'Desember': '12',
        # Bahasa Inggris
        'January': '01', 'February': '02', 'March': '03', 'May': '05',
        'June': '06', 'July': '07', 'August': '08', 'September': '09',
        'October': '10', 'November': '11', 'December': '12',
    }

    try:
        parts = date_str.split()
        if len(parts) >= 3:
            day   = parts[0].zfill(2)
            month = months.get(parts[1], '01')
            year  = parts[2]
            return f"{day}/{month}/{year}"
    except Exception:
        pass
    return date_str


@st.cache_data(show_spinner=False)
def parse_ledger(uploaded_file):
    filename = uploaded_file.name.lower()

    try:
        if filename.endswith('.csv'):
            df_raw = pd.read_csv(uploaded_file, header=None, dtype=str)
        elif filename.endswith('.xls'):
            df_raw = pd.read_excel(uploaded_file, header=None, dtype=str, engine='xlrd')
        elif filename.endswith('.xlsx'):
            df_raw = pd.read_excel(uploaded_file, header=None, dtype=str, engine='openpyxl')
        else:
            st.error("Format file tidak didukung. Gunakan .csv, .xls, atau .xlsx.")
            return pd.DataFrame()
    except Exception as e:
        st.error(f"Gagal membaca file. Error: {e}")
        return pd.DataFrame()

    # --- LOGIKA DINAMIS PENCARIAN KOLOM ---
    header_row_idx = None
    col_map = {}

    for idx, row in df_raw.iterrows():
        row_values = [str(x).lower() for x in row.values]
        if 'tanggal' in row_values and 'debit' in row_values:
            header_row_idx = idx
            for col_idx, val in enumerate(row_values):
                if 'tanggal' in val:
                    col_map['date'] = col_idx
                if 'no. sumber' in val or 'no sumber' in val:
                    col_map['source_no'] = col_idx
                if 'keterangan' in val:
                    col_map['desc'] = col_idx
                if 'debit' in val:
                    col_map['debit'] = col_idx
                if 'kredit' in val:
                    col_map['credit'] = col_idx
                if 'balance' in val or 'saldo' in val:
                    col_map['balance'] = col_idx
            break

    # Fallback jika header tidak terdeteksi (Mode Kompatibilitas)
    if header_row_idx is None:
        st.warning("Format header tidak terdeteksi otomatis. Menggunakan mode kompatibilitas (Format Lama).")
        col_map = {
            'date': 2, 'source_no': 8, 'desc': 12,
            'debit': 19, 'credit': 21, 'balance': 23
        }

    processed_rows = []
    current_account_name = None
    current_account_type = None

    for idx, row in df_raw.iterrows():
        # Lewati baris header dan baris sebelumnya
        if header_row_idx is not None and idx <= header_row_idx:
            continue

        # LOGIKA DETEKSI NAMA AKUN (HEADER AKUN)
        # Ciri: kolom 0 kosong tapi kolom 1 berisi nilai
        if pd.notna(row.iloc[1]) and pd.isna(row.iloc[0]):
            potential_names = []
            for c in range(2, min(10, len(row))):
                val = row.iloc[c]
                if pd.notna(val) and str(val).strip():
                    potential_names.append((c, str(val).strip()))

            if potential_names:
                current_account_name = potential_names[0][1]

                # Coba deteksi tipe akun di sebelah kanan nama akun
                current_account_type = "Umum"
                for c in range(potential_names[0][0] + 1, min(20, len(row))):
                    val = row.iloc[c]
                    if pd.notna(val) and str(val).strip():
                        current_account_type = str(val).strip()
                        break

            # --- MENANGANI SALDO AWAL ---
            idx_balance = col_map.get('balance', 23)
            opening_balance = 0

            # Coba ambil saldo dari kolom Balance
            if idx_balance < len(row):
                opening_balance = row.iloc[idx_balance]

            # Jika kolom balance kosong, cari angka di kolom paling kanan
            if pd.isna(opening_balance) or str(opening_balance).strip() == '':
                for c in range(len(row) - 1, 10, -1):
                    val = row.iloc[c]
                    if pd.notna(val) and str(val).strip():
                        opening_balance = val
                        break

            processed_rows.append({
                "Tanggal"   : "01/01/2025",
                "Nama Akun" : current_account_name,
                "Tipe Akun" : current_account_type,
                "No. Sumber": "-",
                "Keterangan": "Saldo Awal",
                "Debit"     : 0.0,
                "Kredit"    : 0.0,
                "Saldo"     : clean_number(opening_balance)
            })

        # LOGIKA DETEKSI TRANSAKSI (BARIS DATA)
        elif (
            pd.notna(row.iloc[col_map.get('date', 2)])
            and str(row.iloc[col_map.get('date', 2)]).strip() not in ("Tanggal", "")
            and current_account_name
        ):
            # Ambil nilai No. Sumber
            source_val = "-"
            idx_source = col_map.get('source_no')
            if idx_source is not None and idx_source < len(row):
                val = row.iloc[idx_source]
                if pd.notna(val):
                    source_str = str(val)
                    # Bersihkan format float string (misal "1234.0" → "1234")
                    source_val = source_str[:-2] if source_str.endswith('.0') else source_str

            processed_rows.append({
                "Tanggal"   : format_date(row.iloc[col_map.get('date', 2)]),
                "Nama Akun" : current_account_name,
                "Tipe Akun" : current_account_type,
                "No. Sumber": source_val,
                "Keterangan": row.iloc[col_map.get('desc', 12)] if col_map.get('desc', 12) < len(row) else "",
                "Debit"     : clean_number(row.iloc[col_map.get('debit', 19)]) if col_map.get('debit', 19) < len(row) else 0.0,
                "Kredit"    : clean_number(row.iloc[col_map.get('credit', 21)]) if col_map.get('credit', 21) < len(row) else 0.0,
                "Saldo"     : clean_number(row.iloc[col_map.get('balance', 23)]) if col_map.get('balance', 23) < len(row) else 0.0,
            })

    if not processed_rows:
        return pd.DataFrame()

    df = pd.DataFrame(processed_rows)
    df["Keterangan"] = df["Keterangan"].fillna("").astype(str)
    return df


def to_excel(df: pd.DataFrame) -> bytes:
    """Mengonversi DataFrame ke bytes Excel dengan format akuntansi."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='General Ledger')
        workbook  = writer.book
        worksheet = writer.sheets['General Ledger']

        # Format
        header_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#2E75B6', 'font_color': 'white',
            'border': 1, 'align': 'center', 'valign': 'vcenter'
        })
        money_fmt = workbook.add_format({'num_format': '#,##0.00'})

        # Lebar kolom
        worksheet.set_column('A:A', 12)        # Tanggal
        worksheet.set_column('B:B', 30)        # Nama Akun
        worksheet.set_column('C:C', 20)        # Tipe Akun
        worksheet.set_column('D:D', 15)        # No. Sumber
        worksheet.set_column('E:E', 50)        # Keterangan
        worksheet.set_column('F:H', 18, money_fmt)  # Debit, Kredit, Saldo

        # Tulis ulang header dengan format khusus
        for col_num, col_name in enumerate(df.columns):
            worksheet.write(0, col_num, col_name, header_fmt)

        # Freeze baris header
        worksheet.freeze_panes(1, 0)

    return output.getvalue()


# =============================================================================
# UI STREAMLIT
# =============================================================================

st.title("📊 Accurate General Ledger Cleaner")
st.caption("Upload file General Ledger dari Accurate / sistem akuntansi lainnya untuk dibersihkan dan dikonversi ke format standar.")

# --- SIDEBAR ---
with st.sidebar:
    st.header("⚙️ Pengaturan")
    tahun_saldo = st.text_input("Tahun Saldo Awal", value="2025",
                                 help="Tahun yang akan digunakan pada baris 'Saldo Awal'")
    st.divider()
    st.info("**Format yang Didukung:**\n- `.xlsx` / `.xls`\n- `.csv`")
    st.markdown("---")
    st.markdown("**Kolom Output:**")
    st.markdown("Tanggal · Nama Akun · Tipe Akun · No. Sumber · Keterangan · Debit · Kredit · Saldo")

# --- UPLOAD FILE ---
uploaded_file = st.file_uploader(
    "Upload File General Ledger",
    type=["xlsx", "xls", "csv"],
    help="File bisa berformat Excel atau CSV hasil export dari Accurate / sistem lain."
)

if uploaded_file:
    with st.spinner("Memproses file, harap tunggu..."):
        df = parse_ledger(uploaded_file)

    if df is None or df.empty:
        st.error("❌ Tidak ada data yang berhasil diproses. Periksa format file Anda.")
        st.stop()

    # Terapkan tahun saldo awal dari pengaturan sidebar
    if tahun_saldo.strip():
        df["Tanggal"] = df["Tanggal"].apply(
            lambda x: x.replace("2025", tahun_saldo.strip()) if x == "01/01/2025" else x
        )

    st.success(f"✅ Berhasil memproses **{len(df):,} baris** dari **{df['Nama Akun'].nunique()}** akun.")

    # --- STATISTIK RINGKAS ---
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Baris",    f"{len(df):,}")
    col2.metric("Jumlah Akun",    f"{df['Nama Akun'].nunique():,}")
    col3.metric("Total Debit",    f"Rp {df['Debit'].sum():,.2f}")
    col4.metric("Total Kredit",   f"Rp {df['Kredit'].sum():,.2f}")

    st.divider()

    # --- FILTER ---
    with st.expander("🔍 Filter Data", expanded=False):
        fc1, fc2 = st.columns(2)
        akun_list = ["Semua"] + sorted(df["Nama Akun"].dropna().unique().tolist())
        selected_akun = fc1.selectbox("Filter Nama Akun", akun_list)
        tipe_list = ["Semua"] + sorted(df["Tipe Akun"].dropna().unique().tolist())
        selected_tipe = fc2.selectbox("Filter Tipe Akun", tipe_list)

    df_filtered = df.copy()
    if selected_akun != "Semua":
        df_filtered = df_filtered[df_filtered["Nama Akun"] == selected_akun]
    if selected_tipe != "Semua":
        df_filtered = df_filtered[df_filtered["Tipe Akun"] == selected_tipe]

    # --- TABEL PREVIEW ---
    st.subheader(f"📋 Preview Data ({len(df_filtered):,} baris)")
    st.dataframe(
        df_filtered,
        use_container_width=True,
        height=450,
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

    # Excel
    with dc1:
        excel_bytes = to_excel(df_filtered)
        base_name   = uploaded_file.name.rsplit('.', 1)[0]
        st.download_button(
            label    = "📥 Download File Excel (.xlsx)",
            data     = excel_bytes,
            file_name= f"{base_name}_cleaned.xlsx",
            mime     = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    # CSV
    with dc2:
        csv_bytes = df_filtered.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label    = "📄 Download File CSV (.csv)",
            data     = csv_bytes,
            file_name= f"{base_name}_cleaned.csv",
            mime     = "text/csv",
            use_container_width=True
        )

else:
    st.info("👆 Upload file General Ledger di atas untuk memulai.")
    st.markdown("""
    #### Cara Penggunaan
    1. Export General Ledger dari Accurate ke format **Excel / CSV**
    2. Upload file tersebut menggunakan tombol di atas
    3. Sistem akan otomatis mendeteksi kolom dan memproses data
    4. Download hasil dalam format **Excel** atau **CSV**
    """)