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
    """Membersihkan format angka akuntansi menjadi float standar."""
    if pd.isna(value):
        return 0.0
    val = str(value).replace('(Dr)', '').replace('(Cr)', '').replace('(', '').replace(')', '').strip()
    val = val.replace('.', '').replace(',', '.')
    try:
        return float(val)
    except ValueError:
        return 0.0

def format_date(date_str):
    """Mengubah format tanggal menjadi DD/MM/YYYY."""
    if not isinstance(date_str, str):
        if pd.isna(date_str):
            return ""
        try:
            return date_str.strftime('%d/%m/%Y')
        except:
            return str(date_str)

    months = {
        'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04', 'Mei': '05', 'Jun': '06',
        'Jul': '07', 'Agu': '08', 'Sep': '09', 'Okt': '10', 'Nov': '11', 'Des': '12',
        'Agustus': '08', 'Januari': '01', 'Februari': '02', 'Maret': '03', 'April': '04',
        'Juni': '06', 'Juli': '07', 'September': '09', 'Oktober': '10', 'November': '11', 'Desember': '12'
    }
    
    try:
        parts = date_str.split()
        if len(parts) >= 3:
            day = parts[0].zfill(2)
            month_str = parts[1]
            year = parts[2]
            month = months.get(month_str, '01') 
            return f"{day}/{month}/{year}"
    except:
        return date_str
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
            return None
    except Exception as e:
        st.error(f"Gagal membaca file. Error: {e}")
        return pd.DataFrame()

    # --- LOGIKA DINAMIS PENCARIAN KOLOM ---
    # Kita cari baris yang mengandung kata kunci "Tanggal" dan "Debit" untuk jadi patokan header
    header_row_idx = None
    col_map = {}

    for idx, row in df_raw.iterrows():
        row_values = [str(x).lower() for x in row.values]
        if 'tanggal' in row_values and 'debit' in row_values:
            header_row_idx = idx
            # Mapping nama kolom ke index-nya
            for col_idx, val in enumerate(row_values):
                if 'tanggal' in val: col_map['date'] = col_idx
                if 'keterangan' in val: col_map['desc'] = col_idx
                if 'debit' in val: col_map['debit'] = col_idx
                if 'kredit' in val: col_map['credit'] = col_idx
                if 'balance' in val or 'saldo' in val: col_map['balance'] = col_idx
            break
    
    # Jika tidak ketemu header standar, fallback ke index manual (file lama)
    if header_row_idx is None:
        st.warning("Format header tidak terdeteksi otomatis. Menggunakan mode kompatibilitas (Format Lama).")
        col_map = {'date': 2, 'desc': 12, 'debit': 19, 'credit': 21, 'balance': 23}
    
    # Deteksi posisi kolom Nama Akun & Saldo Awal relatif terhadap kolom Debit
    # Biasanya Nama Akun ada jauh di kiri, dan Saldo Awal ada di kolom Balance
    # Kita pakai heuristik: Nama Akun biasanya di baris Header Akun, di kolom yang agak awal.
    
    processed_rows = []
    current_account_name = None
    current_account_type = None

    # Mulai iterasi data (bisa dari baris 0, karena kita pakai logika per-baris)
    for idx, row in df_raw.iterrows():
        # LOGIKA DETEKSI HEADER AKUN (Baris yang punya Kode Akun di kolom 1)
        # Ciri: Kolom 1 ada isi, Kolom 0 kosong.
        if pd.notna(row[1]) and pd.isna(row[0]):
            # Cari Nama Akun & Tipe Akun di baris ini
            # Nama akun biasanya string panjang non-angka.
            # Kita cari kolom yang isinya string di baris ini
            
            # Coba ambil Nama Akun dari kolom 6 (File Lama) atau kolom 8 (File Baru)
            # Kita cari nilai string pertama setelah kolom kode akun
            potential_names = []
            for c in range(2, 10): # Scan kolom 2 sampai 9
                val = row[c]
                if pd.notna(val) and not str(val).replace('.','').isdigit():
                     potential_names.append((c, val))
            
            if potential_names:
                # Ambil yang pertama ditemukan sebagai Nama Akun
                current_account_name = potential_names[0][1]
                
                # Coba cari Tipe Akun (biasanya 'Kas/Bank', 'Akun Piutang', dll)
                # Biasanya ada di sebelah kanan Nama Akun
                current_account_type = "Umum" # Default
                for c in range(potential_names[0][0] + 1, 20):
                    val = row[c]
                    if pd.notna(val) and isinstance(val, str) and len(val) > 3:
                        current_account_type = val
                        break
            
            # Ambil Saldo Awal
            # Biasanya ada di kolom yang sama dengan kolom Balance transaksi
            idx_balance = col_map.get('balance', 23) # Default 23 if not found map
            # Cek offset kolom balance jika file baru (index 35)
            # Jika mapping dinamis aktif, pakai col_map['balance']
            
            # Khusus Saldo Awal, kadang posisinya geser dikit dari kolom Balance transaksi
            # Kita coba cari angka di sekitar kolom balance
            opening_balance = 0
            if idx_balance < len(row):
                 opening_balance = row[idx_balance]
            
            # Jika kosong, coba cari mundur sedikit (kadang alignment beda)
            if pd.isna(opening_balance) or str(opening_balance).strip() == '':
                 if idx_balance-3 < len(row): # Cek kolom balance di file baru (index 20 vs 35)
                     # Di file baru: Saldo Awal di 20, Balance Transaksi di 35. Beda jauh.
                     # Kita cari angka pertama dari kanan di baris header akun
                     for c in range(len(row)-1, 10, -1):
                         val = row[c]
                         if pd.notna(val) and any(char.isdigit() for char in str(val)):
                             opening_balance = val
                             break

            processed_rows.append({
                "Tanggal": "01/01/2025", 
                "Nama Akun": current_account_name,
                "Tipe Akun": current_account_type,
                "Keterangan": "Saldo Awal",
                "Debit": 0.0,
                "Kredit": 0.0,
                "Saldo": clean_number(opening_balance)
            })
            
        # LOGIKA DETEKSI TRANSAKSI
        # Syarat: Kolom Tanggal ada isinya, dan bukan header "Tanggal"
        elif pd.notna(row[col_map.get('date', 2)]) and str(row[col_map.get('date', 2)]).strip() != "Tanggal" and current_account_name:
             processed_rows.append({
                "Tanggal": format_date(row[col_map.get('date', 2)]),
                "Nama Akun": current_account_name,
                "Tipe Akun": current_account_type,
                "Keterangan": row[col_map.get('desc', 12)],
                "Debit": clean_number(row[col_map.get('debit', 19)]),
                "Kredit": clean_number(row[col_map.get('credit', 21)]),
                "Saldo": clean_number(row[col_map.get('balance', 23)])
            })

    return pd.DataFrame(processed_rows)

# --- UI / UX SIDEBAR (Instruksi) ---

with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/2830/2830284.png", width=80)
    st.title("Panduan Penggunaan")
    
    st.info("**Langkah 1: Ekspor Data**")
    st.markdown("""
    Buka Accurate, lalu:
    1. Masuk ke **Laporan** > **Buku Besar**.
    2. Pilih periode yang diinginkan.
    3. Klik **Ekspor** > Pilih **Excel** atau **CSV**.
    """)
    
    st.info("**Langkah 2: Upload & Proses**")
    st.markdown("""
    1. Upload file hasil ekspor di halaman utama.
    2. Tunggu proses cleaning selesai.
    3. Cek ringkasan angka di dashboard.
    """)
    
    st.success("**Langkah 3: Download**")
    st.markdown("Klik tombol **Download Excel** untuk mendapatkan file tabel yang sudah rapi (Flat File).")
    
    st.divider()
    st.caption("Dikembangkan untuk KAP KBS Samarinda")
    st.caption("Oleh Viany Ramadhany")

# --- UI / UX MAIN PAGE ---

st.title("📊 General Ledger Cleaner Tool")
st.markdown("""
Aplikasi ini mengubah format laporan Buku Besar **Accurate** (Hierarki) menjadi format **Tabel Datar (Flat)** yang siap untuk Pivot Table atau analisis lanjutan di Excel.
""")

st.divider()

# File Uploader
col1, col2 = st.columns([2, 1])
with col1:
    uploaded_file = st.file_uploader("📂 Upload File Buku Besar (.csv, .xls, .xlsx)", type=["csv", "xls", "xlsx"])

if uploaded_file:
    with st.spinner('Sedang membersihkan dan merapikan data...'):
        df_result = parse_ledger(uploaded_file)
    
    if df_result is not None and not df_result.empty:
        
        # --- DASHBOARD RINGKASAN ---
        st.success("✅ Data berhasil diproses!")
        
        # Hitung Metrics
        total_debit = df_result['Debit'].sum()
        total_kredit = df_result['Kredit'].sum()
        total_rows = len(df_result)
        total_accounts = df_result['Nama Akun'].nunique()

        # Tampilkan Metrics
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Baris Data", f"{total_rows:,}")
        m2.metric("Jumlah Akun", total_accounts)
        m3.metric("Total Debit", f"Rp {total_debit:,.0f}")
        m4.metric("Total Kredit", f"Rp {total_kredit:,.0f}")

        st.divider()

        # --- TABS FOR INTERACTIVITY ---
        tab1, tab2 = st.tabs(["🔍 Preview & Filter Data", "📥 Download Data"])

        with tab1:
            st.subheader("Eksplorasi Data")
            
            # Interaktif Filter
            all_accounts = df_result['Nama Akun'].unique().tolist()
            selected_accounts = st.multiselect("Filter berdasarkan Nama Akun:", all_accounts, default=None)
            
            if selected_accounts:
                df_display = df_result[df_result['Nama Akun'].isin(selected_accounts)]
            else:
                df_display = df_result
            
            # Tampilkan Data Editor (Lebih interaktif daripada dataframe biasa)
            st.data_editor(
                df_display,
                column_config={
                    "Debit": st.column_config.NumberColumn(format="Rp %.2f"),
                    "Kredit": st.column_config.NumberColumn(format="Rp %.2f"),
                    "Saldo": st.column_config.NumberColumn(format="Rp %.2f"),
                },
                use_container_width=True,
                hide_index=True,
                height=500
            )

        with tab2:
            st.subheader("Download Hasil")
            st.write("Data yang didownload adalah data lengkap (tidak terpengaruh filter di atas).")
            
            # Proses Excel di Memory
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df_result.to_excel(writer, index=False, sheet_name='Data_Rapi')
                
                # Format kolom Excel agar cantik saat dibuka
                workbook  = writer.book
                worksheet = writer.sheets['Data_Rapi']
                money_fmt = workbook.add_format({'num_format': '#,##0.00'})
                date_fmt = workbook.add_format({'num_format': 'dd/mm/yyyy'})
                
                # Auto-adjust column width (simple estimation)
                worksheet.set_column('A:D', 20) # Tanggal - Tipe Akun
                worksheet.set_column('E:E', 50) # Keterangan lebar
                worksheet.set_column('F:H', 18, money_fmt) # Kolom Angka

            st.download_button(
                label="📥 Download File Excel (.xlsx)",
                data=buffer,
                file_name="GL_Cleaned_Data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary" # Tombol menonjol
            )

    else:
        st.warning("File kosong atau format tidak dikenali. Pastikan file berasal dari Accurate.")
else:
    # Tampilan awal jika belum ada file
    st.info("👋 Silakan upload file di atas untuk memulai.")