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
    # Cari posisi terakhir dari titik (.) dan koma (,)
    last_comma = val.rfind(',')
    last_dot = val.rfind('.')
    
    try:
        # KASUS A: Format Indonesia (Desimal pakai Koma)
        # Ciri: Koma muncul SETELAH titik (1.000,00) ATAU hanya ada koma (100,00)
        if last_comma > last_dot:
            # Hapus titik (ribuan), ganti koma jadi titik (desimal)
            clean_val = val.replace('.', '').replace(',', '.')
            return float(clean_val)
            
        # KASUS B: Format US/Inggris (Desimal pakai Titik)
        # Ciri: Titik muncul SETELAH koma (1,000.00) ATAU hanya ada titik (1000.0)
        else:
            # Hapus koma (ribuan), biarkan titik (desimal)
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
    header_row_idx = None
    col_map = {}

    for idx, row in df_raw.iterrows():
        row_values = [str(x).lower() for x in row.values]
        if 'tanggal' in row_values and 'debit' in row_values:
            header_row_idx = idx
            for col_idx, val in enumerate(row_values):
                if 'tanggal' in val: col_map['date'] = col_idx
                if 'keterangan' in val: col_map['desc'] = col_idx
                if 'debit' in val: col_map['debit'] = col_idx
                if 'kredit' in val: col_map['credit'] = col_idx
                if 'balance' in val or 'saldo' in val: col_map['balance'] = col_idx
            break
    
    if header_row_idx is None:
        st.warning("Format header tidak terdeteksi otomatis. Menggunakan mode kompatibilitas (Format Lama).")
        col_map = {'date': 2, 'desc': 12, 'debit': 19, 'credit': 21, 'balance': 23}
    
    processed_rows = []
    current_account_name = None
    current_account_type = None

    for idx, row in df_raw.iterrows():
        if pd.notna(row[1]) and pd.isna(row[0]):
            potential_names = []
            for c in range(2, 10):
                val = row[c]
                if pd.notna(val) and not str(val).replace('.','').isdigit():
                     potential_names.append((c, val))
            
            if potential_names:
                current_account_name = potential_names[0][1]
                
                current_account_type = "Umum"
                for c in range(potential_names[0][0] + 1, 20):
                    val = row[c]
                    if pd.notna(val) and isinstance(val, str) and len(val) > 3:
                        current_account_type = val
                        break
            
            idx_balance = col_map.get('balance', 23)
            opening_balance = 0
            if idx_balance < len(row):
                 opening_balance = row[idx_balance]
            
            if pd.isna(opening_balance) or str(opening_balance).strip() == '':
                 if idx_balance-3 < len(row): 
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
        
        # --- DASHBOARD RINGKASAN UTAMA ---
        st.success("✅ Data berhasil diproses!")
        
        # Hitung Metrics Dasar
        total_rows = len(df_result)
        total_accounts = df_result['Nama Akun'].nunique()

        # Tampilkan Metrics Global
        m1, m2 = st.columns(2)
        m1.metric("Total Baris Data", f"{total_rows:,}")
        m2.metric("Jumlah Akun", total_accounts)

        st.divider()

        # --- TABS FOR INTERACTIVITY ---
        tab1, tab2 = st.tabs(["🔍 Preview & Filter Data", "📥 Download Data"])

        with tab1:
            st.subheader("Eksplorasi Data")
            
            # Interaktif Filter
            all_accounts = df_result['Nama Akun'].unique().tolist()
            selected_accounts = st.multiselect("Filter berdasarkan Nama Akun (Pilih 1 untuk lihat Rincian):", all_accounts, default=None)
            
            # Logic Filter
            if selected_accounts:
                df_display = df_result[df_result['Nama Akun'].isin(selected_accounts)]
                
                # --- FITUR KHUSUS: SALDO HANYA MUNCUL JIKA PILIH 1 AKUN ---
                if len(selected_accounts) == 1:
                    # Ambil nama akun
                    acc_name = selected_accounts[0]
                    
                    # 1. Hitung Saldo Awal (Khusus akun ini)
                    saldo_awal_val = df_display[df_display['Keterangan'] == 'Saldo Awal']['Saldo'].sum()
                    
                    # 2. Hitung Saldo Akhir (Angka terakhir di kolom saldo)
                    if not df_display.empty:
                        saldo_akhir_val = df_display.iloc[-1]['Saldo']
                    else:
                        saldo_akhir_val = 0
                    
                    # 3. Hitung Jumlah Transaksi (Total baris termasuk saldo awal)
                    jumlah_transaksi = len(df_display)

                    # Tampilkan Metric Khusus Akun (3 Kolom)
                    st.markdown(f"**Rincian Akun: {acc_name}**")
                    c1, c2, c3 = st.columns(3)
                    
                    c1.metric("Saldo Awal", f"Rp {saldo_awal_val:,.2f}")
                    c2.metric("Saldo Akhir", f"Rp {saldo_akhir_val:,.2f}")
                    c3.metric("Jml. Transaksi", f"{jumlah_transaksi} Baris")
                    
                    st.divider()

            else:
                df_display = df_result
            
            # Tampilkan Data Editor
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
                
                # Format kolom Excel
                workbook  = writer.book
                worksheet = writer.sheets['Data_Rapi']
                money_fmt = workbook.add_format({'num_format': '#,##0.00'})
                date_fmt = workbook.add_format({'num_format': 'dd/mm/yyyy'})
                
                worksheet.set_column('A:D', 20) 
                worksheet.set_column('E:E', 50) 
                worksheet.set_column('F:H', 18, money_fmt) 

            st.download_button(
                label="📥 Download File Excel (.xlsx)",
                data=buffer,
                file_name="GL_Cleaned_Data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

    else:
        st.warning("File kosong atau format tidak dikenali. Pastikan file berasal dari Accurate.")
else:
    # Tampilan awal jika belum ada file
    st.info("👋 Silakan upload file di atas untuk memulai.")