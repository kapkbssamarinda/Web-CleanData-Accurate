import streamlit as st
import pandas as pd
import io
import re

# 1. Fungsi Pembersih Angka (Debit/Kredit/Saldo)
def clean_number(value):
    if pd.isna(value):
        return 0.0
    # Ubah ke string, hapus (Dr), (Cr), kurung, dan spasi
    val = str(value).replace('(Dr)', '').replace('(Cr)', '').replace('(', '').replace(')', '').strip()
    
    # Menangani format Indonesia (1.000.000,00)
    # Hapus titik ribuan, ganti koma desimal jadi titik
    val = val.replace('.', '').replace(',', '.')
    
    try:
        return float(val)
    except ValueError:
        return 0.0

# 2. Fungsi Format Tanggal (Menjadi DD/MM/YYYY)
def format_date(date_str):
    # Jika input bukan string (misal sudah datetime dari Excel), ubah dulu
    if not isinstance(date_str, str):
        if pd.isna(date_str):
            return ""
        try:
            # Jika Excel otomatis membaca sebagai datetime, format ulang
            return date_str.strftime('%d/%m/%Y')
        except:
            return str(date_str)

    # Mapping Bulan Indonesia ke Angka
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

# 3. Fungsi Utama Parsing (Support CSV & Excel)
def parse_ledger(uploaded_file):
    filename = uploaded_file.name.lower()
    
    try:
        # Deteksi jenis file dan baca sesuai formatnya
        if filename.endswith('.csv'):
            df_raw = pd.read_csv(uploaded_file, header=None, dtype=str)
        elif filename.endswith('.xls'):
            # Engine xlrd untuk file Excel lama (97-2003)
            df_raw = pd.read_excel(uploaded_file, header=None, dtype=str, engine='xlrd')
        elif filename.endswith('.xlsx'):
            # Engine openpyxl untuk file Excel baru
            df_raw = pd.read_excel(uploaded_file, header=None, dtype=str, engine='openpyxl')
        else:
            return None
    except Exception as e:
        st.error(f"Gagal membaca file. Pastikan format valid. Error: {e}")
        return pd.DataFrame()

    processed_rows = []
    current_account_name = None
    current_account_type = None

    # Iterasi data
    for idx, row in df_raw.iterrows():
        # A. Deteksi Header Akun
        # Kolom 1 ada isi, Kolom 0 kosong, Kolom 6 ada Nama Akun
        if pd.notna(row[1]) and pd.isna(row[0]) and pd.notna(row[6]):
            current_account_name = row[6]
            current_account_type = row[10] # Ambil Tipe Akun
            opening_balance = row[14]      # Ambil Saldo Awal
            
            # Buat Baris Saldo Awal (Debit=0, Kredit=0)
            processed_rows.append({
                "Tanggal": "01/01/2025", # Default start date
                "Nama Akun": current_account_name,
                "Tipe Akun": current_account_type,
                "Keterangan": "Saldo Awal",
                "Debit": 0.0,
                "Kredit": 0.0,
                "Saldo": clean_number(opening_balance)
            })
            
        # B. Deteksi Transaksi
        # Kolom 2 ada isi (Tanggal) dan bukan tulisan 'Tanggal'
        elif pd.notna(row[2]) and str(row[2]).strip() != "Tanggal" and current_account_name:
            processed_rows.append({
                "Tanggal": format_date(row[2]),
                "Nama Akun": current_account_name,
                "Tipe Akun": current_account_type,
                "Keterangan": row[12],
                "Debit": clean_number(row[19]),
                "Kredit": clean_number(row[21]),
                "Saldo": clean_number(row[23])
            })

    return pd.DataFrame(processed_rows)

# 4. Antarmuka Web App (Streamlit)
st.set_page_config(page_title="Konversi Buku Besar", layout="wide")
st.title("Aplikasi Konversi Buku Besar (Accurate)")
st.markdown("Support file: **.csv**, **.xls** (Excel 97-2003), dan **.xlsx**")

# Update file uploader untuk menerima excel juga
uploaded_file = st.file_uploader("Upload File", type=["csv", "xls", "xlsx"])

if uploaded_file:
    with st.spinner('Sedang memproses data...'):
        df_result = parse_ledger(uploaded_file)
    
    if df_result is not None and not df_result.empty:
        st.success("Data berhasil diproses!")
        
        # Tampilkan Preview
        st.write("### Preview Hasil Data (50 Baris Pertama):")
        st.dataframe(df_result.head(50), use_container_width=True)
        
        # Proses Download ke Excel
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_result.to_excel(writer, index=False, sheet_name='Data_Rapi')
            
            # Format kolom uang di Excel output agar rapi (opsional)
            workbook  = writer.book
            worksheet = writer.sheets['Data_Rapi']
            money_fmt = workbook.add_format({'num_format': '#,##0.00'})
            
            # Terapkan format uang ke kolom Debit(E), Kredit(F), Saldo(G)
            # (Ingat index dimulai dari 0, A=0, E=4, F=5, G=6)
            worksheet.set_column(4, 6, 18, money_fmt) 
            worksheet.set_column(0, 3, 20) # Lebarkan kolom teks

        st.download_button(
            label="📥 Download File Excel (.xlsx)",
            data=buffer,
            file_name="Laporan_Buku_Besar_Rapi.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("File kosong atau format tidak dikenali.")