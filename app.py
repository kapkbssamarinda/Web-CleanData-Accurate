import streamlit as st
import pandas as pd
import csv
import io
import time

# 1. Fungsi Pembersihan Nominal (Sudah Diperbaiki Jadi Angka Murni)
def clean_currency(val):
    if not val or pd.isna(val) or str(val).strip() == "" or str(val).strip() == "-":
        return 0 
    
    val = str(val).replace('Rp', '').replace('"', '').strip()
    val = val.replace(',', '') 
    
    if val.endswith('.00'):
        val = val[:-3]
    if val.endswith('cr'):
        val = val.replace('cr', '').strip()
        
    try:
        num = int(float(val))
        return num
    except ValueError:
        return 0

# 2. Fungsi Pembersihan Tanggal
def clean_date(val):
    if not val:
        return ""
    return str(val).split(" ")[0].strip()

# 3. Fungsi Utama Proses Data dengan Progress Bar
def process_gl_data(data_rows):
    parsed_data = []
    current_coa = ""
    current_account = ""
    start_reading = False
    
    idx_id, idx_date, idx_memo, idx_debit, idx_credit, idx_eb = 0, 2, 3, 4, 5, 8
    
    total_rows = len(data_rows)
    progress_bar = st.progress(0, text="Memulai proses cleaning...")
    
    for i, row in enumerate(data_rows):
        # Update progress bar secara berkala
        percent_complete = (i + 1) / total_rows
        progress_bar.progress(percent_complete, text=f"Cleaning data baris ke-{i+1} dari {total_rows}...")
        
        clean_row = [str(cell).strip() for cell in row]
        
        if all(cell == "" for cell in clean_row):
            continue
            
        # Deteksi Header
        if "ID#" in clean_row and "Date" in clean_row:
            idx_id = clean_row.index("ID#")
            idx_date = clean_row.index("Date")
            idx_memo = clean_row.index("Memo")
            idx_debit = clean_row.index("Debit")
            idx_credit = clean_row.index("Credit")
            for i_col, col in enumerate(clean_row):
                if "Ending Balance" in col:
                    idx_eb = i_col
                    break
            start_reading = True
            continue
            
        if not start_reading:
            continue
            
        # Deteksi COA
        coa_idx = -1
        for i_cell, cell in enumerate(clean_row):
            if "-" in cell and cell[0].isdigit() and len(cell) >= 5:
                coa_idx = i_cell
                break
                
        if coa_idx != -1 and (len(clean_row) <= idx_date or clean_row[idx_date] == ""):
            current_coa = clean_row[coa_idx]
            current_account = ""
            for cell in clean_row[coa_idx+1:]:
                if cell != "":
                    current_account = cell
                    break
            continue
            
        # Deteksi Saldo Awal
        bb_idx = -1
        for i_cell, cell in enumerate(clean_row):
            if "Beginning Balance" in cell:
                bb_idx = i_cell
                break
                
        if bb_idx != -1:
            balance = ""
            for cell in clean_row[bb_idx+1:]:
                if cell != "" and cell != ":":
                    balance = cell
                    break
            if not balance and len(clean_row) > idx_eb:
                balance = clean_row[idx_eb]
                
            parsed_data.append({
                "ID": "-", "Tanggal": "-", "COA": current_coa, "Nama Akun": current_account,
                "Memo": "Beginning Balance (Saldo Awal)", "Debit": 0, "Kredit": 0, # Menggunakan 0 murni
                "Ending Balance": clean_currency(balance)
            })
            continue
            
        # Deteksi Baris Transaksi
        if len(clean_row) > idx_id and clean_row[idx_id].isdigit():
            _id = clean_row[idx_id]
            _date = clean_row[idx_date] if len(clean_row) > idx_date else ""
            _memo = clean_row[idx_memo] if len(clean_row) > idx_memo else ""
            _debit = clean_row[idx_debit] if len(clean_row) > idx_debit else ""
            _credit = clean_row[idx_credit] if len(clean_row) > idx_credit else ""
            _eb = clean_row[idx_eb] if len(clean_row) > idx_eb else ""
            if not _eb and len(clean_row) > idx_eb + 1:
                _eb = clean_row[idx_eb + 1]
                
            parsed_data.append({
                "ID": _id, "Tanggal": clean_date(_date), "COA": current_coa, "Nama Akun": current_account,
                "Memo": _memo, "Debit": clean_currency(_debit), "Kredit": clean_currency(_credit),
                "Ending Balance": clean_currency(_eb)
            })
    
    progress_bar.empty() # Hapus progress bar setelah selesai
    return pd.DataFrame(parsed_data)

# ================= UI STREAMLIT =================

st.set_page_config(page_title="GL Converter Pro", page_icon="📄", layout="wide")

st.title("📄 Ekstrak Data Buku Besar MyOb")
st.markdown("Pastikan ekstrak dari myob berupa file .txt agar dapat diproses di web-app ini")

uploaded_file = st.file_uploader("Upload file General Ledger (.txt)", type=['txt'])

if uploaded_file is not None:
    # --- TAHAP 1: IMPORT ---
    with st.status("Mengimport file...", expanded=True) as status:
        st.write("Membaca isi file ke memori...")
        file_content = uploaded_file.getvalue().decode("utf-8", errors="replace")
        reader = csv.reader(file_content.splitlines())
        data_rows = list(reader)
        time.sleep(0.5) # Delay kosmetik agar user bisa melihat prosesnya
        status.update(label="Import Selesai!", state="complete", expanded=False)

    # --- TAHAP 2: CLEANING ---
    df = process_gl_data(data_rows)
    
    if not df.empty:
        st.success(f"Berhasil membersihkan {len(df)} data transaksi.")
        st.dataframe(df, use_container_width=True)
        
        # --- TAHAP 3: EXPORT (Sudah Diperbaiki Dengan XlsxWriter Formatting) ---
        with st.status("Menyiapkan file Excel...", expanded=True) as status_export:
            st.write("Membuat format kolom dan styling...")
            output = io.BytesIO()
            
            # Memastikan engine xlsxwriter digunakan
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='GL_Detail')
                
                workbook = writer.book
                worksheet = writer.sheets['GL_Detail']
                
                # Buat format angka akuntansi standar (pemisah ribuan, 0 menjadi strip)
                accounting_format = workbook.add_format({'num_format': '#,##0;-#,##0;"-"'})
                
                # Simulasi progress kecil untuk styling
                for i, col in enumerate(df.columns):
                    column_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
                    
                    # Terapkan format HANYA pada kolom nominal
                    if col in ["Debit", "Kredit", "Ending Balance"]:
                        worksheet.set_column(i, i, min(column_len, 50), accounting_format)
                    else:
                        worksheet.set_column(i, i, min(column_len, 50))
                        
                time.sleep(0.8) # Delay kosmetik
                
            processed_data = output.getvalue()
            status_export.update(label="File Siap Diunduh!", state="complete", expanded=False)

        st.download_button(
            label="📥 Download Hasil (.xlsx)",
            data=processed_data,
            file_name=f"Cleaned_{uploaded_file.name.split('.')[0]}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Data gagal diekstrak. Mohon periksa kembali isi file .txt Anda.")