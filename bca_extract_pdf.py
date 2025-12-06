import pdfplumber
import pandas as pd
import re
import glob # Untuk mencari file
import os   # Untuk cek waktu file

def extract_bca_statement(pdf_path, output_path):
    # Regex untuk mendeteksi baris yang diawali tanggal (Format DD/MM)
    date_pattern = re.compile(r'^(\d{2}/\d{2})')
    
    # Regex untuk menangkap angka keuangan
    amount_pattern = re.compile(r'([\d,]+\.\d{2}(?:\s?DB)?)')

    transactions = []
    
    current_tx = {
        "Tanggal": None,
        "Keterangan": "",
        "Mutasi": None,
        "Saldo": None,
        "Tipe": None
    }

    print(f"Memproses file: {pdf_path}...")

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text: continue # Skip halaman kosong
            lines = text.split('\n')
            
            for line in lines:
                # 1. Cek apakah baris dimulai dengan Tanggal
                date_match = date_pattern.match(line)
                
                if date_match:
                    # Simpan transaksi sebelumnya
                    if current_tx["Tanggal"] is not None:
                        transactions.append(current_tx)
                    
                    # Reset untuk transaksi baru
                    current_tx = {
                        "Tanggal": date_match.group(1) + "/2025", 
                        "Keterangan": "",
                        "Mutasi": 0,
                        "Saldo": 0,
                        "Tipe": "CR"
                    }
                    
                    # Cari angka
                    amounts = amount_pattern.findall(line)
                    clean_line = line
                    
                    if len(amounts) >= 1:
                        # Angka terakhir adalah Saldo
                        saldo_str = amounts[-1]
                        # Hapus koma agar bisa jadi float (asumsi format 1,000.00)
                        current_tx["Saldo"] = float(saldo_str.replace(',', '').replace(' DB', ''))
                        
                        if len(amounts) >= 2:
                            # Angka kedua dari belakang adalah Mutasi
                            mutasi_str = amounts[-2]
                            is_db = "DB" in mutasi_str
                            
                            nominal = float(mutasi_str.replace(',', '').replace(' DB', ''))
                            current_tx["Mutasi"] = nominal
                            current_tx["Tipe"] = "DB" if is_db else "CR"
                            
                            clean_line = clean_line.replace(saldo_str, '').replace(mutasi_str, '')
                    
                    # Bersihkan tanggal dari keterangan
                    clean_line = clean_line.replace(date_match.group(1), '', 1).strip()
                    current_tx["Keterangan"] += clean_line + " "
                    
                else:
                    # 2. Baris lanjutan (multiline)
                    if "BCA" in line or "SALDO AWAL" in line or "HALAMAN" in line:
                        continue
                        
                    if current_tx["Tanggal"] is not None:
                        amounts = amount_pattern.findall(line)
                        if not amounts: 
                            current_tx["Keterangan"] += line.strip() + " "

    # Append transaksi terakhir
    if current_tx["Tanggal"] is not None:
        transactions.append(current_tx)

    # Buat DataFrame
    df = pd.DataFrame(transactions)

    # Cek jika DataFrame kosong
    if df.empty:
        print("Tidak ada transaksi yang ditemukan atau format PDF tidak sesuai.")
        return

    # Data Cleaning
    df['Keterangan'] = df['Keterangan'].str.strip()
    df['Debet'] = df.apply(lambda x: x['Mutasi'] if x['Tipe'] == 'DB' else 0, axis=1)
    df['Kredit'] = df.apply(lambda x: x['Mutasi'] if x['Tipe'] == 'CR' else 0, axis=1)
    
    # Reorder columns: A=Tanggal, B=Keterangan, C=Debet, D=Kredit, E=Saldo
    final_cols = ['Tanggal', 'Keterangan', 'Debet', 'Kredit', 'Saldo']
    df = df[final_cols]
    
    print(f"Menyimpan ke Excel: {output_path}")
    
    # Gunakan ExcelWriter dengan engine xlsxwriter
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Mutasi')
        
        workbook = writer.book
        worksheet = writer.sheets['Mutasi']
        
        # Format Angka: Menggunakan pemisah ribuan dan 2 desimal
        number_format = workbook.add_format({'num_format': '#,##0.00'})
        
        # Auto-fit Column Width
        for i, col in enumerate(df.columns):
            # Hitung panjang maksimum teks di kolom tersebut
            max_len = max(
                df[col].astype(str).map(len).max(), # Panjang data
                len(col) # Panjang header
            ) + 2 
            
            # Terapkan lebar kolom
            if i in [2, 3, 4]: # Kolom Debet, Kredit, Saldo
                worksheet.set_column(i, i, max_len, number_format)
            else:
                worksheet.set_column(i, i, max_len)

    print("Berhasil dikonversi!")

# --- EKSEKUSI OTOMATIS ---

def main():
    # 1. Cari semua file PDF di folder ini
    pdf_files = glob.glob("*.pdf")
    
    if not pdf_files:
        print("Tidak ada file PDF ditemukan di folder ini.")
        return

    # 2. Ambil file PDF yang paling baru (berdasarkan waktu modifikasi)
    # Ini berguna jika folder Downloads menumpuk, dia akan ambil yang barusan diunduh.
    latest_file = max(pdf_files, key=os.path.getctime)
    
    print(f"File terdeteksi: {latest_file}")
    
    # 3. Buat nama output dinamis agar tidak menimpa file lama
    # Contoh: 'Statement_Jan.pdf' -> 'Statement_Jan_Excel.xlsx'
    file_output = latest_file.replace('.pdf', '_Excel.xlsx')
    
    # Cek apakah file output sudah ada (opsional, untuk safety)
    if os.path.exists(file_output):
        response = input(f"File '{file_output}' sudah ada. Timpa? (y/n): ")
        if response.lower() != 'y':
            print("Dibatalkan pengguna.")
            return

    try:
        extract_bca_statement(latest_file, file_output)
    except Exception as e:
        print(f"Terjadi error: {e}")

if __name__ == "__main__":
    main()