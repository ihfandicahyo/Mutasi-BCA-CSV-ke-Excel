import pandas as pd
import os
import glob
import re
from openpyxl.utils import get_column_letter 

def split_db_cr_columns(df):
    """
    Fungsi untuk memindai kolom yang berisi nominal + DB/CR (misal: "100.00 CR")
    dan memecahnya menjadi dua kolom terpisah (DB dan CR).
    """
    processed_data = []
    
    # Pola Regex: Mencari angka (termasuk koma/titik) diikuti spasi dan DB atau CR
    # Flags re.IGNORECASE agar tidak peduli huruf besar/kecil (db, DB, Cr, CR)
    pattern = re.compile(r'^\s*([\d.,]+)\s*(DB|CR)\s*$', re.IGNORECASE)

    for col in df.columns:
        # Ambil data kolom sebagai string
        col_data = df[col].astype(str)
        
        # Cek apakah kolom ini mayoritas berisi pola DB/CR
        # Cek sampel 20 baris pertama (atau kurang jika data sedikit) untuk efisiensi
        sample = col_data.head(20)
        matches = sample.apply(lambda x: bool(pattern.match(x)) if x and x.lower() != 'nan' else False)
        
        # Jika ditemukan pola yang cocok pada sampel, maka akan memproses seluruh kolom
        if matches.any():
            print(f"      -> Mendeteksi kolom '{col}' sebagai format DB/CR. Memecah kolom...")
            
            # Membuat daftar untuk menampung data baru
            db_values = []
            cr_values = []
            
            for val in col_data:
                match = pattern.match(val) if isinstance(val, str) else None
                if match:
                    nominal_str = match.group(1).replace(',', '') # Hapus koma pemisah ribuan
                    tipe = match.group(2).upper()
                    
                    try:
                        nominal = float(nominal_str)
                    except ValueError:
                        nominal = 0
                        
                    if tipe == 'DB':
                        db_values.append(nominal)
                        cr_values.append(None) # Atau 0 jika ingin angka 0
                    elif tipe == 'CR':
                        db_values.append(None)
                        cr_values.append(nominal)
                else:
                    # Jika data tidak cocok dengan pola (misal header atau kosong), biarkan kosong
                    db_values.append(None)
                    cr_values.append(None)
            
            # Menambahkan sebagai DataFrame baru (2 Kolom)
            temp_df = pd.DataFrame({'DB': db_values, 'CR': cr_values})
            processed_data.append(temp_df)
            
        else:
            # Jika bukan kolom DB/CR, biarkan seperti aslinya
            processed_data.append(df[col])

    # Menggabungkan kembali semua kolom
    if processed_data:
        return pd.concat(processed_data, axis=1)
    return df

def convert_csv_to_excel_autofit():
    # Atur folder saat ini
    current_folder = os.path.dirname(os.path.abspath(__file__))
    print(f"Mencari file CSV di folder: {current_folder}")

    # Cari semua file .csv
    csv_files = glob.glob(os.path.join(current_folder, "*.csv"))

    if not csv_files:
        print("Tidak ditemukan file CSV di folder ini.")
        return

    print(f"Ditemukan {len(csv_files)} file CSV. Memulai proses...\n")

    for input_path in csv_files:
        try:
            base_name = os.path.splitext(os.path.basename(input_path))[0]
            output_filename = f"{base_name}.xlsx"
            output_path = os.path.join(current_folder, output_filename)

            print(f"-> Memproses: {os.path.basename(input_path)} ...")

            # Baca data
            df = pd.read_csv(input_path, sep=None, engine='python', quotechar='"', header=None)

            # --- Memproses pembersihan data ---
            def clean_text(text):
                if isinstance(text, str):
                    return text.strip().strip(',')
                return text
            df_clean = df.map(clean_text)

            # --- Memproses split DB/CR (Baru) ---
            df_final = split_db_cr_columns(df_clean)

            # Simpan ke Excel dengan auto-fit
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                # Header=True agar label "DB" dan "CR" tertulis di baris pertama
                # Index=False agar nomor baris pandas tidak ikut
                
                # Cek apakah kita punya header hasil generate (DB/CR) atau integer murni
                has_named_columns = any(isinstance(c, str) for c in df_final.columns)
                
                df_final.to_excel(writer, index=False, header=has_named_columns, sheet_name='Sheet1')
                
                # Akses worksheet
                worksheet = writer.sheets['Sheet1']

                # Loop setiap kolom untuk mencari teks terpanjang (Auto-fit)
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = get_column_letter(column[0].column)
                    
                    for cell in column:
                        try:
                            if cell.value:
                                cell_len = len(str(cell.value))
                                if cell_len > max_length:
                                    max_length = cell_len
                        except:
                            pass
                    
                    adjusted_width = (max_length + 2) 
                    worksheet.column_dimensions[column_letter].width = adjusted_width

            print(f"   [BERHASIL] Disimpan & Auto-fit: {output_filename}")

        except Exception as e:
            print(f"   [GAGAL] Error pada file {os.path.basename(input_path)}: {e}")
        
        print("-" * 40)

    print("\nSemua proses selesai!")

if __name__ == "__main__":
    convert_csv_to_excel_autofit()