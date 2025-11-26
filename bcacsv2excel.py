import pandas as pd
import os
import glob
from openpyxl.utils import get_column_letter 

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

            # Baca & bersihkan data
            df = pd.read_csv(input_path, sep=None, engine='python', quotechar='"', header=None)

            def clean_text(text):
                if isinstance(text, str):
                    return text.strip().strip(',')
                return text

            df_clean = df.map(clean_text)

            # Simpan ke Excel dengan auto-fit mirip dengan model (ALT+H+O+I)
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                # Tulis data ke worksheet
                df_clean.to_excel(writer, index=False, header=False, sheet_name='Sheet1')
                
                # Akses worksheet
                worksheet = writer.sheets['Sheet1']

                # Loop setiap kolom untuk mencari teks terpanjang
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = get_column_letter(column[0].column)
                    
                    for cell in column:
                        try:
                            # Hitung panjang karakter di sel tersebut
                            if cell.value:
                                cell_len = len(str(cell.value))
                                if cell_len > max_length:
                                    max_length = cell_len
                        except:
                            pass
                    
                    # Atur lebar kolom
                    adjusted_width = (max_length + 2) 
                    worksheet.column_dimensions[column_letter].width = adjusted_width

            print(f"   [BERHASIL] Disimpan & Auto-fit: {output_filename}")

        except Exception as e:
            print(f"   [GAGAL] Error pada file {os.path.basename(input_path)}: {e}")
        
        print("-" * 40)

    print("\nSemua proses selesai!")

if __name__ == "__main__":
    convert_csv_to_excel_autofit()