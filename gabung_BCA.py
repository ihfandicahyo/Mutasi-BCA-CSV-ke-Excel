import pandas as pd
import glob
import os
import re

def clean_merge_excel_strict_recalc():
    try:
        print("--- PROGRAM PENGGABUNGAN DATA BCA ---")
        
        files = glob.glob("*.xlsx")
        output_filename = 'Hasil_Gabungan_Mutasi_BCA.xlsx'
        
        files = [f for f in files if output_filename not in f and not f.startswith('~$')]
        
        if not files:
            print("File Excel (.xlsx) tidak ditemukan.")
            return

        print(f"Ditemukan {len(files)} file Excel. Memulai pemrosesan...")
        
        grouped_data = {}

        for file in files:
            try:
                base_name = os.path.basename(file)
                match = re.search(r'(\d+)', base_name)
                group_name = match.group(1) if match else "Lainnya"

                temp_df = pd.read_excel(file, header=None, nrows=25)
                
                header_idx = -1
                for idx, row in temp_df.iterrows():
                    row_str = row.astype(str).str.cat(sep=' ')
                    if 'Tanggal Transaksi' in row_str and 'Keterangan' in row_str:
                        header_idx = idx
                        break
                
                if header_idx == -1: continue

                df = pd.read_excel(file, skiprows=header_idx, dtype=str)
                df.columns = [str(c).strip() for c in df.columns]
                
                date_pattern = r'\d{2}/\d{2}/\d{4}'
                if 'Tanggal Transaksi' in df.columns:
                    df = df[df['Tanggal Transaksi'].astype(str).str.match(date_pattern, na=False)]
                    
                    if 'Jumlah' in df.columns:
                        def clean_money(val):
                            if not isinstance(val, str): return val
                            val = val.replace(',', '')
                            if 'CR' in val: return float(val.replace('CR', '').strip())
                            if 'DB' in val: return -float(val.replace('DB', '').strip())
                            try: return float(val)
                            except: return 0.0

                        df['Temp_Jumlah'] = df['Jumlah'].apply(clean_money)
                        df['Debit'] = df['Temp_Jumlah'].apply(lambda x: abs(x) if x < 0 else 0)
                        df['Kredit'] = df['Temp_Jumlah'].apply(lambda x: x if x >= 0 else 0)
                        
                        if 'Saldo' in df.columns:
                            df['Saldo'] = df['Saldo'].apply(clean_money)
                        else:
                            df['Saldo'] = 0.0
                    
                    cols_needed = ['Tanggal Transaksi', 'Keterangan', 'Cabang', 'Debit', 'Kredit', 'Saldo']
                    final_cols = [c for c in cols_needed if c in df.columns]
                    df = df[final_cols].copy()
                    
                    df['Tanggal Transaksi'] = pd.to_datetime(df['Tanggal Transaksi'], dayfirst=True)
                    df['_original_idx'] = df.index
                    
                    if group_name not in grouped_data:
                        grouped_data[group_name] = []
                    grouped_data[group_name].append(df)
                    
            except Exception as e:
                print(f"Error pada file {file}: {e}")

        if grouped_data:
            print("Mengurutkan data dan menghitung ulang saldo...")
            writer = pd.ExcelWriter(output_filename, engine='xlsxwriter')
            workbook = writer.book
            
            money_fmt = workbook.add_format({'num_format': '#,##0.00'})
            text_fmt = workbook.add_format({'num_format': '@'})

            sorted_keys = sorted(grouped_data.keys())

            for sheet_name in sorted_keys:
                df_list = grouped_data[sheet_name]
                final_df = pd.concat(df_list, ignore_index=True)
                
                final_df = final_df.sort_values(by=['Tanggal Transaksi', '_original_idx'], ascending=[True, True]).reset_index(drop=True)
                
                if not final_df.empty:
                    first_saldo = final_df.loc[0, 'Saldo']
                    first_debit = final_df.loc[0, 'Debit']
                    first_kredit = final_df.loc[0, 'Kredit']
                    
                    saldo_awal = first_saldo - (first_kredit - first_debit)
                    net_flow = final_df['Kredit'] - final_df['Debit']
                    final_df['Saldo'] = saldo_awal + net_flow.cumsum()

                final_df.drop(columns=['_original_idx'], inplace=True, errors='ignore')
                final_df['Tanggal Transaksi'] = final_df['Tanggal Transaksi'].dt.strftime('%d/%m/%Y')
                
                final_df.to_excel(writer, index=False, sheet_name=sheet_name)
                worksheet = writer.sheets[sheet_name]

                for i, col in enumerate(final_df.columns):
                    column_len = max(final_df[col].astype(str).map(len).max(), len(col)) + 3
                    if col in ['Debit', 'Kredit', 'Saldo']:
                        worksheet.set_column(i, i, column_len, money_fmt)
                    elif col == 'Tanggal Transaksi':
                        worksheet.set_column(i, i, column_len, text_fmt)
                    else:
                        worksheet.set_column(i, i, column_len)

            writer.close()
            print(f"SUKSES! File tersimpan: {output_filename}")
        else:
            print("Tidak ada data valid.")
            
    except Exception as e:
        print(f"ERROR: {e}")

if __name__ == "__main__":
    clean_merge_excel_strict_recalc()
    input("\nTekan Enter untuk keluar...")