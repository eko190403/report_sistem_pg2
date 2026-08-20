import os
from flask import Flask, render_template, request, send_file, flash, redirect, url_for
import pandas as pd
from openpyxl.styles import PatternFill, Alignment, Font
from openpyxl.utils.dataframe import dataframe_to_rows
import openpyxl
import tempfile
import shutil
from datetime import datetime
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = 'super_secret_key_change_me_in_production'

# Pastikan folder templates dan static ada
os.makedirs('templates', exist_ok=True)
os.makedirs('static', exist_ok=True)

def process_data_logic(master_path, report_path, output_path):
    # 1. Load Master Data
    master_df = pd.read_excel(master_path, sheet_name='Sheet1')
    master_df['Pers.No.'] = master_df['Pers.No.'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    master_df['Kode Mandor'] = master_df['Kode Mandor'].fillna('').astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    
    name_lookup_pers = master_df.set_index('Pers.No.')['Full Name'].to_dict()
    jabatan_lookup = master_df.set_index('Pers.No.')['ZJABATAN'].to_dict()
    bagian_lookup = master_df.set_index('Pers.No.')['Organizational Unit'].to_dict()
    
    master_mandor = master_df[master_df['Kode Mandor'] != '']
    name_lookup_kode = master_mandor.set_index('Kode Mandor')['Nama mandor'].to_dict()

    # 2. Load Report Data
    xl_report = pd.ExcelFile(report_path)
    report_df = xl_report.parse('Worksheet')
    
    report_df['KIT TK'] = report_df['KIT TK'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    report_df['Mandor'] = report_df['Mandor'].fillna('').astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    report_df['Nama TK'] = report_df['Nama TK'].astype(str).str.strip() # Menghapus spasi tersembunyi
    
    # 3. Hapus kolom jika sudah ada dari percobaan sebelumnya
    cols_to_drop = ['nama mandor', 'Jabatan', 'bagian', 'Totall']
    for col in cols_to_drop:
        if col in report_df.columns:
            report_df = report_df.drop(columns=[col])
            
    # 4. Tambahkan Kolom Baru
    mapped_by_pers = report_df['Mandor'].map(name_lookup_pers)
    mapped_by_kode = report_df['Mandor'].map(name_lookup_kode)
    nama_mandor_series = mapped_by_pers.fillna(mapped_by_kode)
    
    jabatan_series = report_df['KIT TK'].map(jabatan_lookup)
    bagian_series = report_df['KIT TK'].map(bagian_lookup)
    
    date_counts = report_df.groupby('KIT TK')['Date Shift'].nunique().to_dict()
    totall_series = report_df['KIT TK'].map(date_counts)
    
    col_mandor_idx = report_df.columns.get_loc('Mandor')
    report_df.insert(col_mandor_idx + 1, 'nama mandor', nama_mandor_series)
    report_df.insert(col_mandor_idx + 2, 'Jabatan', jabatan_series)
    report_df.insert(col_mandor_idx + 3, 'bagian', bagian_series)
    report_df.insert(col_mandor_idx + 4, 'Totall', totall_series)
    
    # Urutkan data berdasarkan Totall (Terbesar ke Terkecil) lalu berdasarkan Nama TK (A-Z)
    report_df = report_df.sort_values(by=['Totall', 'Nama TK'], ascending=[False, True])

    # Ganti NaNs dengan string kosong agar tidak error di openpyxl
    report_df = report_df.fillna("")

    # Gunakan file Template yang sudah mengandung PivotTable asli
    template_path = os.path.join(app.root_path, 'Template_Report.xlsx')
    shutil.copy(template_path, output_path)

    # Buka file hasil kopian menggunakan openpyxl
    wb = openpyxl.load_workbook(output_path)
    
    if 'Worksheet' in wb.sheetnames:
        ws = wb['Worksheet']
        # Hapus data lama di Worksheet (sisakan header di baris 1)
        ws.delete_rows(2, ws.max_row)
        
        # Tulis data baru dari pandas DataFrame mulai baris 2
        for r_idx, row in enumerate(dataframe_to_rows(report_df, index=False, header=False), 2):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
                
        # Format excel cells
        green_fill = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")
        center_alignment = Alignment(horizontal="center", vertical="center")
        header_font = Font(color="000000", bold=True)
        
        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = center_alignment
                if cell.row == 1:
                    cell.fill = green_fill
                    cell.font = header_font
                    
        # Auto-adjust column widths roughly
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            ws.column_dimensions[column].width = max_length + 2

    wb.save(output_path)
    return output_path

def process_hko_logic(report_path, output_path):
    df = pd.read_excel(report_path, sheet_name='Sheet1')
    
    if 'Pers.No.' not in df.columns or 'Start Date' not in df.columns or 'Total Biaya' not in df.columns:
        raise Exception("Kolom 'Pers.No.', 'Start Date', atau 'Total Biaya' tidak ditemukan di file HKO.")
        
    df = df.dropna(subset=['Pers.No.', 'Start Date'])
    
    pivot_df = pd.pivot_table(df, values='Total Biaya', index=['Pers.No.'], columns=['Start Date'], aggfunc='sum')
    
    pivot_df['Grand Total'] = pivot_df.sum(axis=1)
    
    date_columns = [col for col in pivot_df.columns if col != 'Grand Total']
    pivot_df['Kehadiran (>20)'] = (pivot_df[date_columns] > 20).sum(axis=1)
    
    # Buang data yang total kehadirannya 0
    pivot_df = pivot_df[pivot_df['Kehadiran (>20)'] > 0]
    
    pivot_df.columns = [col.strftime('%d-%m-%Y') if isinstance(col, pd.Timestamp) or isinstance(col, datetime) else col for col in pivot_df.columns]
    pivot_df = pivot_df.reset_index()
    
    pivot_df = pivot_df.fillna("")
    
    # Cara super cepat: Copy file mentahnya langsung, baru kita tambahkan Sheet Pivot!
    import shutil
    shutil.copy(report_path, output_path)
    
    with pd.ExcelWriter(output_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        # Tulis hasil pivot saja
        pivot_df.to_excel(writer, sheet_name='Pivot HKO', index=False)
        
        workbook = writer.book
        worksheet_pivot = writer.sheets['Pivot HKO']
        
        green_fill = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")
        center_alignment = Alignment(horizontal="center", vertical="center")
        header_font = Font(color="000000", bold=True)
        
        # Styling Pivot HKO
        for row in worksheet_pivot.iter_rows():
            for cell in row:
                cell.alignment = center_alignment
                if cell.row == 1:
                    cell.fill = green_fill
                    cell.font = header_font
                    
        for col in worksheet_pivot.columns:
            column = col[0].column_letter
            worksheet_pivot.column_dimensions[column].width = 15
            
    return output_path

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process_files():
    if 'master_file' not in request.files or 'report_file' not in request.files:
        return "Missing files", 400
        
    master_file = request.files['master_file']
    report_file = request.files['report_file']
    
    if master_file.filename == '' or report_file.filename == '':
        return "No selected file", 400
        
    temp_dir = tempfile.mkdtemp()
    
    master_path = os.path.join(temp_dir, secure_filename(master_file.filename))
    report_path = os.path.join(temp_dir, secure_filename(report_file.filename))
    
    # Pastikan ekstensi selalu lowercase agar tidak error di pandas ExcelWriter
    base_name = secure_filename(report_file.filename).rsplit('.', 1)[0]
    output_filename = "Processed_" + base_name + ".xlsx"
    output_path = os.path.join(temp_dir, output_filename)
    
    master_file.save(master_path)
    report_file.save(report_path)
    
    try:
        process_data_logic(master_path, report_path, output_path)
        return send_file(output_path, as_attachment=True, download_name=output_filename)
    except Exception as e:
        return str(e), 500

@app.route('/process_hko', methods=['POST'])
def process_hko():
    if 'hko_file' not in request.files:
        return "Missing file", 400
        
    hko_file = request.files['hko_file']
    if hko_file.filename == '':
        return "No selected file", 400
        
    temp_dir = tempfile.mkdtemp()
    hko_path = os.path.join(temp_dir, secure_filename(hko_file.filename))
    
    # Pastikan ekstensi selalu lowercase agar tidak error di pandas ExcelWriter
    base_name = secure_filename(hko_file.filename).rsplit('.', 1)[0]
    output_filename = "Processed_HKO_" + base_name + ".xlsx"
    output_path = os.path.join(temp_dir, output_filename)
    
    hko_file.save(hko_path)
    
    try:
        process_hko_logic(hko_path, output_path)
        return send_file(output_path, as_attachment=True, download_name=output_filename)
    except Exception as e:
        return str(e), 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)
