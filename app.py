import os
from flask import Flask, render_template, request, send_file, flash, redirect, url_for
import pandas as pd
from openpyxl.styles import PatternFill, Alignment, Font
import tempfile
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
    
    # Create the Pivot summary
    pivot_df = report_df.groupby(['Nama TK', 'nama mandor', 'Jabatan', 'bagian', 'Date Shift'])['Totall'].sum().reset_index()
    # Sort just like in the screenshot, but prioritized by Totall (largest to smallest)
    pivot_df = pivot_df.sort_values(
        by=['Totall', 'Nama TK', 'nama mandor', 'Jabatan', 'bagian', 'Date Shift'], 
        ascending=[False, True, True, True, True, True]
    )

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        report_df.to_excel(writer, sheet_name='Worksheet', index=False)
        pivot_df.to_excel(writer, sheet_name='Pivot Summary', index=False)
        
        for sheet_name in xl_report.sheet_names:
            if sheet_name != 'Worksheet' and sheet_name != 'Sheet3': # Skip old pivot
                other_df = xl_report.parse(sheet_name)
                other_df.to_excel(writer, sheet_name=sheet_name, index=False)
                
        # Format excel cells
        workbook = writer.book
        green_fill = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")
        center_alignment = Alignment(horizontal="center", vertical="center")
        header_font = Font(color="FFFFFF", bold=True)
        
        for sheet_name in workbook.sheetnames:
            worksheet = workbook[sheet_name]
            for row in worksheet.iter_rows():
                for cell in row:
                    cell.alignment = center_alignment
                    if cell.row == 1:
                        cell.fill = green_fill
                        cell.font = header_font
                        
            # Auto-adjust column widths roughly
            for col in worksheet.columns:
                max_length = 0
                column = col[0].column_letter # Get the column name
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = (max_length + 2)
                worksheet.column_dimensions[column].width = adjusted_width
                
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
        
    # Use temporary directory for processing
    temp_dir = tempfile.mkdtemp()
    
    master_path = os.path.join(temp_dir, secure_filename(master_file.filename))
    report_path = os.path.join(temp_dir, secure_filename(report_file.filename))
    
    output_filename = "Processed_" + secure_filename(report_file.filename)
    output_path = os.path.join(temp_dir, output_filename)
    
    master_file.save(master_path)
    report_file.save(report_path)
    
    try:
        process_data_logic(master_path, report_path, output_path)
        return send_file(output_path, as_attachment=True, download_name=output_filename)
    except Exception as e:
        return str(e), 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)
