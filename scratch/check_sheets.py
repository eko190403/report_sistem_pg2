import pandas as pd
file_path = r"C:\xampp\htdocs\data\hok\HKO4.XLSX"
try:
    xl = pd.ExcelFile(file_path)
    print("Sheets in HKO4.XLSX:", xl.sheet_names)
except Exception as e:
    print("Error:", e)
