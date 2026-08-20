import pandas as pd
file_path = r"C:\xampp\htdocs\data\hok\HKO4.XLSX"
try:
    df = pd.read_excel(file_path, sheet_name='Sheet1', nrows=5)
    print("Columns:", df.columns.tolist()[:15])
except Exception as e:
    print("Error:", e)
