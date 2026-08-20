import pandas as pd
file_path = r"C:\xampp\htdocs\data\hok\HKO4.XLSX"
try:
    df = pd.read_excel(file_path, nrows=5)
    print("Columns:", df.columns.tolist())
    print("\nFirst 3 rows:")
    print(df.head(3).to_string())
except Exception as e:
    print("Error:", e)
