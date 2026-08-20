import pandas as pd
file_path = r"C:\xampp\htdocs\data\hok\HKO4.XLSX"
try:
    df = pd.read_excel(file_path, skiprows=2, nrows=10)
    print("Columns:", df.columns.tolist()[:10])
    print("\nData:")
    print(df.head(10).iloc[:, :10].to_string())
    print("\nLast 5 columns:")
    print(df.head(10).iloc[:, -5:].to_string())
except Exception as e:
    print("Error:", e)
