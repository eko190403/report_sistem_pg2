import pandas as pd
file_path = r"C:\xampp\htdocs\data\hok\HKO4.XLSX"
df = pd.read_excel(file_path, sheet_name='Sheet1')
print(f"Number of rows in HKO4.XLSX: {len(df)}")
