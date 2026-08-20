import pandas as pd
import time

report_path = r"C:\xampp\htdocs\data\hok\HKO4.XLSX"

print("Reading with default (openpyxl)...")
t0 = time.time()
df1 = pd.read_excel(report_path, sheet_name='Sheet1')
print(f"Default took: {time.time()-t0:.2f}s")

print("Reading with calamine...")
t1 = time.time()
df2 = pd.read_excel(report_path, sheet_name='Sheet1', engine='calamine')
print(f"Calamine took: {time.time()-t1:.2f}s")
