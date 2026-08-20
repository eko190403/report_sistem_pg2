import pandas as pd
import time

report_path = r"C:\xampp\htdocs\data\hok\HKO4.XLSX"
output_path = "Fast_Pivot_Test_XlsxWriter.xlsx"

t0 = time.time()
print("Reading...")
df = pd.read_excel(report_path, sheet_name='Sheet1')
print(f"Read took {time.time()-t0:.2f}s")

t1 = time.time()
pivot_df = pd.pivot_table(df, values='Total Biaya', index=['Pers.No.'], columns=['Start Date'], aggfunc='sum')
pivot_df['Grand Total'] = pivot_df.sum(axis=1)
date_columns = [col for col in pivot_df.columns if col != 'Grand Total']
pivot_df['Kehadiran (>20)'] = (pivot_df[date_columns] > 20).sum(axis=1)
pivot_df = pivot_df[pivot_df['Kehadiran (>20)'] > 0]
pivot_df = pivot_df.reset_index()
print(f"Pivot took {time.time()-t1:.2f}s")

t2 = time.time()
print("Writing using xlsxwriter...")
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
    df.to_excel(writer, sheet_name='Data Awal', index=False)
    pivot_df.to_excel(writer, sheet_name='Pivot HKO', index=False)

print(f"Write took {time.time()-t2:.2f}s")
print(f"Total time: {time.time()-t0:.2f}s")
