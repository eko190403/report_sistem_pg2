import pandas as pd
from datetime import datetime

file_path = r"C:\xampp\htdocs\data\hok\HKO4.XLSX"

print("Loading data...")
df = pd.read_excel(file_path, sheet_name='Sheet1')

if 'Pers.No.' not in df.columns or 'Start Date' not in df.columns or 'Total Biaya' not in df.columns:
    print("Columns missing!")
    exit(1)

print("Pivoting data...")
pivot_df = pd.pivot_table(df, values='Total Biaya', index=['Pers.No.'], columns=['Start Date'], aggfunc='sum')

# Add Grand Total
pivot_df['Grand Total'] = pivot_df.sum(axis=1)

# Count if > 20
date_columns = [col for col in pivot_df.columns if col != 'Grand Total']
pivot_df['Kehadiran'] = (pivot_df[date_columns] > 20).sum(axis=1)

# Format column names to string if they are datetime
pivot_df.columns = [col.strftime('%Y-%m-%d') if isinstance(col, pd.Timestamp) or isinstance(col, datetime) else col for col in pivot_df.columns]

pivot_df = pivot_df.reset_index()

print("Result preview:")
print(pivot_df.head(5))

output_path = "Pivot_HKO_Test.xlsx"
pivot_df.to_excel(output_path, index=False)
print(f"Saved to {output_path}")
