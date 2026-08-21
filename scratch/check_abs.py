import pandas as pd
df = pd.read_excel('Overtime 7 2026.XLSX', engine='calamine')
print("Unique Absence Type:", df['Absence Type'].unique())
print("Non-null Absence Type count:", df['Absence Type'].notna().sum())
print("Total rows:", len(df))
