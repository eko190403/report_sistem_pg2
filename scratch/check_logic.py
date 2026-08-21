import pandas as pd
df = pd.read_excel('Overtime 7 2026.XLSX', engine='calamine')

# Sheet 1: Absence Type
sheet1 = df.dropna(subset=['Absence Type'])
print(f"Sheet 1 (Absence Type) potential rows: {len(sheet1)}")

# Sheet 2: Piket
df['Total Piket'] = df['Jam Piket Biasa'].fillna(0) + df['Jam Piket Libur'].fillna(0)
sheet2 = df[df['Total Piket'] > 0]
print(f"Sheet 2 (Piket) potential rows: {len(sheet2)}")
