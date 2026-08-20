import pandas as pd

report_path = r"C:\xampp\htdocs\data\hok\HKO4.XLSX"
df = pd.read_excel(report_path, sheet_name='Sheet1')
df = df.dropna(subset=['Pers.No.', 'Start Date'])

# Convert Start Date to datetime if not already
df['Start Date'] = pd.to_datetime(df['Start Date'])

daily_sum = df.groupby(['Pers.No.', 'Start Date'])['Total Biaya'].sum().reset_index()
valid_days = daily_sum[daily_sum['Total Biaya'] > 20].copy()

valid_days['Bulan'] = valid_days['Start Date'].dt.month
valid_days['Tahun'] = valid_days['Start Date'].dt.year

result_df = valid_days.groupby(['Pers.No.', 'Bulan', 'Tahun']).size().reset_index(name='Kehadiran')

print(result_df.head(10))
print(f"Total rows in result: {len(result_df)}")
