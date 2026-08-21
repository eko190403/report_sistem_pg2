import pandas as pd
import json

try:
    df = pd.read_excel('Overtime 7 2026.XLSX', engine='calamine')
    
    print("Columns:", df.columns.tolist())
    print("\nShape:", df.shape)
    
    # Print first 5 rows, converting to dict for cleaner output
    print("\nFirst 5 rows:")
    print(df.head(5).to_string())
    
    # Describe numerical columns
    print("\nSummary statistics:")
    print(df.describe().to_string())
    
except Exception as e:
    print("Error reading file:", str(e))
