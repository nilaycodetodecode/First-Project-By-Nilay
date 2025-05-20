import pandas as pd

# Read the Excel file into a DataFrame
df = pd.read_excel(r"C:\Users\bnila\Downloads\task_.xlsx")
import os
from openpyxl import load_workbook

def validate_excel_file(file_path):
    """Validate basic Excel file properties"""
    
    # Check if file exists
    if not os.path.exists(file_path):
        return False, "File does not exist"
    
    # Check if it's an Excel file
    if not file_path.lower().endswith(('.xlsx', '.xls')):
        return False, "Not an Excel file"
    
    try:
        # Try to load the workbook
        wb = load_workbook(file_path)
        return True, "File is valid"
    except Exception as e:
        return False, f"Invalid Excel file: {str(e)}"

# Usage
is_valid, message = validate_excel_file(r"C:\Users\bnila\Downloads\task_.xlsx")
print(f"Validation result: {is_valid}, Message: {message}")
# Print the DataFrame
print(df)
df.rename(columns={"ERP USER ID": "t_user", "PROJECT": "t_cpry", "DICIPLINE": "t_htyp", "REMARKS": "t_rema", "DATE": "t_date","START TIME": "t_fsts", "END TIME": "t_nstp"}, inplace= True)
df.to_excel(r"C:\Users\bnila\Downloads\task_niloy.xlsx", index=False)

print(df.columns)
# Ensure columns are in datetime format
df['t_fsts'] = pd.to_datetime(df['t_fsts'])
df['t_nstp'] = pd.to_datetime(df['t_nstp'])
df.to_excel(r"C:\Users\bnila\Downloads\task_niloy.xlsx", index=False)
print(df.columns)
 
print("\nData types after conversion:")
print(df.dtypes)

# 3. Calculate the difference
df['t_fsts - t_nstp'] = df['t_nstp'] - df['t_fsts']

print("\nDataFrame with Time Difference (Timedelta object):")
print(df)
print("\nData type of 'Time Difference':", df['t_fsts - t_nstp'].dtype)
df.to_excel(r"C:\Users\bnila\Downloads\task_niloy.xlsx", index=False)
