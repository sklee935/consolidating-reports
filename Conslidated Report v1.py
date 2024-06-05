import pandas as pd
from openpyxl import load_workbook
import os

# Set file paths
input_file_path = r"C:\Users\slee\OneDrive - SBP\Tax Returns\xSungkeun\Monthly Task\2024 FA Addition & Disposal recon - Alex US\Alex US TB's Input Template.xlsx"
output_file_path = r"C:\Users\slee\OneDrive - SBP\Tax Returns\xSungkeun\Monthly Task\2024 FA Addition & Disposal recon - Alex US\Alex US TB's Output Template.xlsx"

# List of sheet names to process
sheets = ['West', 'NE', 'MW', 'NSSUS', 'Direct']

# Collecting unique 'Ledger account' and 'Name' from all sheets
ledger_accounts = []
for sheet_name in sheets:
    df_temp = pd.read_excel(input_file_path, sheet_name=sheet_name, usecols=['Ledger account', 'Name'], skiprows=0)
    ledger_accounts.append(df_temp[['Ledger account', 'Name']].drop_duplicates().set_index('Ledger account'))

# Creating a dataframe with unique ledger accounts and names
ledger_accounts_df = pd.concat(ledger_accounts).drop_duplicates()
consolidated_df = ledger_accounts_df.reset_index().drop_duplicates().set_index('Ledger account')

# Processing each sheet and summing up
for sheet_name in sheets:
    # Reading required columns from the sheet
    df = pd.read_excel(input_file_path, sheet_name=sheet_name, usecols=['Ledger account', 'Opening balance', 'February'], skiprows=0)
    df.set_index('Ledger account', inplace=True)
    # Calculating 'Closing Balance' for each sheet by summing 'Opening balance' and 'February', handling any non-numeric values gracefully with `pd.to_numeric` and `fillna(0)`
    df['Closing Balance_' + sheet_name] = pd.to_numeric(df['Opening balance'], errors='coerce').fillna(0) + pd.to_numeric(df['February'], errors='coerce').fillna(0)
    # Joining calculated 'Closing Balance' with the consolidated dataframe
    consolidated_df = consolidated_df.join(df[['Closing Balance_' + sheet_name]], how='outer')

# Replacing missing values (NA) with 0
consolidated_df.fillna(0, inplace=True)

# Defining the order of columns for the final dataframe
column_order = ['Name'] + ['Closing Balance_' + sheet for sheet in sheets] + ['Subtotal']
# Calculating 'Subtotal' as the sum of all 'Closing Balance' columns
consolidated_df['Subtotal'] = consolidated_df.filter(like='Closing Balance').sum(axis=1)
# Reordering the columns as per the defined order and including 'Ledger account' as the first column
consolidated_df = consolidated_df.reset_index()[['Ledger account'] + column_order]

# Saving the results to a new file
if not os.path.exists(output_file_path):
    # If the output file doesn't exist, create it and write the dataframe
    with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
        consolidated_df.to_excel(writer, index=False, sheet_name='Alex US Cons')
else:
    # If the output file exists, open it and replace the 'Alex US Cons' sheet with the updated dataframe
    book = load_workbook(output_file_path)
    with pd.ExcelWriter(output_file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        writer.book = book
        consolidated_df.to_excel(writer, index=False, sheet_name='Alex US Cons')

print("Process completed!")
