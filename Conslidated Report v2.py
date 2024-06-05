import pandas as pd
from openpyxl import load_workbook
import os

# 파일 경로 설정
input_file_path = r"C:\Users\slee\OneDrive - SBP\Tax Returns\xSungkeun\Monthly Task\2024 FA Addition & Disposal recon - Alex US\Alex US TB's Input Template.xlsx"
output_file_path = r"C:\Users\slee\OneDrive - SBP\Tax Returns\xSungkeun\Monthly Task\2024 FA Addition & Disposal recon - Alex US\Alex US TB's Output Template.xlsx"

# 시트 이름 목록
sheets = ['West', 'NE', 'MW', 'NSSUS', 'Direct']

# 모든 시트에서 계정 번호와 이름을 수집
ledger_accounts = []
for sheet_name in sheets:
    df_temp = pd.read_excel(input_file_path, sheet_name=sheet_name, usecols=['Ledger account', 'Name'], skiprows=0)
    ledger_accounts.append(df_temp[['Ledger account', 'Name']].drop_duplicates().set_index('Ledger account'))

# 유니크 계정 번호 및 이름을 가지고 있는 데이터프레임 생성
ledger_accounts_df = pd.concat(ledger_accounts).drop_duplicates()
consolidated_df = ledger_accounts_df.reset_index().drop_duplicates().set_index('Ledger account')

# 각 시트 처리 및 합산
for sheet_name in sheets:
    df = pd.read_excel(input_file_path, sheet_name=sheet_name, usecols=['Ledger account', 'Opening balance', 'February'], skiprows=0)
    df.set_index('Ledger account', inplace=True)
    df['Closing Balance_' + sheet_name] = pd.to_numeric(df['Opening balance'], errors='coerce').fillna(0) + pd.to_numeric(df['February'], errors='coerce').fillna(0)
    consolidated_df = consolidated_df.join(df[['Closing Balance_' + sheet_name]], how='outer')

# 누락된 값(NA)을 0으로 대체
consolidated_df.fillna(0, inplace=True)

# 컬럼 순서 재정의
column_order = ['Name'] + ['Closing Balance_' + sheet for sheet in sheets] + ['Subtotal']
consolidated_df['Subtotal'] = consolidated_df.filter(like='Closing Balance').sum(axis=1)
consolidated_df = consolidated_df.reset_index()[['Ledger account'] + column_order]

# 결과를 새 파일에 저장
if not os.path.exists(output_file_path):
    with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
        consolidated_df.to_excel(writer, index=False, sheet_name='Alex US Cons')
else:
    book = load_workbook(output_file_path)
    with pd.ExcelWriter(output_file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        writer.book = book
        consolidated_df.to_excel(writer, index=False, sheet_name='Alex US Cons')

print("작업 완료!")
