import pandas as pd
from openpyxl import load_workbook
import os

# 파일 경로 설정
input_file_path = "C:\\Users\\slee\\OneDrive - SBP\\Tax Returns\\xSungkeun\\Monthly Task\\2024 FA Addition & Disposal recon - Alex US\\Alex US TB's Input Template.xlsx"
output_file_path = "C:\\Users\\slee\\OneDrive - SBP\\Tax Returns\\xSungkeun\\Monthly Task\\2024 FA Addition & Disposal recon - Alex US\\Alex US TB's Output Template.xlsx"

# 시트 이름 목록
sheets = ['West', 'NE', 'MW', 'NSSUS', 'Direct']

# 모든 시트에서 계정 번호를 수집
all_ledger_accounts = set()
for sheet_name in sheets:
    df_temp = pd.read_excel(input_file_path, sheet_name=sheet_name)
    all_ledger_accounts.update(df_temp.iloc[:, 0].tolist())  # A열(0번째 열)에서 계정 번호 수집

# 유니크 계정 번호를 가지고 있는 데이터프레임 생성
consolidated_df = pd.DataFrame(sorted(all_ledger_accounts), columns=['Ledger account']).set_index('Ledger account')

# 각 시트 처리 및 합산
for sheet_name in sheets:
    df = pd.read_excel(input_file_path, sheet_name=sheet_name)
    df.set_index(df.columns[0], inplace=True)  # A열을 인덱스로 설정
    # C열(2번째 열)과 F열(5번째 열)의 합산하여 'Closing Balance' 계산
    df['Closing Balance_' + sheet_name] = pd.to_numeric(df.iloc[:, 2], errors='coerce') + pd.to_numeric(df.iloc[:, 5], errors='coerce')
    consolidated_df = consolidated_df.join(df[['Closing Balance_' + sheet_name]], how='left')

# 누락된 값(NA)을 0으로 대체
consolidated_df.fillna(0, inplace=True)

# 결과를 새 파일에 저장
if not os.path.exists(output_file_path):
    with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
        consolidated_df.to_excel(writer, sheet_name='Alex US Cons')
else:
    book = load_workbook(output_file_path)
    with pd.ExcelWriter(output_file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        writer.book = book
        consolidated_df.to_excel(writer, sheet_name='Alex US Cons')

print("작업 완료!")
