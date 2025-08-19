import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import os

def load_data():
    print('[DEBUG] Loading Excel files...')
    df_putaway = pd.read_excel('input/CEVA putaway records.xlsx', engine='openpyxl')
    df_test = pd.read_excel('input/Test Records.xlsx', engine='openpyxl')
    print('[DEBUG] Columns in CEVA putaway records.xlsx:')
    print(df_putaway.columns.tolist())
    print('[DEBUG] Columns in Test Records.xlsx:')
    print(df_test.columns.tolist())
    return df_putaway, df_test

def validate_row(row):
    # Debug for each row validation
    if pd.isna(row['Part Number_putaway']):
        print(f"[DEBUG] SN {row['Serial Number']} missing in putaway records.")
        return 'NOT OK: Serial Number missing in putaway records'
    part_match = row['Part Number_test'] == row['Part Number_putaway']
    if str(row['Test Result']).strip().upper() == 'PASS':
        rack_bin_ok = str(row['Rack Bin']).strip().upper().endswith('P')
        result = part_match and rack_bin_ok
        if not result:
            print(f"[DEBUG] SN {row['Serial Number']} failed PASS check: part_match={part_match}, rack_bin_ok={rack_bin_ok}")
        reason = 'OK' if result else 'NOT OK: Part mismatch or Rack Bin not ending with P'
    elif str(row['Test Result']).strip().upper() in ['FAIL', 'MAYBE']:
        # Exception: if Rack Bin is 'Testwip' (case-insensitive), mark as OK
        if str(row['Rack Bin']).strip().upper() == 'TESTWIP':
            result = True
            rack_bin_ok = True
            print(f"[DEBUG] SN {row['Serial Number']} is exception: Rack Bin is Testwip.")
            reason = 'OK (Exception: Rack Bin is Testwip)'
        else:
            rack_bin_ok = not str(row['Rack Bin']).strip().upper().endswith('P')
            result = part_match and rack_bin_ok
            if not result:
                print(f"[DEBUG] SN {row['Serial Number']} failed FAIL/MAYBE check: part_match={part_match}, rack_bin_ok={rack_bin_ok}")
            reason = 'OK' if result else 'NOT OK: Part mismatch or Rack Bin ends with P'
    else:
        print(f"[DEBUG] SN {row['Serial Number']} has unknown test result: {row['Test Result']}")
        reason = 'NOT OK: Unknown Test Result'
    return reason

def check_duplicates(df_test, df_putaway):
    print('[DEBUG] Checking for duplicate Serial Numbers...')
    dup_test = df_test['Serial Number'].duplicated()
    dup_putaway = df_putaway['Serial Number'].duplicated()
    
    test_duplicates = df_test[dup_test]['Serial Number'].tolist() if dup_test.any() else []
    putaway_duplicates = df_putaway[dup_putaway]['Serial Number'].tolist() if dup_putaway.any() else []
    
    dup_test_msg = f'Duplicate Serial Numbers found in Test Records: {", ".join(map(str, test_duplicates))}' if test_duplicates else 'No duplicate Serial Numbers in Test Records'
    dup_putaway_msg = f'Duplicate Serial Numbers found in Putaway Records: {", ".join(map(str, putaway_duplicates))}' if putaway_duplicates else 'No duplicate Serial Numbers in Putaway Records'
    
    print(f'[DEBUG] {dup_test_msg}')
    print(f'[DEBUG] {dup_putaway_msg}')
    return dup_test_msg, dup_putaway_msg

def format_dates(df_merged):
    print('[DEBUG] Formatting date columns...')
    for col in ['Testing Date_test', 'Testing Date_putaway']:
        if col in df_merged.columns:
            df_merged[col] = pd.to_datetime(df_merged[col], errors='coerce').dt.date
    return df_merged

def save_with_autofit(df_merged, output_path):
    print(f'[DEBUG] Saving DataFrame to {output_path} and autofitting columns...')
    df_merged.to_excel(output_path, index=False)
    wb = load_workbook(output_path)
    ws = wb.active
    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        ws.column_dimensions[col_letter].width = max_length + 2
    wb.save(output_path)
    print('[DEBUG] File saved and columns autofitted.')

def main():
    print('[DEBUG] Starting validation process...')
    df_putaway, df_test = load_data()
    print(f'[DEBUG] Loaded {len(df_putaway)} rows from putaway records and {len(df_test)} rows from test records.')
    print('[DEBUG] Merging test and putaway records on Serial Number...')
    df_merged = pd.merge(df_test, df_putaway, on='Serial Number', how='left', suffixes=('_test', '_putaway'))
    print(f'[DEBUG] Merged DataFrame has {len(df_merged)} rows.')
    print('[DEBUG] First 10 rows of merged DataFrame:')
    print(df_merged.head(10))
    print('[DEBUG] Applying validation logic to each row...')
    df_merged['Validation Result'] = df_merged.apply(validate_row, axis=1)
    print('[DEBUG] Validation logic applied. Here are the first 10 validation results:')
    print(df_merged['Validation Result'].head(10))
    print('[DEBUG] Formatting date columns...')
    df_merged = format_dates(df_merged)
    print('[DEBUG] Date columns formatted.')
    print('[DEBUG] Checking for duplicate Serial Numbers in both files...')
    dup_test_msg, dup_putaway_msg = check_duplicates(df_test, df_putaway)
    print(f'[DEBUG] Duplicate check results: {dup_test_msg}, {dup_putaway_msg}')
    
    # Add an empty row for separation before duplicates summary
    summary = {col: '' for col in df_merged.columns}
    summary['Serial Number'] = '--- DUPLICATE ANALYSIS ---'
    summary['Validation Result'] = 'Summary of duplicate serial numbers found'
    
    summary_test = {col: '' for col in df_merged.columns}
    summary_test['Serial Number'] = 'TEST RECORDS DUPLICATES'
    summary_test['Validation Result'] = dup_test_msg
    
    summary_putaway = {col: '' for col in df_merged.columns}
    summary_putaway['Serial Number'] = 'PUTAWAY RECORDS DUPLICATES'
    summary_putaway['Validation Result'] = dup_putaway_msg
    
    df_merged = pd.concat([df_merged, pd.DataFrame([summary, summary_test, summary_putaway])], ignore_index=True)
    output_path = 'output/merged_validation.xlsx'
    print(f'[DEBUG] Saving results to {output_path}...')
    save_with_autofit(df_merged, output_path)
    print("[DEBUG] Validation complete. Results saved to output/merged_validation.xlsx with autofit columns.")
    print('[DEBUG] Opening the Excel file for review...')
    os.system('start excel.exe output\\merged_validation.xlsx')

if __name__ == "__main__":
    main()
