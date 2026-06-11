import pandas as pd
import sys

sys.stdout.reconfigure(encoding='utf-8')

print("Starting verification of HQ1_Q1_2026_form_filled.xlsx...")

# Load generated file
try:
    xl = pd.ExcelFile(r"D:\bv_hqqu_dashboard\HQ1_Q1_2026_form_filled.xlsx")
except Exception as e:
    print(f"Error loading file: {e}")
    sys.exit(1)

print("Sheet names:", xl.sheet_names)
if 'Sheet1' not in xl.sheet_names or 'Bồi thường' not in xl.sheet_names:
    print("Error: Missing expected sheets.")
    sys.exit(1)

df_rev = xl.parse('Sheet1')
df_clm = xl.parse('Bồi thường')

print(f"Revenue sheet rows: {len(df_rev)}")
print(f"Claims sheet rows: {len(df_clm)}")

# Check 80 rows
if len(df_rev) != 80:
    print(f"Warning: Expected 80 rows in Revenue, got {len(df_rev)}")
if len(df_clm) != 80:
    print(f"Warning: Expected 80 rows in Claims, got {len(df_clm)}")

# Load original raw data to compare
df_raw = pd.read_excel(r"D:\bv_hqqu_dashboard\HQ1_Q1_2026.xlsx")
df_raw = df_raw[df_raw['company_name_sun'].notna()]
df_raw['company_name_sun'] = df_raw['company_name_sun'].astype(str).str.strip()
df_raw = df_raw[~df_raw['company_name_sun'].str.startswith('Applied filters')]
df_raw = df_raw[df_raw['company_name_sun'] != 'company_name_sun']

print(f"Raw 2026 data rows: {len(df_raw)}")

# Check mapping of LOBs
lob_cols = ['Y TE', 'DU_LICH', 'CON_NGUOI', 'XCG']

for sheet_name, df, prefix in [('Sheet1', df_rev, 'DT_'), ('Bồi thường', df_clm, 'BT_')]:
    print(f"\nVerifying sheet: {sheet_name}")
    # Align by unit name
    df = df.set_index('CTTV')
    
    # Check that TONG = sum(LOBs)
    sum_lobs = df[lob_cols].sum(axis=1)
    diff_tong = (df['TONG'] - sum_lobs).abs()
    max_diff_tong = diff_tong.max()
    print(f"Max difference between TONG and sum of LOBs: {max_diff_tong}")
    if max_diff_tong > 1e-2:
        print("Error: TONG does not match sum of LOBs!")
        sys.exit(1)
        
    # Check against raw file
    raw_aligned = df_raw.set_index('company_name_sun')
    
    # Check TONG matches
    raw_col_total = 'Tổng Doanh thu' if prefix == 'DT_' else 'Tổng Bồi thường'
    diff_raw_total = (df['TONG'] - raw_aligned[raw_col_total]).abs()
    print(f"Max difference with raw total column: {diff_raw_total.max()}")
    
    # Check LOB values
    for col in lob_cols:
        raw_col = prefix + col
        # Normalize casing / spacing for matching raw columns
        # Map: Y TE -> Y tế, DU_LICH -> Du lịch / Du Lịch, CON_NGUOI -> Con người / Con Người, XCG -> Xe Cơ Giới / Xe cơ giới
        if prefix == 'DT_':
            map_name = {'Y TE': 'DT_Y tế', 'DU_LICH': 'DT_Du lịch', 'CON_NGUOI': 'DT_Con người', 'XCG': 'DT_Xe Cơ Giới'}
        else:
            map_name = {'Y TE': 'BT_Y tế', 'DU_LICH': 'BT_Du Lịch', 'CON_NGUOI': 'BT_Con Người', 'XCG': 'BT_Xe Cơ Giới'}
        
        raw_val = raw_aligned[map_name[col]].fillna(0)
        diff_val = (df[col] - raw_val).abs().max()
        print(f"  Column {col} max difference: {diff_val}")
        if diff_val > 1e-2:
            print(f"Error: Mismatch in column {col}!")
            sys.exit(1)

    # Check seasonality groups
    groups = df['Phân nhóm'].unique()
    print("Seasonality groups found:", groups)
    if '-' in groups or None in groups or pd.isna(groups).any():
        print("Warning: Some rows have unclassified seasonality groups ('-')")

print("\nVerification completed successfully!")
