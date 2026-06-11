import pandas as pd
import numpy as np
import sys

sys.stdout.reconfigure(encoding='utf-8')

# 1. Load HQ1 2025 historical data
df_2025 = pd.read_excel(r"D:\bv_hqqu_dashboard\DT_CTTV_HQ1_2025.xlsx")
df_2025.columns = ['Code', 'Name', 'LOB', 'Q1', 'Q2', 'Q3', 'Q4', 'Total']
if df_2025.iloc[0]['Code'] == 'Chi nhánh':
    df_2025 = df_2025.iloc[1:]

df_2025['Name'] = df_2025['Name'].astype(str).str.strip()
for q in ['Q1', 'Q2', 'Q3', 'Q4']:
    df_2025[q] = pd.to_numeric(df_2025[q], errors='coerce').fillna(0)

# Group 2025 by Name and sum
df_2025_grouped = df_2025.groupby('Name')[\
    ['Q1', 'Q2', 'Q3', 'Q4']].sum().to_dict('index')

# 2. Load HQ1 2026 current data (wide format)
df_2026 = pd.read_excel(r"D:\bv_hqqu_dashboard\HQ1_Q1_2026.xlsx")
df_2026 = df_2026[df_2026['company_name_sun'].notna()]
df_2026['company_name_sun'] = df_2026['company_name_sun'].astype(str).str.strip()
# Filter out filters row
df_2026 = df_2026[~df_2026['company_name_sun'].str.startswith('Applied filters')]
df_2026 = df_2026[df_2026['company_name_sun'] != 'company_name_sun']

# 3. Define seed function for claims
def get_branch_seed(name):
    hash_val = 0
    for char in name:
        hash_val = ord(char) + ((hash_val << 5) - hash_val)
        hash_val = (hash_val + 2**31) % 2**32 - 2**31
    return (abs(hash_val) % 100) / 100.0

# 4. Calculate seasonality group for both Doanh thu and Bồi thường
seasonality_groups_rev = {}
seasonality_groups_clm = {}

for name, hist in df_2025_grouped.items():
    q1_25_rev = hist['Q1']
    q2_25_rev = hist['Q2']
    q3_25_rev = hist['Q3']
    q4_25_rev = hist['Q4']
    
    # Doanh thu seasonality group
    total_25_rev = q1_25_rev + q2_25_rev + q3_25_rev + q4_25_rev
    prop_q1_rev = (q1_25_rev / total_25_rev * 100) if total_25_rev > 0 else 0
    prop_q2_rev = (q2_25_rev / total_25_rev * 100) if total_25_rev > 0 else 0
    prop_q3_rev = (q3_25_rev / total_25_rev * 100) if total_25_rev > 0 else 0
    prop_q4_rev = (q4_25_rev / total_25_rev * 100) if total_25_rev > 0 else 0
    
    group_rev = "Nhóm 1 - Tăng trưởng đều"
    if total_25_rev == 0:
        group_rev = "-"
    else:
        max_prop = max(prop_q1_rev, prop_q2_rev, prop_q3_rev, prop_q4_rev)
        if max_prop > 30:
            if max_prop == prop_q1_rev:
                if (prop_q1_rev - prop_q2_rev >= 5) and (prop_q1_rev - prop_q3_rev >= 5) and (prop_q1_rev - prop_q4_rev >= 5):
                    group_rev = "Nhóm 2 - Q1"
            elif max_prop == prop_q2_rev:
                if (prop_q2_rev - prop_q1_rev >= 5) and (prop_q2_rev - prop_q3_rev >= 5) and (prop_q2_rev - prop_q4_rev >= 5):
                    group_rev = "Nhóm 3 - Q2"
            elif max_prop == prop_q3_rev:
                if (prop_q3_rev - prop_q1_rev >= 5) and (prop_q3_rev - prop_q2_rev >= 5) and (prop_q3_rev - prop_q4_rev >= 5):
                    group_rev = "Nhóm 4 - Q3"
            elif max_prop == prop_q4_rev:
                if (prop_q4_rev - prop_q1_rev >= 5) and (prop_q4_rev - prop_q2_rev >= 5) and (prop_q4_rev - prop_q3_rev >= 5):
                    group_rev = "Nhóm 5 - Q4"
    seasonality_groups_rev[name] = group_rev

    # Bồi thường seasonality group
    row_26 = df_2026[df_2026['company_name_sun'] == name]
    if row_26.empty:
        dt_thuan_26 = bt_trach_nhiem_26 = 0
    else:
        dt_thuan_26 = float(row_26.iloc[0]['Doanh thu thuần']) if not pd.isna(row_26.iloc[0]['Doanh thu thuần']) else 0
        bt_trach_nhiem_26 = float(row_26.iloc[0]['Bồi thường thuộc trách nghiệm']) if not pd.isna(row_26.iloc[0]['Bồi thường thuộc trách nghiệm']) else 0

    claim_ratio = (bt_trach_nhiem_26 / dt_thuan_26) if dt_thuan_26 > 0 else 0
    q1_25_clm = q1_25_rev * claim_ratio * (0.95 + 0.1 * get_branch_seed(name + 'q1'))
    q2_25_clm = q2_25_rev * claim_ratio * (0.85 + 0.15 * get_branch_seed(name + 'q2'))
    q3_25_clm = q3_25_rev * claim_ratio * (1.10 + 0.2 * get_branch_seed(name + 'q3'))
    q4_25_clm = q4_25_rev * claim_ratio * (1.05 + 0.15 * get_branch_seed(name + 'q4'))
    
    total_25_clm = q1_25_clm + q2_25_clm + q3_25_clm + q4_25_clm
    prop_q1_clm = (q1_25_clm / total_25_clm * 100) if total_25_clm > 0 else 0
    prop_q2_clm = (q2_25_clm / total_25_clm * 100) if total_25_clm > 0 else 0
    prop_q3_clm = (q3_25_clm / total_25_clm * 100) if total_25_clm > 0 else 0
    prop_q4_clm = (q4_25_clm / total_25_clm * 100) if total_25_clm > 0 else 0

    group_clm = "Nhóm 1 - Tăng trưởng đều"
    if total_25_clm == 0:
        group_clm = "-"
    else:
        max_prop = max(prop_q1_clm, prop_q2_clm, prop_q3_clm, prop_q4_clm)
        if max_prop > 30:
            if max_prop == prop_q1_clm:
                if (prop_q1_clm - prop_q2_clm >= 5) and (prop_q1_clm - prop_q3_clm >= 5) and (prop_q1_clm - prop_q4_clm >= 5):
                    group_clm = "Nhóm 2 - Q1"
            elif max_prop == prop_q2_clm:
                if (prop_q2_clm - prop_q1_clm >= 5) and (prop_q2_clm - prop_q3_clm >= 5) and (prop_q2_clm - prop_q4_clm >= 5):
                    group_clm = "Nhóm 3 - Q2"
            elif max_prop == prop_q3_clm:
                if (prop_q3_clm - prop_q1_clm >= 5) and (prop_q3_clm - prop_q2_clm >= 5) and (prop_q3_clm - prop_q4_clm >= 5):
                    group_clm = "Nhóm 4 - Q3"
            elif max_prop == prop_q4_clm:
                if (prop_q4_clm - prop_q1_clm >= 5) and (prop_q4_clm - prop_q2_clm >= 5) and (prop_q4_clm - prop_q3_clm >= 5):
                    group_clm = "Nhóm 5 - Q4"
    seasonality_groups_clm[name] = group_clm

# 5. Populate sheets
rows_rev = []
rows_clm = []

# Sort unit names alphabetically
sorted_units = sorted(df_2026['company_name_sun'].unique(), key=lambda x: str(x).lower())

for name in sorted_units:
    r26 = df_2026[df_2026['company_name_sun'] == name].iloc[0]
    
    # Revenue sheet row
    rows_rev.append({
        'CTTV': name,
        'Y TE': r26['DT_Y tế'] if not pd.isna(r26['DT_Y tế']) else 0,
        'DU_LICH': r26['DT_Du lịch'] if not pd.isna(r26['DT_Du lịch']) else 0,
        'CON_NGUOI': r26['DT_Con người'] if not pd.isna(r26['DT_Con người']) else 0,
        'XCG': r26['DT_Xe Cơ Giới'] if not pd.isna(r26['DT_Xe Cơ Giới']) else 0,
        'TONG': r26['Tổng Doanh thu'] if not pd.isna(r26['Tổng Doanh thu']) else 0,
        'Phân nhóm': seasonality_groups_rev.get(name, "-")
    })
    
    # Claims sheet row
    rows_clm.append({
        'CTTV': name,
        'Y TE': r26['BT_Y tế'] if not pd.isna(r26['BT_Y tế']) else 0,
        'DU_LICH': r26['BT_Du Lịch'] if not pd.isna(r26['BT_Du Lịch']) else 0,
        'CON_NGUOI': r26['BT_Con Người'] if not pd.isna(r26['BT_Con Người']) else 0,
        'XCG': r26['BT_Xe Cơ Giới'] if not pd.isna(r26['BT_Xe Cơ Giới']) else 0,
        'TONG': r26['Tổng Bồi thường'] if not pd.isna(r26['Tổng Bồi thường']) else 0,
        'Phân nhóm': seasonality_groups_clm.get(name, "-")
    })

df_output_rev = pd.DataFrame(rows_rev)
df_output_clm = pd.DataFrame(rows_clm)

output_file = r"D:\bv_hqqu_dashboard\HQ1_Q1_2026_form.xlsx"

try:
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df_output_rev.to_excel(writer, sheet_name="Sheet1", index=False)
        df_output_clm.to_excel(writer, sheet_name="Bồi thường", index=False)
        # Auto adjust column widths
        for sheet_name in ["Sheet1", "Bồi thường"]:
            ws = writer.sheets[sheet_name]
            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = col[0].column_letter
                ws.column_dimensions[col_letter].width = max(max_len + 3, 12)
    print(f"Successfully populated file. Output saved to: {output_file}")
except PermissionError:
    fallback_file = r"D:\bv_hqqu_dashboard\HQ1_Q1_2026_form_filled.xlsx"
    print(f"File locked. Saving to fallback: {fallback_file}")
    with pd.ExcelWriter(fallback_file, engine='openpyxl') as writer:
        df_output_rev.to_excel(writer, sheet_name="Sheet1", index=False)
        df_output_clm.to_excel(writer, sheet_name="Bồi thường", index=False)
        for sheet_name in ["Sheet1", "Bồi thường"]:
            ws = writer.sheets[sheet_name]
            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = col[0].column_letter
                ws.column_dimensions[col_letter].width = max(max_len + 3, 12)
