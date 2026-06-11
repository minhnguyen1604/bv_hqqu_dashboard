import pandas as pd
import numpy as np
import sys

sys.stdout.reconfigure(encoding='utf-8')

# 1. Load HQ1 2025 historical data (using 'Tổng doanh thu quý' sheet)
df_2025 = pd.read_excel(r"D:\bv_hqqu_dashboard\DT_CTTV_HQ1_2025.xlsx", sheet_name='Tổng doanh thu quý')
df_2025.columns = ['Name', 'Q1', 'Q2', 'Q3', 'Q4']
df_2025['Name'] = df_2025['Name'].astype(str).str.strip()

for q in ['Q1', 'Q2', 'Q3', 'Q4']:
    df_2025[q] = pd.to_numeric(df_2025[q], errors='coerce').fillna(0)

# Group by Name and sum (just in case there are duplicates, though it has 80 unique rows)
df_2025_grouped = df_2025.groupby('Name')[['Q1', 'Q2', 'Q3', 'Q4']].sum().to_dict('index')

# Calculate seasonality group for Doanh thu
seasonality_groups_rev = {}

for name, hist in df_2025_grouped.items():
    q1_25 = hist['Q1']
    q2_25 = hist['Q2']
    q3_25 = hist['Q3']
    q4_25 = hist['Q4']
    
    total_25 = q1_25 + q2_25 + q3_25 + q4_25
    prop_q1 = (q1_25 / total_25 * 100) if total_25 > 0 else 0
    prop_q2 = (q2_25 / total_25 * 100) if total_25 > 0 else 0
    prop_q3 = (q3_25 / total_25 * 100) if total_25 > 0 else 0
    prop_q4 = (q4_25 / total_25 * 100) if total_25 > 0 else 0
    
    group = "Nhóm 1 - Tăng trưởng đều"
    if total_25 == 0:
        group = "-"
    else:
        max_prop = max(prop_q1, prop_q2, prop_q3, prop_q4)
        if max_prop > 30:
            if max_prop == prop_q1:
                if (prop_q1 - prop_q2 >= 5) and (prop_q1 - prop_q3 >= 5) and (prop_q1 - prop_q4 >= 5):
                    group = "Nhóm 2 - Q1"
            elif max_prop == prop_q2:
                if (prop_q2 - prop_q1 >= 5) and (prop_q2 - prop_q3 >= 5) and (prop_q2 - prop_q4 >= 5):
                    group = "Nhóm 3 - Q2"
            elif max_prop == prop_q3:
                if (prop_q3 - prop_q1 >= 5) and (prop_q3 - prop_q2 >= 5) and (prop_q3 - prop_q4 >= 5):
                    group = "Nhóm 4 - Q3"
            elif max_prop == prop_q4:
                if (prop_q4 - prop_q1 >= 5) and (prop_q4 - prop_q2 >= 5) and (prop_q4 - prop_q3 >= 5):
                    group = "Nhóm 5 - Q4"
    seasonality_groups_rev[name] = group

# 2. Load HQ1 2026 current data (wide format)
df_2026 = pd.read_excel(r"D:\bv_hqqu_dashboard\HQ1_Q1_2026.xlsx")
df_2026 = df_2026[df_2026['company_name_sun'].notna()]
df_2026['company_name_sun'] = df_2026['company_name_sun'].astype(str).str.strip()
df_2026 = df_2026[~df_2026['company_name_sun'].str.startswith('Applied filters')]
df_2026 = df_2026[df_2026['company_name_sun'] != 'company_name_sun']

# 3. Calculate allocated LOB-level HQQU
rows_hqqu = []

# Sort unit names alphabetically
sorted_units = sorted(df_2026['company_name_sun'].unique(), key=lambda x: str(x).lower())

for name in sorted_units:
    r26 = df_2026[df_2026['company_name_sun'] == name].iloc[0]
    
    # Extract values and handle NaN
    dt_con_nguoi = float(r26['DT_Con người']) if not pd.isna(r26['DT_Con người']) else 0.0
    dt_du_lich = float(r26['DT_Du lịch']) if not pd.isna(r26['DT_Du lịch']) else 0.0
    dt_xe_co_gioi = float(r26['DT_Xe Cơ Giới']) if not pd.isna(r26['DT_Xe Cơ Giới']) else 0.0
    dt_y_te = float(r26['DT_Y tế']) if not pd.isna(r26['DT_Y tế']) else 0.0
    
    bt_con_nguoi = float(r26['BT_Con Người']) if not pd.isna(r26['BT_Con Người']) else 0.0
    bt_du_lich = float(r26['BT_Du Lịch']) if not pd.isna(r26['BT_Du Lịch']) else 0.0
    bt_xe_co_gioi = float(r26['BT_Xe Cơ Giới']) if not pd.isna(r26['BT_Xe Cơ Giới']) else 0.0
    bt_y_te = float(r26['BT_Y tế']) if not pd.isna(r26['BT_Y tế']) else 0.0
    
    dt_thuan = float(r26['Doanh thu thuần']) if not pd.isna(r26['Doanh thu thuần']) else 0.0
    bt_trach_nhiem = float(r26['Bồi thường thuộc trách nghiệm']) if not pd.isna(r26['Bồi thường thuộc trách nghiệm']) else 0.0
    du_phong_phi = float(r26['Dự phòng phí']) if not pd.isna(r26['Dự phòng phí']) else 0.0
    chenh_lech_dpbt = float(r26['Chênh lệch Dự phòng bồi thường']) if not pd.isna(r26['Chênh lệch Dự phòng bồi thường']) else 0.0
    cpdm = float(r26['CPDM_HQ1']) if not pd.isna(r26['CPDM_HQ1']) else 0.0
    hqqu_total = float(r26['HQQU']) if not pd.isna(r26['HQQU']) else 0.0
    
    # Allocated calculation lists: index 0: Con nguoi, index 1: Du lich, index 2: Xe co gioi, index 3: Y te (last LOB)
    lobs = [
        {'name': 'Con người', 'rev': dt_con_nguoi, 'clm': bt_con_nguoi},
        {'name': 'Du lịch', 'rev': dt_du_lich, 'clm': bt_du_lich},
        {'name': 'Xe Cơ Giới', 'rev': dt_xe_co_gioi, 'clm': bt_xe_co_gioi},
        {'name': 'Y tế', 'rev': dt_y_te, 'clm': bt_y_te}
    ]
    
    sum_dt_thuan = 0.0
    sum_bt_trach_nhiem = 0.0
    sum_du_phong = 0.0
    sum_chenh_lech = 0.0
    sum_cpdm = 0.0
    sum_hqqu = 0.0
    
    hqqu_by_lob = {}
    
    for idx, lob in enumerate(lobs):
        if idx < len(lobs) - 1:
            dt_thuan_lob = lob['rev']
            bt_trach_nhiem_lob = lob['clm']
            
            allocated_dp = du_phong_phi * (lob['rev'] / dt_thuan) if dt_thuan > 0 else 0.0
            allocated_cp = cpdm * (lob['rev'] / dt_thuan) if dt_thuan > 0 else 0.0
            allocated_cl = 0.0
            if bt_trach_nhiem > 0:
                allocated_cl = chenh_lech_dpbt * (lob['clm'] / bt_trach_nhiem)
            else:
                allocated_cl = chenh_lech_dpbt * (lob['rev'] / dt_thuan) if dt_thuan > 0 else 0.0
                
            allocated_hqqu = dt_thuan_lob - bt_trach_nhiem_lob - allocated_dp - allocated_cl - allocated_cp
            
            sum_dt_thuan += dt_thuan_lob
            sum_bt_trach_nhiem += bt_trach_nhiem_lob
            sum_du_phong += allocated_dp
            sum_chenh_lech += allocated_cl;
            sum_cpdm += allocated_cp
            sum_hqqu += allocated_hqqu
            
            hqqu_by_lob[lob['name']] = allocated_hqqu
        else:
            # Last LOB (Y tế) gets the remainder
            allocated_hqqu = hqqu_total - sum_hqqu
            hqqu_by_lob[lob['name']] = allocated_hqqu

    # Append row
    rows_hqqu.append({
        'CTTV': name,
        'Y TE': hqqu_by_lob['Y tế'],
        'DU_LICH': hqqu_by_lob['Du lịch'],
        'CON_NGUOI': hqqu_by_lob['Con người'],
        'XCG': hqqu_by_lob['Xe Cơ Giới'],
        'TONG': hqqu_total,
        'Phân nhóm': seasonality_groups_rev.get(name, "-")
    })

df_output = pd.DataFrame(rows_hqqu)

# 4. Save to Excel
output_file = r"D:\bv_hqqu_dashboard\HQ1_Q1_2026_form.xlsx"

def save_workbook(path):
    with pd.ExcelWriter(path, engine='openpyxl') as writer:
        df_output.to_excel(writer, sheet_name="Sheet1", index=False)
        ws = writer.sheets["Sheet1"]
        # Auto adjust column widths
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = col[0].column_letter
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

try:
    save_workbook(output_file)
    print(f"Successfully populated file. Output saved to: {output_file}")
except PermissionError:
    fallback_file = r"D:\bv_hqqu_dashboard\HQ1_Q1_2026_form_filled.xlsx"
    print(f"File locked. Saving to fallback: {fallback_file}")
    save_workbook(fallback_file)
