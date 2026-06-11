import pandas as pd
import sys

sys.stdout.reconfigure(encoding='utf-8')

print("Starting verification of HQ1_Q1_2026_form.xlsx (allocated HQQU values)...")

# Load generated file
try:
    df_gen = pd.read_excel(r"D:\bv_hqqu_dashboard\HQ1_Q1_2026_form.xlsx", sheet_name='Sheet1')
except Exception as e:
    print(f"Error loading file: {e}")
    sys.exit(1)

print(f"Generated Sheet1 rows: {len(df_gen)}")
if len(df_gen) != 80:
    print(f"Warning: Expected 80 rows, got {len(df_gen)}")

# Load original raw data to compare
df_raw = pd.read_excel(r"D:\bv_hqqu_dashboard\HQ1_Q1_2026.xlsx")
df_raw = df_raw[df_raw['company_name_sun'].notna()]
df_raw['company_name_sun'] = df_raw['company_name_sun'].astype(str).str.strip()
df_raw = df_raw[~df_raw['company_name_sun'].str.startswith('Applied filters')]
df_raw = df_raw[df_raw['company_name_sun'] != 'company_name_sun']

# Align both by index
df_gen = df_gen.set_index('CTTV')
df_raw_aligned = df_raw.set_index('company_name_sun')

# Check that TONG = sum(LOBs)
sum_lobs = df_gen[['Y TE', 'DU_LICH', 'CON_NGUOI', 'XCG']].sum(axis=1)
diff_tong = (df_gen['TONG'] - sum_lobs).abs()
print(f"Max difference between TONG and sum of LOBs: {diff_tong.max()}")
if diff_tong.max() > 1e-2:
    print("Error: TONG column does not equal the sum of LOB columns!")
    sys.exit(1)

# Check TONG matches HQQU column from raw
diff_raw_hqqu = (df_gen['TONG'] - df_raw_aligned['HQQU']).abs().max()
print(f"Max difference with raw HQQU: {diff_raw_hqqu}")
if diff_raw_hqqu > 1e-2:
    print("Error: Total HQQU does not match raw HQQU!")
    sys.exit(1)

# Verify allocated math for CON_NGUOI, DU_LICH, XCG
for name, r_gen in df_gen.iterrows():
    r_raw = df_raw_aligned.loc[name]
    
    dt_con_nguoi = float(r_raw['DT_Con người']) if not pd.isna(r_raw['DT_Con người']) else 0.0
    dt_du_lich = float(r_raw['DT_Du lịch']) if not pd.isna(r_raw['DT_Du lịch']) else 0.0
    dt_xe_co_gioi = float(r_raw['DT_Xe Cơ Giới']) if not pd.isna(r_raw['DT_Xe Cơ Giới']) else 0.0
    dt_y_te = float(r_raw['DT_Y tế']) if not pd.isna(r_raw['DT_Y tế']) else 0.0
    
    bt_con_nguoi = float(r_raw['BT_Con Người']) if not pd.isna(r_raw['BT_Con Người']) else 0.0
    bt_du_lich = float(r_raw['BT_Du Lịch']) if not pd.isna(r_raw['BT_Du Lịch']) else 0.0
    bt_xe_co_gioi = float(r_raw['BT_Xe Cơ Giới']) if not pd.isna(r_raw['BT_Xe Cơ Giới']) else 0.0
    bt_y_te = float(r_raw['BT_Y tế']) if not pd.isna(r_raw['BT_Y tế']) else 0.0
    
    dt_thuan = float(r_raw['Doanh thu thuần']) if not pd.isna(r_raw['Doanh thu thuần']) else 0.0
    bt_trach_nhiem = float(r_raw['Bồi thường thuộc trách nghiệm']) if not pd.isna(r_raw['Bồi thường thuộc trách nghiệm']) else 0.0
    du_phong_phi = float(r_raw['Dự phòng phí']) if not pd.isna(r_raw['Dự phòng phí']) else 0.0
    chenh_lech_dpbt = float(r_raw['Chênh lệch Dự phòng bồi thường']) if not pd.isna(r_raw['Chênh lệch Dự phòng bồi thường']) else 0.0
    cpdm = float(r_raw['CPDM_HQ1']) if not pd.isna(r_raw['CPDM_HQ1']) else 0.0
    hqqu_total = float(r_raw['HQQU']) if not pd.isna(r_raw['HQQU']) else 0.0
    
    # Calculate allocated HQQU for CON_NGUOI
    con_nguoi_dp = du_phong_phi * (dt_con_nguoi / dt_thuan) if dt_thuan > 0 else 0.0
    con_nguoi_cp = cpdm * (dt_con_nguoi / dt_thuan) if dt_thuan > 0 else 0.0
    con_nguoi_cl = chenh_lech_dpbt * (bt_con_nguoi / bt_trach_nhiem) if bt_trach_nhiem > 0 else (chenh_lech_dpbt * (dt_con_nguoi / dt_thuan) if dt_thuan > 0 else 0.0)
    con_nguoi_hqqu = dt_con_nguoi - bt_con_nguoi - con_nguoi_dp - con_nguoi_cl - con_nguoi_cp
    
    # Calculate allocated HQQU for DU_LICH
    du_lich_dp = du_phong_phi * (dt_du_lich / dt_thuan) if dt_thuan > 0 else 0.0
    du_lich_cp = cpdm * (dt_du_lich / dt_thuan) if dt_thuan > 0 else 0.0
    du_lich_cl = chenh_lech_dpbt * (bt_du_lich / bt_trach_nhiem) if bt_trach_nhiem > 0 else (chenh_lech_dpbt * (dt_du_lich / dt_thuan) if dt_thuan > 0 else 0.0)
    du_lich_hqqu = dt_du_lich - bt_du_lich - du_lich_dp - du_lich_cl - du_lich_cp
    
    # Calculate allocated HQQU for XCG
    xcg_dp = du_phong_phi * (dt_xe_co_gioi / dt_thuan) if dt_thuan > 0 else 0.0
    xcg_cp = cpdm * (dt_xe_co_gioi / dt_thuan) if dt_thuan > 0 else 0.0
    xcg_cl = chenh_lech_dpbt * (bt_xe_co_gioi / bt_trach_nhiem) if bt_trach_nhiem > 0 else (chenh_lech_dpbt * (dt_xe_co_gioi / dt_thuan) if dt_thuan > 0 else 0.0)
    xcg_hqqu = dt_xe_co_gioi - bt_xe_co_gioi - xcg_dp - xcg_cl - xcg_cp

    # Check matches
    if abs(r_gen['CON_NGUOI'] - con_nguoi_hqqu) > 1e-2:
        print(f"Error for branch {name}: CON_NGUOI hqqu mismatch. Generated: {r_gen['CON_NGUOI']}, Expected: {con_nguoi_hqqu}")
        sys.exit(1)
        
    if abs(r_gen['DU_LICH'] - du_lich_hqqu) > 1e-2:
        print(f"Error for branch {name}: DU_LICH hqqu mismatch. Generated: {r_gen['DU_LICH']}, Expected: {du_lich_hqqu}")
        sys.exit(1)
        
    if abs(r_gen['XCG'] - xcg_hqqu) > 1e-2:
        print(f"Error for branch {name}: XCG hqqu mismatch. Generated: {r_gen['XCG']}, Expected: {xcg_hqqu}")
        sys.exit(1)

# Check seasonality groups
groups = df_gen['Phân nhóm'].unique()
print("Seasonality groups found:", groups)
if '-' in groups or None in groups or pd.isna(groups).any():
    print("Warning: Some rows have unclassified seasonality groups ('-')")

print("All checks passed! The file contains correct allocated performance numbers.")
