import pandas as pd
import numpy as np
import os
import json
from datetime import datetime

def apply_zscore(df, uph_col):
    """ตัด outliers ด้วย Z-Score (±3 std)"""
    mean = df[uph_col].mean()
    std = df[uph_col].std()
    if std == 0:
        return df
    z_scores = (df[uph_col] - mean) / std
    filtered = df[(z_scores >= -3) & (z_scores <= 3)].copy()
    filtered['Outlier_Method'] = 'Z-Score ±3'
    return filtered

def apply_iqr(df, uph_col):
    """ตัด outliers ด้วย IQR"""
    Q1 = df[uph_col].quantile(0.25)
    Q3 = df[uph_col].quantile(0.75)
    IQR = Q3 - Q1
    lower = Q1 - 1.5 * IQR
    upper = Q3 + 1.5 * IQR
    filtered = df[(df[uph_col] >= lower) & (df[uph_col] <= upper)].copy()
    filtered['Outlier_Method'] = 'IQR'
    return filtered

def has_outlier(df, uph_col):
    """ตรวจสอบ outliers"""
    Q1 = df[uph_col].quantile(0.25)
    Q3 = df[uph_col].quantile(0.75)
    IQR = Q3 - Q1
    lower = Q1 - 1.5 * IQR
    upper = Q3 + 1.5 * IQR
    return ((df[uph_col] < lower) | (df[uph_col] > upper)).sum() > 0

def remove_outliers_auto(df_model, uph_col, max_iter=20):
    """ตัด outliers อัตโนมัติ"""
    df_model[uph_col] = pd.to_numeric(df_model[uph_col], errors='coerce')
    df_model = df_model.dropna(subset=[uph_col])

    if len(df_model) < 15:
        df_model['Outlier_Method'] = 'ไม่ตัด (ข้อมูลน้อย)'
        return df_model

    current_df = df_model.copy()
    for i in range(max_iter):
        z_df = apply_zscore(current_df, uph_col)
        if not has_outlier(z_df, uph_col):
            z_df['Outlier_Method'] = f'Z-Score Loop ×{i+1}'
            return z_df

        iqr_df = apply_iqr(z_df, uph_col)
        if not has_outlier(iqr_df, uph_col):
            iqr_df['Outlier_Method'] = f'IQR Loop ×{i+1}'
            return iqr_df
        current_df = iqr_df

    current_df['Outlier_Method'] = f'IQR-Z-Score Loop ×{max_iter}+'
    return current_df

def get_column_names(df):
    """หาชื่อคอลัมน์ที่ต้องการ"""
    col_map = {col.lower(): col for col in df.columns}
    
    uph_col = col_map.get('uph')
    if not uph_col:
        raise KeyError("ไม่พบคอลัมน์ UPH")
    
    model_col = col_map.get('machine model') or col_map.get('machine_model')
    if not model_col:
        raise KeyError("ไม่พบคอลัมน์ Machine Model")
    
    bom_col = col_map.get('bom_no') or col_map.get('bom no')
    if not bom_col:
        raise KeyError("ไม่พบคอลัมน์ bom_no")
    
    date_col = None
    for col_name in df.columns:
        if any(keyword in col_name.lower() for keyword in ['date', 'time', 'วัน', 'เวลา']):
            date_col = col_name
            break
    
    return uph_col, model_col, bom_col, date_col

def load_file(file_path):
    """อ่านไฟล์ตามประเภท"""
    if file_path.endswith('.xlsx'):
        return pd.read_excel(file_path, engine='openpyxl')
    elif file_path.endswith('.xls'):
        return pd.read_excel(file_path, engine='xlrd')
    elif file_path.endswith('.csv'):
        return pd.read_csv(file_path)
    elif file_path.endswith('.json'):
        with open(file_path, 'r', encoding='utf-8') as f:
            json_data = json.load(f)
        if isinstance(json_data, list):
            return pd.DataFrame(json_data)
        elif isinstance(json_data, dict):
            for key in ['data', 'results', 'items', 'records']:
                if key in json_data and isinstance(json_data[key], list):
                    return pd.DataFrame(json_data[key])
            return pd.DataFrame([json_data])
    else:
        return pd.read_excel(file_path, engine='openpyxl')

def remove_outliers(df):
    """ตัด outliers ตามกลุ่ม"""
    uph_col, model_col, bom_col, _ = get_column_names(df)
    result_dfs = []
    
    for (bom_no, machine_model), group_df in df.groupby([bom_col, model_col]):
        before_count = len(group_df)
        cleaned_group = remove_outliers_auto(group_df, uph_col)
        after_count = len(cleaned_group)
        cleaned_group['DataPoints_Before'] = before_count
        cleaned_group['DataPoints_After'] = after_count
        result_dfs.append(cleaned_group)
    
    return pd.concat(result_dfs, ignore_index=True)

def process_date_column(df):
    """ประมวลผลคอลัมน์วันที่"""
    _, _, _, date_col = get_column_names(df)
    
    if not date_col:
        print("ไม่พบคอลัมน์วันที่")
        return df
    
    print(f"ใช้คอลัมน์วันที่: {date_col}")
    df['date_time_start'] = pd.to_datetime(df[date_col], errors='coerce')
    df['date_time_start'] = df['date_time_start'].dt.strftime('%Y/%m/%d')
    
    invalid_dates = df['date_time_start'].isna().sum()
    if invalid_dates > 0:
        print(f"พบวันที่ที่แปลงไม่ได้: {invalid_dates} แถว")
        df = df.dropna(subset=['date_time_start'])
    
    return df

def get_date_range(df, start_date=None, end_date=None):
    """ได้ช่วงวันที่"""
    if start_date and end_date:
        return start_date, end_date
    
    max_date = df['date_time_start'].max()
    min_date = df['date_time_start'].min()
    print(f"ใช้ช่วงวันที่ทั้งหมด: {min_date} ถึง {max_date}")
    return min_date, max_date

def filter_by_date_range(df, start_date, end_date):
    """กรองข้อมูลตามช่วงวันที่"""
    filtered_df = df[df['date_time_start'].between(start_date, end_date)].copy()
    
    if len(filtered_df) == 0:
        raise Exception("ไม่พบข้อมูลในช่วงวันที่ที่เลือก")
    
    print(f"กรองข้อมูล: {len(filtered_df)}/{len(df)} แถว")
    return filtered_df

def calculate_group_average(df, start_date, end_date):
    """คำนวณค่าเฉลี่ยตามกลุ่ม"""
    uph_col, model_col, bom_col, _ = get_column_names(df)
    
    grouped = df.groupby([bom_col, model_col], as_index=False).agg({uph_col: 'mean'})
    grouped[uph_col] = grouped[uph_col].round(3)
    
    other_cols = ['operation', 'optn_code'] + (['DataPoints_Before', 'DataPoints_After'] if 'DataPoints_Before' in df.columns else [])
    if other_cols:
        firsts = df.groupby([bom_col, model_col], as_index=False)[other_cols].first()
        grouped = pd.merge(grouped, firsts, on=[bom_col, model_col], how='left')
    
    print(f"=== ค่าเฉลี่ย UPH ({start_date} ถึง {end_date}) ===")
    print(grouped)
    return grouped

def save_results(df_cleaned, grouped_average, start_date, end_date, output_dir):
    """บันทึกผลลัพธ์"""
    os.makedirs(output_dir, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    date_range = f"{start_date.replace('/', '')}_to_{end_date.replace('/', '')}"
    
    cleaned_file = os.path.join(output_dir, f"cleaned_data_{date_range}_{timestamp}.xlsx")
    average_file = os.path.join(output_dir, f"group_average_{date_range}_{timestamp}.xlsx")
    
    df_cleaned.to_excel(cleaned_file, index=False, engine='openpyxl')
    grouped_average.to_excel(average_file, index=False, engine='openpyxl')
    
    print(f"บันทึกไฟล์: {cleaned_file}")
    print(f"บันทึกไฟล์: {average_file}")
    
    return cleaned_file, average_file

def process_die_attack_data(file_path, start_date=None, end_date=None):
    """ประมวลผลข้อมูล Die Attack"""
    print("=== ประมวลผลข้อมูล Die Attack ===")
    
    df = load_file(file_path)
    print(f"ข้อมูลเริ่มต้น: {len(df)} แถว")
    
    df = process_date_column(df)
    
    if start_date and end_date:
        start_date = start_date.replace("-", "/")
        end_date = end_date.replace("-", "/")
    else:
        start_date, end_date = get_date_range(df)
    
    df_filtered = filter_by_date_range(df, start_date, end_date)
    
    print("ตัด outliers...")
    df_cleaned = remove_outliers(df_filtered)
    print(f"ข้อมูลหลังตัด outliers: {len(df_cleaned)} แถว")
    
    grouped_average = calculate_group_average(df_cleaned, start_date, end_date)
    
    return df_cleaned, grouped_average, start_date, end_date

def preview_date_range(file_path):
    """แสดงข้อมูลวันที่ในไฟล์"""
    try:
        df = load_file(file_path)
        print(f"ไฟล์มีข้อมูล: {len(df):,} แถว")
        
        date_cols = [col for col in df.columns 
                    if any(keyword in col.lower() for keyword in ['date', 'time', 'วัน', 'เวลา'])]
        
        if not date_cols:
            print("ไม่พบคอลัมน์วันที่")
            return None
        
        date_col = date_cols[0]
        df['temp_date'] = pd.to_datetime(df[date_col], errors='coerce')
        valid_dates = df.dropna(subset=['temp_date'])
        
        if len(valid_dates) == 0:
            print("ไม่มีข้อมูลวันที่ที่ถูกต้อง")
            return None
        
        min_date = valid_dates['temp_date'].min()
        max_date = valid_dates['temp_date'].max()
        
        print(f"วันที่: {min_date.strftime('%Y-%m-%d')} ถึง {max_date.strftime('%Y-%m-%d')}")
        print(f"ข้อมูลถูกต้อง: {len(valid_dates):,} แถว")
        
        return {
            'min_date': min_date.strftime('%Y-%m-%d'),
            'max_date': max_date.strftime('%Y-%m-%d'),
            'valid_records': len(valid_dates),
            'total_records': len(df)
        }
        
    except Exception as e:
        print(f"เกิดข้อผิดพลาด: {str(e)}")
        return None

def map_data(average_file):
    """Map ข้อมูลเพิ่มเติมจากไฟล์ Part bom pkg ในโฟลเดอร์ data_MAP"""
    print("=== Map ข้อมูลเพิ่มเติม ===")
    
    try:
        # โหลดไฟล์ average
        df_average = pd.read_excel(average_file, engine='openpyxl')
        print(f"📊 ข้อมูล average: {len(df_average)} แถว")

        # หาไฟล์ mapping
        current_dir = os.path.dirname(os.path.abspath(__file__))
        map_folder = os.path.join(current_dir, "..", "data_MAP")

        mapping_file = os.path.join(map_folder, "Part bom pkg.xlsx")
        mapping_file2 = os.path.join(map_folder, "DIE_ATTACH_Fallout_P08.xlsx")

        if not os.path.exists(mapping_file):
            print(f"⚠️ ไม่พบไฟล์: {mapping_file}")
            return average_file

        # โหลดไฟล์ mapping แรก
        df_map = pd.read_excel(mapping_file, engine='openpyxl')
        print(f"📊 ข้อมูล mapping: {len(df_map)} แถว")

        # ตรวจสอบคอลัมน์ที่จำเป็น
        required_cols = ["Package Code", "Cust Code", "Product Number", "bom_no"]
        missing_cols = [col for col in required_cols if col not in df_map.columns]
        
        if missing_cols:
            print(f"⚠️ ไม่พบคอลัมน์: {missing_cols}")
            return average_file

        # สร้างคอลัมน์ Device จากไฟล์แรก
        df_map["Device"] = df_map[["Package Code", "Cust Code", "Product Number"]].astype(str).agg('_'.join, axis=1)
        
        # เลือกคอลัมน์ที่ต้องการจากไฟล์แรก
        map_cols = ["bom_no", "Device"]
        if "#of Die" in df_map.columns:
            map_cols.append("#of Die")
        elif "of Die" in df_map.columns:
            map_cols.append("of Die")
        
        df_map_selected = df_map[map_cols]

        # Merge ข้อมูลจากไฟล์แรก
        df_merged = df_average.merge(df_map_selected, on="bom_no", how="left")
        print(f"✅ Map ไฟล์แรกสำเร็จ: {len(df_merged)} แถว")

        # Filter เฉพาะ Device ที่มีในไฟล์ที่สอง (ถ้ามี)
        if os.path.exists(mapping_file2):
            df_map2 = pd.read_excel(mapping_file2, engine='openpyxl')
            print(f"📊 ข้อมูล mapping2: {len(df_map2)} แถว")
            print(f"📋 คอลัมน์ใน mapping2: {list(df_map2.columns)}")
            
            # ถ้ามีคอลัมน์ Device ในไฟล์ที่สอง
            if "Device" in df_map2.columns:
                # ดึงเฉพาะรายการ Device ที่มีในไฟล์ที่สอง
                devices_in_file2 = set(df_map2['Device'].dropna().unique())
                print(f"🔍 Device ในไฟล์ที่ 2: {len(devices_in_file2)} รายการ")
                
                # แสดงจำนวนข้อมูลก่อน filter
                before_filter = len(df_merged)
                
                # Filter เฉพาะ Device ที่มีในไฟล์ที่สอง
                df_merged = df_merged[df_merged['Device'].isin(devices_in_file2)].copy()
                
                after_filter = len(df_merged)
                removed_count = before_filter - after_filter
                
                print(f"✅ Filter Device สำเร็จ: {after_filter} แถว")
                print(f"📊 ข้อมูลก่อน filter: {before_filter} แถว")
                print(f"🔍 ข้อมูลที่เหลือ (มี Device ในไฟล์ที่ 2): {after_filter} แถว")
                print(f"🗑️ ข้อมูลที่ถูกตัดออก: {removed_count} แถว")
                
                if after_filter == 0:
                    print("⚠️ ไม่มี Device ใดที่ตรงกันระหว่างไฟล์แรกและไฟล์ที่สอง")
                    # ถ้าไม่มีข้อมูลที่ match สามารถ return ไฟล์เดิมได้
                    return average_file
            else:
                print("⚠️ ไม่พบคอลัมน์ Device ในไฟล์ที่สอง")
        else:
            print(f"⚠️ ไม่พบไฟล์: {mapping_file2}")

        # บันทึกไฟล์ที่ map แล้ว
        output_dir = os.path.dirname(average_file)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        mapped_file = os.path.join(output_dir, f"mapped_data_{timestamp}.xlsx")
        
        df_merged.to_excel(mapped_file, index=False, engine='openpyxl')
        print(f"💾 บันทึกไฟล์ที่ map แล้ว: {mapped_file}")
        
        return mapped_file

    except Exception as e:
        print(f"❌ เกิดข้อผิดพลาดในการ map ข้อมูล: {e}")
        return average_file

def DA_AUTO_UPH(file_path, temp_root, start_date=None, end_date=None):
    """ฟังก์ชันหลักสำหรับประมวลผล Die Attack"""
    try:
        # ตรวจสอบ input type
        if isinstance(file_path, list):
            if len(file_path) == 0:
                print("❌ ไม่มีไฟล์ในรายการ")
                return None
            actual_file_path = file_path[0]  # ใช้ไฟล์แรก
            print(f"⚠️ รับรายการไฟล์ ({len(file_path)} ไฟล์) ใช้ไฟล์แรก: {actual_file_path}")
        else:
            actual_file_path = file_path
        
        df_cleaned, grouped_average, used_start_date, used_end_date = process_die_attack_data(
            actual_file_path, start_date, end_date)
        
        cleaned_file, average_file = save_results(
        df_cleaned, grouped_average, used_start_date, used_end_date, temp_root)
        
        print(f"ช่วงวันที่: {used_start_date} ถึง {used_end_date}")
        
        if not os.path.exists(average_file):
            print("❌ ไม่พบไฟล์ average_file")
            return None

        # Map ข้อมูลเพิ่มเติม
        mapped_file = map_data(average_file)
        print(f"📁 ส่งคืนไฟล์: {mapped_file}")
        return mapped_file

    except Exception as e:
        print(f"❌ DA_AUTO_UPH error: {e}")
        return None

