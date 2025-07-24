import pandas as pd
import numpy as np
import os
import sys
import json
import requests
from datetime import datetime
from urllib.parse import urlparse
import matplotlib.pyplot as plt
import io
import base64

def load_data_from_source(source):
    """
    โหลดข้อมูลจากแหล่งต่างๆ (Excel, JSON file, JSON API)
    """
    # รองรับทั้งกรณี source เป็น list หรือ str
    if isinstance(source, list):
        df_list = [load_data_from_source(s) for s in source]
        return pd.concat(df_list, ignore_index=True)
    
    try:
        # ตรวจสอบว่าเป็น URL หรือไม่
        parsed = urlparse(source)
        is_url = bool(parsed.netloc)
        
        if is_url:
            print(f"📡 กำลังเรียกข้อมูลจาก API: {source}")
            response = requests.get(source, timeout=30)
            response.raise_for_status()
            
            # ตรวจสอบ Content-Type
            content_type = response.headers.get('content-type', '').lower()
            
            if 'application/json' in content_type or source.endswith('.json'):
                # ข้อมูล JSON จาก API
                json_data = response.json()
                df = process_json_data(json_data)
                print(f"✅ โหลดข้อมูลจาก API สำเร็จ: {len(df)} แถว")
                
            else:
                raise ValueError(f"ประเภทข้อมูลไม่รองรับ: {content_type}")
                
        else:
            # ไฟล์ในเครื่อง
            if not os.path.exists(source):
                raise FileNotFoundError(f"ไม่พบไฟล์: {source}")
                
            file_ext = os.path.splitext(source)[1].lower()
            
            if file_ext in ['.xlsx', '.xls']:
                print(f"📄 กำลังโหลดไฟล์ Excel: {source}")
                df = pd.read_excel(source)
                print(f"✅ โหลดไฟล์ Excel สำเร็จ: {len(df)} แถว")
                
            elif file_ext == '.json':
                print(f"📄 กำลังโหลดไฟล์ JSON: {source}")
                with open(source, 'r', encoding='utf-8') as f:
                    json_data = json.load(f)
                df = process_json_data(json_data)
                print(f"✅ โหลดไฟล์ JSON สำเร็จ: {len(df)} แถว")
                
            elif file_ext == '.csv':
                print(f"📄 กำลังโหลดไฟล์ CSV: {source}")
                df = pd.read_csv(source)
                print(f"✅ โหลดไฟล์ CSV สำเร็จ: {len(df)} แถว")
                
            else:
                raise ValueError(f"ประเภทไฟล์ไม่รองรับ: {file_ext}")
        
        return df
        
    except requests.RequestException as e:
        raise Exception(f"❌ เกิดข้อผิดพลาดในการเรียก API: {str(e)}")
    except json.JSONDecodeError as e:
        raise Exception(f"❌ ข้อมูล JSON ไม่ถูกต้อง: {str(e)}")
    except Exception as e:
        raise Exception(f"❌ เกิดข้อผิดพลาดในการโหลดข้อมูล: {str(e)}")

def process_json_data(json_data):
    """
    ประมวลผลข้อมูล JSON เป็น DataFrame
    
    Parameters:
    json_data: ข้อมูล JSON ที่ได้จาก API หรือไฟล์
    
    Returns:
    pandas.DataFrame: ข้อมูลที่ประมวลผลแล้ว
    """
    print("🔄 กำลังประมวลผลข้อมูล JSON...")
    
    # กรณีที่ 1: ข้อมูลเป็น list ของ objects
    if isinstance(json_data, list):
        df = pd.DataFrame(json_data)
        
    # กรณีที่ 2: ข้อมูลอยู่ใน key ใดๆ
    elif isinstance(json_data, dict):
        # ลองหา key ที่มี list ของข้อมูล
        possible_keys = ['data', 'results', 'items', 'records', 'rows', 'content']
        
        data_found = False
        for key in possible_keys:
            if key in json_data and isinstance(json_data[key], list):
                df = pd.DataFrame(json_data[key])
                print(f"📋 ใช้ข้อมูลจาก key: '{key}'")
                data_found = True
                break
        
        if not data_found:
            # ถ้าไม่เจอ ลองใช้ key แรกที่เป็น list
            for key, value in json_data.items():
                if isinstance(value, list):
                    df = pd.DataFrame(value)
                    print(f"📋 ใช้ข้อมูลจาก key: '{key}'")
                    data_found = True
                    break
            
            if not data_found:
                # ถ้ายังไม่เจอ ลองแปลง dict ตรงๆ
                df = pd.DataFrame([json_data])
                print("📋 ใช้ข้อมูล JSON ทั้งหมดเป็น 1 แถว")
    
    else:
        raise ValueError("รูปแบบข้อมูล JSON ไม่รองรับ")
    
    print(f"✅ ประมวลผล JSON เสร็จสิ้น: {len(df)} แถว, {len(df.columns)} คอลัมน์")
    
    # แสดงชื่อคอลัมน์
    print("📊 คอลัมน์ที่พบ:")
    for i, col in enumerate(df.columns):
        print(f"  {i+1}. {col}")
    
    return df

def load_data_with_config(source, config=None):

    try:
        parsed = urlparse(source)
        is_url = bool(parsed.netloc)
        
        if is_url and config:
            print(f"📡 กำลังเรียกข้อมูลจาก API พร้อม config: {source}")
            
            # ตั้งค่า request parameters
            request_params = {
                'timeout': config.get('timeout', 30)
            }
            
            if 'headers' in config:
                request_params['headers'] = config['headers']
                
            if 'params' in config:
                request_params['params'] = config['params']
                
            if 'auth' in config:
                request_params['auth'] = tuple(config['auth'])
            
            response = requests.get(source, **request_params)
            response.raise_for_status()
            
            json_data = response.json()
            df = process_json_data(json_data)
            print(f"✅ โหลดข้อมูลจาก API พร้อม config สำเร็จ: {len(df)} แถว")
            
            return df
        else:
            return load_data_from_source(source)
            
    except Exception as e:
        raise Exception(f"❌ เกิดข้อผิดพลาดในการโหลดข้อมูลพร้อม config: {str(e)}")

# ฟังก์ชันเดิมทั้งหมด (เหมือนเดิมแต่แก้ไข Wire Per Hour)
def apply_zscore(df):
    """ตัด outliers ด้วยวิธี Z-Score (±3 standard deviations)"""
    col_map = {col.lower(): col for col in df.columns}
    if 'uph' not in col_map:
        raise KeyError("ไม่พบคอลัมน์ UPH ในข้อมูล")
    uph_col = col_map['uph']

    mean = df[uph_col].mean()
    std = df[uph_col].std()
    if std == 0:
        return df  # ถ้า std = 0, return ค่าเดิม
    z_scores = (df[uph_col] - mean) / std
    filtered = df[(z_scores >= -3) & (z_scores <= 3)].copy()
    filtered['Outlier_Method'] = 'Z-Score ±3'
    filtered['Wire Per Hour'] = 1  # แก้ไข: กำหนดค่าให้ถูกต้อง

    return filtered

def has_outlier(df):
    """ตรวจสอบว่ามี outliers หรือไม่ด้วยวิธี IQR"""
    col_map = {col.lower(): col for col in df.columns}
    if 'uph' not in col_map:
        raise KeyError("ไม่พบคอลัมน์ UPH ในข้อมูล")
    uph_col = col_map['uph']

    Q1 = df[uph_col].quantile(0.25)
    Q3 = df[uph_col].quantile(0.75)
    IQR = Q3 - Q1
    lower = Q1 - 1.5 * IQR
    upper = Q3 + 1.5 * IQR
    return ((df[uph_col] < lower) | (df[uph_col] > upper)).sum() > 0

def apply_iqr(df):
    """ตัด outliers ด้วยวิธี IQR (Interquartile Range)"""
    col_map = {col.lower(): col for col in df.columns}
    if 'uph' not in col_map:
        raise KeyError("ไม่พบคอลัมน์ UPH ในข้อมูล")
    uph_col = col_map['uph']

    Q1 = df[uph_col].quantile(0.25)
    Q3 = df[uph_col].quantile(0.75)
    IQR = Q3 - Q1
    lower = Q1 - 1.5 * IQR
    upper = Q3 + 1.5 * IQR
    filtered = df[(df[uph_col] >= lower) & (df[uph_col] <= upper)].copy()
    filtered['Outlier_Method'] = 'IQR'
    filtered['Wire Per Hour'] = 1  # แก้ไข: เพิ่มการกำหนดค่า
    return filtered

def remove_outliers_auto(df_model, max_iter=20):
    """ตัด outliers อัตโนมัติด้วยการใช้ Z-Score และ IQR แบบวนลูป"""
    col_map = {col.lower(): col for col in df_model.columns}
    if 'uph' not in col_map:
        raise KeyError("ไม่พบคอลัมน์ UPH ในข้อมูล")
    uph_col = col_map['uph']

    df_model[uph_col] = pd.to_numeric(df_model[uph_col], errors='coerce')
    df_model = df_model.dropna(subset=[uph_col])

    if len(df_model) < 15:  
        df_model['Outlier_Method'] = 'ไม่ตัด (ข้อมูลน้อย)'
        df_model['Wire Per Hour'] = 1  # แก้ไข: เพิ่มการกำหนดค่า
        return df_model

    current_df = df_model.copy()

    for i in range(max_iter):
        print(f"=== รอบที่ {i+1} ===")
        z_df = apply_zscore(current_df)
        if not has_outlier(z_df):
            z_df['Outlier_Method'] = f'Z-Score Loop ×{i+1}'
            z_df['Wire Per Hour'] = 1  # แก้ไข: เพิ่มการกำหนดค่า
            return z_df

        iqr_df = apply_iqr(z_df)
        if not has_outlier(iqr_df):
            iqr_df['Outlier_Method'] = f'IQR Loop ×{i+1}'
            iqr_df['Wire Per Hour'] = 1  # แก้ไข: เพิ่มการกำหนดค่า
            return iqr_df

        current_df = iqr_df

    current_df['Outlier_Method'] = f'IQR-Z-Score Loop ×{max_iter}+'
    current_df['Wire Per Hour'] = 1  # แก้ไข: เพิ่มการกำหนดค่า
    return current_df

def remove_outliers(df):
    """ตัด outliers ตามกลุ่ม BOM และ Machine Model และเพิ่มคอลัมน์จำนวนข้อมูลก่อน/หลังตัด"""
    col_map = {col.lower(): col for col in df.columns}
    
    # หาคอลัมน์ Machine Model
    model_col = None
    if 'machine model' in col_map:
        model_col = col_map['machine model']
    elif 'machine_model' in col_map:
        model_col = col_map['machine_model']
    else:
        raise KeyError("ไม่พบคอลัมน์ Machine Model หรือ Machine_Model ในข้อมูล")
    
    # หาคอลัมน์ bom_no
    bom_col = None
    if 'bom_no' in col_map:
        bom_col = col_map['bom_no']
    elif 'bom no' in col_map:
        bom_col = col_map['bom no']
    else:
        raise KeyError("ไม่พบคอลัมน์ bom_no ในข้อมูล")
    
    result_dfs = []
    
    for (bom_no, machine_model), group_df in df.groupby([bom_col, model_col]):
        before_count = len(group_df)
        cleaned_group = remove_outliers_auto(group_df)
        after_count = len(cleaned_group)
        # เพิ่มคอลัมน์ใหม่
        cleaned_group['DataPoints_Before'] = before_count
        cleaned_group['DataPoints_After'] = after_count
        result_dfs.append(cleaned_group)
    
    return pd.concat(result_dfs, ignore_index=True)

def time_series_analysis(df):
    """แปลงข้อมูลวันที่ให้อยู่ในรูปแบบที่ใช้งานได้"""
    col_map = {col.lower(): col for col in df.columns}
    
    # หาคอลัมน์วันที่ที่เป็นไปได้
    date_cols = []
    for col_name in df.columns:
        if any(keyword in col_name.lower() for keyword in ['date', 'time', 'วัน', 'เวลา']):
            date_cols.append(col_name)
    
    if not date_cols:
        print("ไม่พบคอลัมน์วันที่ในข้อมูล")
        return df
    
    # ใช้คอลัมน์วันที่แรกที่พบ
    date_col = date_cols[0]
    print(f"ใช้คอลัมน์วันที่: {date_col}")
    
    # แปลงเป็น datetime และจัดรูปแบบ
    df['date_time_start'] = pd.to_datetime(df[date_col], errors='coerce')
    
    # จัดรูปแบบเป็น YYYY/MM/DD
    df['date_time_start'] = df['date_time_start'].dt.strftime('%Y/%m/%d')
    
    # กรองข้อมูลที่แปลงไม่ได้
    invalid_dates = df['date_time_start'].isna().sum()
    if invalid_dates > 0:
        print(f"พบวันที่ที่แปลงไม่ได้: {invalid_dates} แถว")
        df = df.dropna(subset=['date_time_start'])
    
    print(f"แปลงวันที่เสร็จสิ้น รูปแบบ: {df['date_time_start'].iloc[0] if len(df) > 0 else 'ไม่มีข้อมูล'}")
    
    return df
    
def find_max_or_min_date(df):
    if 'date_time_start' not in df.columns:
        raise KeyError("ไม่พบคอลัมน์ date_time_start ในข้อมูล")

    max_date = df['date_time_start'].max()
    min_date = df['date_time_start'].min()
    return max_date, min_date

def get_date_range_auto(df):
    """สร้างช่วงวันที่อัตโนมัติ (ใช้ทั้งหมด)"""
    max_date, min_date = find_max_or_min_date(df)
    print(f"ใช้ช่วงวันที่ทั้งหมด: {min_date} ถึง {max_date}")
    return min_date, max_date

def filter_data_by_date(df, start_date, end_date):
    # กรองข้อมูลตามช่วงวันที่
    filtered_df = df[df['date_time_start'].between(start_date, end_date)].copy()
    
    if len(filtered_df) == 0:
        print("ไม่พบข้อมูลในช่วงวันที่ที่เลือก")
        return None
    
    print(f"พบข้อมูล {len(filtered_df)} แถว ในช่วงวันที่ที่เลือก")
    print(f"ข้อมูลเดิม: {len(df)} แถว")
    print(f"ข้อมูลที่กรอง: {len(filtered_df)} แถว ({len(filtered_df)/len(df)*100:.1f}%)")
    
    return filtered_df

def calculate_group_average(df, start_date, end_date):
    """คำนวณค่าเฉลี่ยตามกลุ่ม และแนบคอลัมน์เดิมที่ต้องการ"""
    col_map = {col.lower(): col for col in df.columns}
    model_col = col_map.get('machine model') or col_map.get('machine_model')
    bom_col = col_map.get('bom_no') or col_map.get('bom no')
    uph_col = col_map.get('uph')

    # เลือกคอลัมน์ที่ต้องการแสดงในไฟล์ average
    columns_to_keep = [bom_col, 'operation', model_col, uph_col]

    # คำนวณ mean เฉพาะ uph
    grouped = df.groupby([bom_col, model_col], as_index=False).agg({uph_col: 'mean'})

    # ดึงค่าแรกของคอลัมน์อื่นในแต่ละกลุ่ม
    other_cols = [c for c in columns_to_keep if c not in [bom_col, model_col, uph_col]]
    firsts = df.groupby([bom_col, model_col], as_index=False)[other_cols].first()

    # เพิ่ม DataPoints_Before/After (ใช้ค่าแรกในกลุ่ม)
    if 'DataPoints_Before' in df.columns and 'DataPoints_After' in df.columns:
        data_points = df.groupby([bom_col, model_col], as_index=False)[['DataPoints_Before', 'DataPoints_After']].first()
        grouped_average = pd.merge(grouped, firsts, on=[bom_col, model_col], how='left')
        grouped_average = pd.merge(grouped_average, data_points, on=[bom_col, model_col], how='left')
    else:
        grouped_average = pd.merge(grouped, firsts, on=[bom_col, model_col], how='left')

    print(f"\n=== ค่าเฉลี่ย UPH ตามกลุ่ม (ช่วงวันที่ {start_date} ถึง {end_date}) ===")
    print(grouped_average)
    return grouped_average

def save_results(df_cleaned, grouped_average, start_date, end_date, output_dir):
    """บันทึกผลลัพธ์ลงไฟล์"""
    os.makedirs(output_dir, exist_ok=True)
    
    # สร้างชื่อไฟล์ด้วยวันที่และเวลา
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    date_range = f"{start_date.replace('/', '')}_to_{end_date.replace('/', '')}"
    
    # บันทึกข้อมูลที่ตัด outliers แล้ว
    cleaned_file = os.path.join(output_dir, f"cleaned_data_{date_range}_{timestamp}.xlsx")
    df_cleaned.to_excel(cleaned_file, index=False)
    print(f"บันทึกข้อมูลที่ตัด outliers แล้ว: {cleaned_file}")
    
    # บันทึกค่าเฉลี่ยตามกลุ่ม
    average_file = os.path.join(output_dir, f"group_average_{date_range}_{timestamp}.xlsx")
    grouped_average.to_excel(average_file, index=False)
    print(f"บันทึกค่าเฉลี่ยตามกลุ่ม: {average_file}")
    
    return cleaned_file, average_file

def process_die_attack_data(source):
    """ประมวลผลข้อมูล Die Attack - รองรับ JSON API"""
    print("=== เริ่มต้นการประมวลผลข้อมูล Die Attack ===")
    
    # อ่านข้อมูลจากแหล่งต่างๆ
    try:
        df = load_data_from_source(source)
        print(f"ข้อมูลเริ่มต้น: {len(df)} แถว")
    except Exception as e:
        raise Exception(f"ไม่สามารถโหลดข้อมูลได้: {str(e)}")
    
    # ขั้นตอนที่ 1: แปลงข้อมูลวันที่
    print("\n1. แปลงข้อมูลวันที่...")
    df = time_series_analysis(df)
    
    # ขั้นตอนที่ 2: ใช้ช่วงวันที่ทั้งหมด (อัตโนมัติ)
    print("\n2. กำหนดช่วงวันที่...")
    start_date, end_date = get_date_range_auto(df)
    
    # ขั้นตอนที่ 3: กรองข้อมูลตามวันที่
    print("\n3. กรองข้อมูลตามวันที่...")
    df_filtered = filter_data_by_date(df, start_date, end_date)
    
    if df_filtered is None:
        raise Exception("ไม่มีข้อมูลในช่วงวันที่ที่เลือก")
    
    # ขั้นตอนที่ 4: ตัด outliers
    print("\n4. ตัด outliers...")
    df_cleaned = remove_outliers(df_filtered)
    df_cleaned = df_cleaned.reset_index(drop=True)
    
    print(f"ข้อมูลหลังตัด outliers: {len(df_cleaned)} แถว")
    
    # ขั้นตอนที่ 5: คำนวณค่าเฉลี่ยตามกลุ่ม
    print("\n5. คำนวณค่าเฉลี่ยตามกลุ่ม...")
    grouped_average = calculate_group_average(df_cleaned, start_date, end_date)
    
    print("\n=== การประมวลผลเสร็จสิ้น ===")
    print(f"ข้อมูลสุดท้าย: {len(df_cleaned)} แถว")
    print(f"จำนวนกลุ่ม: {len(grouped_average)} กลุ่ม")
    
    return df_cleaned, grouped_average, start_date, end_date

def process_die_attack_data_with_date_range(file_path, start_date, end_date):
    """ประมวลผลข้อมูล Die Attack ด้วยช่วงวันที่ที่กำหนด"""
    print("=== เริ่มต้นการประมวลผลข้อมูล Die Attack (ช่วงวันที่กำหนด) ===")
    
    # อ่านข้อมูลจากแหล่งต่างๆ (รองรับ Excel, CSV, JSON, API)
    try:
        df = load_data_from_source(file_path)
        print(f"ข้อมูลเริ่มต้น: {len(df)} แถว")
    except Exception as e:
        raise Exception(f"ไม่สามารถโหลดข้อมูลได้: {str(e)}")
    
    # ขั้นตอนที่ 1: แปลงข้อมูลวันที่
    print("\n1. แปลงข้อมูลวันที่...")
    df = time_series_analysis(df)
    
    # ขั้นตอนที่ 2: แปลงวันที่จาก YYYY-MM-DD เป็น YYYY/MM/DD
    print(f"\n2. ใช้ช่วงวันที่ที่กำหนด: {start_date} ถึง {end_date}")
    formatted_start_date = start_date.replace('-', '/')
    formatted_end_date = end_date.replace('-', '/')
    
    # ขั้นตอนที่ 3: กรองข้อมูลตามวันที่
    print("\n3. กรองข้อมูลตามวันที่...")
    df_filtered = filter_data_by_date(df, formatted_start_date, formatted_end_date)
    
    if df_filtered is None:
        raise Exception("ไม่มีข้อมูลในช่วงวันที่ที่เลือก")
    
    # ขั้นตอนที่ 4: ตัด outliers
    print("\n4. ตัด outliers...")
    df_cleaned = remove_outliers(df_filtered)
    df_cleaned = df_cleaned.reset_index(drop=True)
    
    print(f"ข้อมูลหลังตัด outliers: {len(df_cleaned)} แถว")
    
    # ขั้นตอนที่ 5: คำนวณค่าเฉลี่ยตามกลุ่ม
    print("\n5. คำนวณค่าเฉลี่ยตามกลุ่ม...")
    grouped_average = calculate_group_average(df_cleaned, formatted_start_date, formatted_end_date)
    
    print("\n=== การประมวลผลเสร็จสิ้น ===")
    print(f"ข้อมูลสุดท้าย: {len(df_cleaned)} แถว")
    print(f"จำนวนกลุ่ม: {len(grouped_average)} กลุ่ม")
    
    return df_cleaned, grouped_average, formatted_start_date, formatted_end_date


def preview_date_range(file_path):
    """แสดงข้อมูลวันที่ในไฟล์ก่อนประมวลผล"""
    try:
        print("📅 กำลังตรวจสอบช่วงวันที่ในไฟล์...")
        
        # อ่านไฟล์ข้อมูล
        ext = os.path.splitext(file_path)[-1].lower()
        if ext in [".xlsx", ".xls"]:
            df = pd.read_excel(file_path, engine="openpyxl")
        elif ext == ".csv":
            df = pd.read_csv(file_path)
        elif ext == ".json":
            df = pd.read_json(file_path)
        else:
            raise ValueError("ไม่รองรับไฟล์ประเภทนี้")
        print(f"📄 ไฟล์มีข้อมูลทั้งหมด: {len(df):,} แถว")
        
        # หาคอลัมน์วันที่
        date_cols = []
        for col_name in df.columns:
            if any(keyword in col_name.lower() for keyword in ['date', 'time', 'วัน', 'เวลา']):
                date_cols.append(col_name)
        
        if not date_cols:
            print("⚠️ ไม่พบคอลัมน์วันที่ในข้อมูล")
            return None
        
        date_col = date_cols[0]
        print(f"🗓️ ใช้คอลัมน์วันที่: '{date_col}'")
        
        # แปลงข้อมูลวันที่
        df['temp_date'] = pd.to_datetime(df[date_col], errors='coerce')
        
        # กรองข้อมูลที่แปลงวันที่ได้
        valid_dates = df.dropna(subset=['temp_date'])
        invalid_count = len(df) - len(valid_dates)
        
        if len(valid_dates) == 0:
            print("❌ ไม่มีข้อมูลวันที่ที่ถูกต้องในไฟล์")
            return None
        
        # หาวันที่เริ่มต้นและสิ้นสุด
        min_date = valid_dates['temp_date'].min()
        max_date = valid_dates['temp_date'].max()
        
        # แสดงผลลัพธ์
        print(f"\n📊 สรุปข้อมูลวันที่:")
        print(f"  🗓️ วันที่เริ่มต้น: {min_date.strftime('%Y-%m-%d')} (ค.ศ.)")
        print(f"  🗓️ วันที่สิ้นสุด: {max_date.strftime('%Y-%m-%d')} (ค.ศ.)")
        print(f"  📈 จำนวนวัน: {(max_date - min_date).days + 1} วัน")
        print(f"  ✅ ข้อมูลวันที่ถูกต้อง: {len(valid_dates):,} แถว")
        
        if invalid_count > 0:
            print(f"  ⚠️ ข้อมูลวันที่ไม่ถูกต้อง: {invalid_count:,} แถว")
        
        # แสดงตัวอย่างข้อมูลตามช่วงเวลา
        print(f"\n📋 การกระจายข้อมูลตามเดือน:")
        monthly_counts = valid_dates.groupby(valid_dates['temp_date'].dt.to_period('M')).size()
        for period, count in monthly_counts.head(10).items():
            print(f"  📅 {period}: {count:,} แถว")
        
        if len(monthly_counts) > 10:
            print(f"  ... และอีก {len(monthly_counts) - 10} เดือน")
        
        return {
            'min_date': min_date.strftime('%Y-%m-%d'),
            'max_date': max_date.strftime('%Y-%m-%d'),
            'total_days': (max_date - min_date).days + 1,
            'valid_records': len(valid_dates),
            'invalid_records': invalid_count,
            'date_column': date_col,
            'monthly_distribution': {str(period): count for period, count in monthly_counts.to_dict().items()}
        }
        
    except Exception as e:
        print(f"❌ เกิดข้อผิดพลาดในการตรวจสอบวันที่: {str(e)}")
        return None

def DA_AUTO_UPH(file_path, temp_root, start_date=None, end_date=None):
    try:
        if start_date and end_date:
            start_date_fmt = start_date.replace("-", "/")
            end_date_fmt = end_date.replace("-", "/")
            df_cleaned, grouped_average, used_start_date, used_end_date = process_die_attack_data_with_date_range(file_path, start_date_fmt, end_date_fmt)
        else:
            df_cleaned, grouped_average, used_start_date, used_end_date = process_die_attack_data(file_path)
        cleaned_file, average_file = save_results(df_cleaned, grouped_average, used_start_date, used_end_date, temp_root)
        print("DEBUG: average_file path =", average_file)
        print(f"✅ ช่วงวันที่ที่ประมวลผลจริง: {used_start_date} ถึง {used_end_date}")
        if not os.path.exists(average_file):
            print("❌ ไม่พบไฟล์ average_file:", average_file)
            return None
        return average_file
    except Exception as e:
        print(f"❌ DA_AUTO_UPH error: {e}")
        return None

