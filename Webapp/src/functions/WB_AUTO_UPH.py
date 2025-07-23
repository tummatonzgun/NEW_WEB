import pandas as pd
import numpy as np
from scipy.stats import zscore
import os
from datetime import datetime

class WireBondingAnalyzer:
    def __init__(self):
        self.nobump_df = None
        self.wb_data = None
        self.efficiency_df = None
        self.raw_data = None
    
    def normalize_model_name(self, model_name):
        """ทำความสะอาดและรวมชื่อรุ่นเครื่องที่คล้ายกัน"""
        if not isinstance(model_name, str):
            model_name = str(model_name)
        
        model_name = model_name.strip().upper()
        
        # รวม WB3100 ทุกเวอร์ชัน
        if 'WB3100' in model_name:
            return 'WB3100'
        
        # สามารถเพิ่มกฎการรวมรุ่นอื่นๆ ที่นี่
        if 'WB3200' in model_name:
            return 'WB3200'
        
        if 'WB3300' in model_name:
            return 'WB3300'
            
        return model_name

    def clean_model_names(self, df):
        """ทำความสะอาดชื่อรุ่นเครื่อง"""
        df = df.copy()
        if 'machine model' in df.columns:
            df['machine model'] = df['machine model'].apply(self.normalize_model_name)
        return df
    
    def find_wire_data_file(self, directory_path=None):
        """หาไฟล์ Wire Data ในโฟลเดอร์ data_wireWB (สำหรับเว็บ)"""
        try:
            # ใช้ path ตายตัวสำหรับ data_wireWB
            if directory_path is None:
                # หา path ของ data_wireWB จาก current directory
                current_dir = os.path.dirname(os.path.abspath(__file__))
                src_dir = os.path.dirname(current_dir)
                wire_dir = os.path.join(src_dir, "data_MAP")
            else:
                wire_dir = directory_path
            
            if not os.path.exists(wire_dir):
                print(f"Wire data directory not found: {wire_dir}")
                return None
            
            wire_files = []
            for filename in os.listdir(wire_dir):
                if (filename.lower().endswith(('.xlsx', '.xls')) and 
                    ('wire' in filename.lower() or 'book' in filename.lower())):
                    wire_files.append(os.path.join(wire_dir, filename))
            
            if not wire_files:
                print(f"No Wire data file found in: {wire_dir}")
                return None
            
            if len(wire_files) > 1:
                print(f"Multiple Wire files found: {[os.path.basename(f) for f in wire_files]}")
                print(f"Using the first one: {os.path.basename(wire_files[0])}")
            
            print(f"🔗 Using Wire data file: {os.path.basename(wire_files[0])}")
            return wire_files[0]
        
        except Exception as e:
            print(f"Error finding Wire data file: {e}")
            return None
    
    def load_data(self, uph_path, wire_data_path=None):
        """โหลดข้อมูลที่จำเป็น"""
        try:
            # ถ้าไม่ระบุ wire_data_path ให้หาในโฟลเดอร์เดียวกับ uph_path
            if wire_data_path is None:
                directory_path = os.path.dirname(uph_path)
                wire_data_path = self.find_wire_data_file(directory_path)
                
                if wire_data_path is None:
                    print("Wire data file not found. Please specify the path manually.")
                    return False
            
            # โหลดข้อมูล Wire Data
            print(f"📊 Loading Wire data from: {os.path.basename(wire_data_path)}")
            try:
                self.nobump_df = pd.read_excel(wire_data_path)
                self.nobump_df.columns = self.nobump_df.columns.str.strip().str.upper()
                print(f"✅ Wire data loaded: {len(self.nobump_df)} rows, columns: {list(self.nobump_df.columns)}")
            except Exception as e:
                print(f"❌ Error loading Wire data: {e}")
                return False
            
            # โหลดข้อมูล UPH
            print(f"📊 Loading UPH data from: {os.path.basename(uph_path)}")
            try:
                if uph_path.endswith('.csv'):
                    self.raw_data = pd.read_csv(uph_path, encoding='utf-8-sig')
                else:
                    self.raw_data = pd.read_excel(uph_path)
                
                # ทำความสะอาดคอลัมน์
                self.raw_data.columns = self.raw_data.columns.str.strip().str.lower()
                
                # แก้ไขชื่อคอลัมน์ให้เป็นมาตรฐาน
                if 'machine_model' in self.raw_data.columns:
                    self.raw_data.rename(columns={'machine_model': 'machine model'}, inplace=True)
                
                print(f"✅ UPH data loaded: {len(self.raw_data)} rows, columns: {list(self.raw_data.columns)}")
                
                # ตรวจสอบคอลัมน์ที่จำเป็น
                required_columns = ['uph', 'machine model', 'bom_no']
                missing_columns = [col for col in required_columns if col not in self.raw_data.columns]
                
                if missing_columns:
                    print(f"❌ Missing required columns in UPH data: {missing_columns}")
                    print(f"📋 Available columns: {list(self.raw_data.columns)}")
                    return False
                
            except Exception as e:
                print(f"❌ Error loading UPH data: {e}")
                return False
            
            print("✅ Data loaded successfully!")
            return True
            
        except Exception as e:
            print(f"❌ Error loading data: {e}")
            return False
    
    def calculate_wire_per_unit(self, bom_no):
        """คำนวณจำนวนสายต่อหน่วย"""
        try:
            bom_no = str(bom_no).strip().upper()
            bom_data = self.nobump_df[self.nobump_df['BOM_NO'].astype(str).str.strip().str.upper() == bom_no]
            
            if bom_data.empty:
                return 1.0
            
            no_bump = float(bom_data['NO_BUMP'].iloc[0]) if 'NO_BUMP' in bom_data.columns and not bom_data['NO_BUMP'].empty else 0
            num_required = float(bom_data['NUMBER_REQUIRED'].iloc[0]) if 'NUMBER_REQUIRED' in bom_data.columns and not bom_data['NUMBER_REQUIRED'].empty else 0
            
            wire_per_unit = (no_bump / 2) + num_required
            return wire_per_unit if wire_per_unit > 0 else 1.0
        except Exception as e:
            print(f"Error calculating wire per unit for BOM {bom_no}: {e}")
            return 1.0
    
    def remove_outliers(self, df):
        """ลบ outliers จากข้อมูลแบ่งตาม BOM และ Machine Model"""
        try:
            if df.empty:
                return df, {}
                
            df = self.clean_model_names(df)
            
            # ตรวจสอบคอลัมน์ที่จำเป็น
            required_cols = ['uph', 'machine model', 'bom_no']
            missing_cols = [col for col in required_cols if col not in df.columns]
            if missing_cols:
                raise KeyError(f"Missing required columns: {missing_cols}")
            
            # แบ่งข้อมูลตาม BOM และ Machine Model
            grouped = df.groupby(['bom_no', 'machine model'])
            cleaned_data = []
            outlier_info = {}
            
            for (bom_no, model), group_data in grouped:
                group_data = group_data.copy()
                original_count = len(group_data)
                
                # ข้ามถ้าข้อมูลน้อยกว่า 15 จุด
                if len(group_data) < 15:
                    cleaned_data.append(group_data)
                    outlier_info[(bom_no, model)] = {
                        'original_count': original_count,
                        'removed_count': 0,
                        'final_count': original_count
                    }
                    continue
                
                # กระบวนการตัด Outlier แบบอัตโนมัติ
                current_data = group_data
                
                for iteration in range(20):  # จำกัดจำนวนรอบ
                    # ใช้ Z-Score (±3σ)
                    z_threshold = 3
                    z_scores = zscore(current_data['uph'])
                    z_filtered = current_data[(z_scores >= -z_threshold) & (z_scores <= z_threshold)]
                    
                    # ตรวจสอบว่ายังมี Outlier หรือไม่
                    if not self._has_outliers(z_filtered['uph']):
                        current_data = z_filtered
                        break
                    
                    # ใช้ IQR (1.5*IQR)
                    Q1 = current_data['uph'].quantile(0.25)
                    Q3 = current_data['uph'].quantile(0.75)
                    IQR = Q3 - Q1
                    iqr_filtered = current_data[
                        (current_data['uph'] >= Q1 - 1.5*IQR) & 
                        (current_data['uph'] <= Q3 + 1.5*IQR)]
                    
                    if not self._has_outliers(iqr_filtered['uph']):
                        current_data = iqr_filtered
                        break
                    
                    current_data = iqr_filtered
                
                cleaned_data.append(current_data)
                final_count = len(current_data)
                
                # เก็บข้อมูลการตัด outlier
                outlier_info[(bom_no, model)] = {
                    'original_count': original_count,
                    'removed_count': original_count - final_count,
                    'final_count': final_count
                }
            result_df = pd.concat(cleaned_data) if cleaned_data else df
            return result_df, outlier_info
        
        except Exception as e:
            print(f"Error in remove_outliers: {e}")
            return df, {}
    
    def _has_outliers(self, series):
        """ตรวจสอบว่ายังมี Outlier หรือไม่"""
        if len(series) < 3:
            return False
        z_scores = zscore(series)
        return (abs(z_scores) > 3).any()
    
    def preprocess_data(self):
        """เตรียมข้อมูลก่อนคำนวณ"""
        try:
            if self.raw_data is None:
                raise ValueError("No data loaded")
            
            # คัดลอกข้อมูลและทำความสะอาด
            df = self.raw_data.copy()
            df.columns = df.columns.str.strip().str.lower()
            
            # ตรวจสอบคอลัมน์ที่จำเป็น
            required_cols = ['uph', 'machine model', 'bom_no']
            missing_cols = [col for col in required_cols if col not in df.columns]
            if missing_cols:
                raise KeyError(f"Missing required columns: {missing_cols}")
            
            # แปลงประเภทข้อมูล
            df['uph'] = pd.to_numeric(df['uph'], errors='coerce')
            df['bom_no'] = df['bom_no'].astype(str).str.strip().str.upper()
            
            # ลบแถวที่ไม่มีค่า UPH หรือ BOM_NO
            df = df.dropna(subset=['uph', 'bom_no'])
            
            # ทำความสะอาดชื่อรุ่นเครื่อง (เวอร์ชันปรับปรุง)
            df = self.clean_model_names(df)
            
            self.wb_data = df
            return True
        
        except Exception as e:
            print(f"Error in preprocess_data: {e}")
            return False
    
    def calculate_efficiency(self):
        """คำนวณประสิทธิภาพการทำงาน"""
        try:
            print(f"🔄 Starting calculate_efficiency...")
            
            if not self.preprocess_data():
                print(f"❌ Preprocess data failed")
                return None
            
            print(f"📊 Preprocessing completed. Data shape: {self.wb_data.shape}")
            
            # ตัด Outlier และเก็บข้อมูลการตัด
            cleaned_data, outlier_info = self.remove_outliers(self.wb_data)
            
            if cleaned_data.empty:
                print(f"❌ No data remaining after outlier removal")
                return None
            
            print(f"📊 After outlier removal. Data shape: {cleaned_data.shape}")
            
            # กลุ่มข้อมูลตาม BOM และรุ่นเครื่อง
            grouped = cleaned_data.groupby(['bom_no', 'machine model'])
            results = []
            
            print(f"📊 Processing {len(grouped)} groups...")
            
            for i, ((bom_no, model), group) in enumerate(grouped):
                try:
                    print(f"🔍 Processing group {i+1}/{len(grouped)}: BOM={bom_no}, Model={model}")
                    
                    # คำนวณค่าเฉลี่ย UPH
                    mean_uph = group['uph'].mean()
                    count = len(group)
                    
                    print(f"   📈 Mean UPH: {mean_uph:.2f}, Count: {count}")
                    
                    # คำนวณ Wire Per Unit
                    wire_per_unit = self.calculate_wire_per_unit(bom_no)
                    print(f"   🔌 Wire Per Unit: {wire_per_unit:.2f}")
                    
                    # คำนวณประสิทธิภาพ (UPH)
                    efficiency = mean_uph / wire_per_unit if wire_per_unit > 0 else 0
                    print(f"   ⚡ Efficiency (UPH): {efficiency:.3f}")
                    
                    # ดึงข้อมูลเพิ่มเติม
                    operation = group['operation'].iloc[0] if 'operation' in group.columns else 'N/A'
                    optn_code = group['optn_code'].iloc[0] if 'optn_code' in group.columns else 'N/A'
                    
                    # ดึงข้อมูล date_time_start ถ้ามี
                    date_time_start = group['date_time_start'].iloc[0] if 'date_time_start' in group.columns else 'N/A'
                    
                    # ดึงข้อมูลการตัด outlier
                    outlier_data = outlier_info.get((bom_no, model), {
                        'original_count': count,
                        'removed_count': 0,
                        'final_count': count
                    })
                    
                    result_entry = {
                        'Date_Time_Start': date_time_start,
                        'BOM': bom_no,
                        'Model': model,
                        'Operation': operation,
                        'Optn_Code': optn_code,
                        'Wire Per Hour': round(mean_uph, 2),
                        'Wire_Per_Unit': round(wire_per_unit, 2),
                        'UPH': round(efficiency, 3),
                        'Data_Points': count,
                        'Original_Count': outlier_data['original_count'],
                        'Outliers_Removed': outlier_data['removed_count']
                    }
                    
                    results.append(result_entry)
                    print(f"   ✅ Group processed successfully")
                    
                except Exception as group_error:
                    print(f"   ❌ Error processing group {bom_no}-{model}: {group_error}")
                    continue
            
            if not results:
                print(f"❌ No results generated")
                return None
            
            self.efficiency_df = pd.DataFrame(results)
            print(f"✅ Efficiency calculation completed. Generated {len(self.efficiency_df)} results")
            
            return self.efficiency_df
        
        except Exception as e:
            print(f"❌ Error in calculate_efficiency: {e}")
            import traceback
            print(f"🔍 Traceback: {traceback.format_exc()}")
            return None
    
    def export_to_excel(self, file_path=None):
        """ส่งออกผลลัพธ์เป็นไฟล์ Excel"""
        print(f"🔄 Starting export_to_excel...")
        
        if self.efficiency_df is None:
            print(f"❌ Export failed: efficiency_df is None")
            return False
            
        if self.efficiency_df.empty:
            print(f"❌ Export failed: efficiency_df is empty")
            return False
        
        print(f"📊 Data to export: {len(self.efficiency_df)} rows")
        
        try:
            # สร้างโฟลเดอร์ output_WB_AUTO_UPH หากยังไม่มี
            output_dir = 'output_WB_AUTO_UPH'
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
                print(f"📁 Created output directory: {output_dir}")
            
            # กำหนดชื่อไฟล์เริ่มต้นหากไม่ระบุ
            if file_path is None:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                file_path = os.path.join(output_dir, f'wb_analysis_results_{timestamp}.xlsx')
            else:
                # ตรวจสอบว่าได้ระบุโฟลเดอร์หรือไม่ หากไม่ระบุให้ใช้ output_dir
                if not os.path.dirname(file_path):
                    file_path = os.path.join(output_dir, file_path)
            
            print(f"📄 Export file path: {file_path}")
            
            # ตรวจสอบและสร้างโฟลเดอร์ที่ต้องการ
            output_directory = os.path.dirname(file_path)
            if output_directory and not os.path.exists(output_directory):
                os.makedirs(output_directory)
                print(f"📁 Created directory: {output_directory}")
            
            # ตรวจสอบสิทธิ์การเขียน
            try:
                test_file = os.path.join(output_directory, 'test_write.tmp')
                with open(test_file, 'w') as f:
                    f.write('test')
                os.remove(test_file)
                print(f"✅ Write permission verified for: {output_directory}")
            except Exception as perm_error:
                print(f"❌ Write permission error: {perm_error}")
                return False
            
            # เริ่มส่งออก Excel
            print(f"📝 Starting Excel export...")
            
            try:
                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                    # Sheet 1: ผลลัพธ์ UPH
                    print(f"✏️ Writing UPH_Results sheet...")
                    self.efficiency_df.to_excel(
                        writer, sheet_name='UPH_Results', index=False)
                    
                    # Sheet 2: สรุปตามรุ่นเครื่อง (ถ้ามีข้อมูลเพียงพอ)
                    if len(self.efficiency_df) > 0:
                        try:
                            print(f"✏️ Writing Model_Summary sheet...")
                            model_summary = self.efficiency_df.groupby('Model').agg({
                                'UPH': ['mean', 'std', 'count', 'min', 'max'],
                                'Wire Per Hour': 'mean',
                                'Wire_Per_Unit': 'mean'
                            }).round(3)
                            model_summary.to_excel(writer, sheet_name='Model_Summary')
                        except Exception as model_error:
                            print(f"⚠️ Warning: Could not create Model_Summary sheet: {model_error}")
                    
                    # Sheet 3: สรุปภาพรวม
                    try:
                        print(f"✏️ Writing Overall_Summary sheet...")
                        overall_stats = {
                            'Average_UPH': round(self.efficiency_df['UPH'].mean(), 3),
                            'Average_WPH': round(self.efficiency_df['Wire Per Hour'].mean(), 2),
                            'Total_Groups': len(self.efficiency_df),
                            'Total_Data_Points': self.efficiency_df['Data_Points'].sum(),
                            'Total_Outliers_Removed': self.efficiency_df['Outliers_Removed'].sum()
                        }
                        overall_df = pd.DataFrame.from_dict(
                            overall_stats, orient='index', columns=['Value'])
                        overall_df.to_excel(writer, sheet_name='Overall_Summary')
                    except Exception as overall_error:
                        print(f"⚠️ Warning: Could not create Overall_Summary sheet: {overall_error}")
                
                print(f"✅ Excel file created successfully")
                        
            except Exception as excel_error:
                print(f"❌ Excel export error: {excel_error}")
                print(f"🔄 Trying alternative method with xlsxwriter...")
                
                # ลองใช้ xlsxwriter แทน
                try:
                    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                        self.efficiency_df.to_excel(
                            writer, sheet_name='UPH_Results', index=False)
                    print(f"✅ Excel file created with xlsxwriter")
                except Exception as xlsxwriter_error:
                    print(f"❌ xlsxwriter also failed: {xlsxwriter_error}")
                    return False
            
            # ตรวจสอบว่าไฟล์ถูกสร้างจริงและมีขนาดมากกว่า 0
            if os.path.exists(file_path):
                file_size = os.path.getsize(file_path)
                print(f"✅ File created successfully: {file_path} (size: {file_size} bytes)")
                if file_size > 0:
                    return True
                else:
                    print(f"❌ File was created but is empty")
                    return False
            else:
                print(f"❌ File was not created: {file_path}")
                return False
        
        except Exception as e:
            print(f"❌ Unexpected error in export_to_excel: {e}")
            import traceback
            print(f"🔍 Traceback: {traceback.format_exc()}")
            return False

# === Web Interface Functions ===
def get_available_uph_files():
    """ดึงรายชื่อไฟล์ UPH ที่มีในโฟลเดอร์ data_WB สำหรับเว็บ"""
    try:
        # หา path ของ data_WB จาก current directory
        current_dir = os.path.dirname(os.path.abspath(__file__))
        src_dir = os.path.dirname(current_dir)
        uph_dir = os.path.join(src_dir, "data_WB")
        
        if not os.path.exists(uph_dir):
            return []
        
        uph_files = []
        for filename in os.listdir(uph_dir):
            if (filename.lower().endswith(('.xlsx', '.xls')) and 
                ('uph' in filename.lower() or 'apl' in filename.lower() or 'wb_data' in filename.lower())):
                uph_files.append({
                    'filename': filename,
                    'filepath': os.path.join(uph_dir, filename),
                    'size': os.path.getsize(os.path.join(uph_dir, filename))
                })
        
        # เรียงตามชื่อไฟล์
        uph_files.sort(key=lambda x: x['filename'])
        return uph_files
        
    except Exception as e:
        print(f"Error getting UPH files: {e}")
        return []

def get_wire_data_file():
    """ดึง path ของไฟล์ Wire Data จากโฟลเดอร์ data_MAP สำหรับเว็บ"""
    try:
        current_dir = os.path.dirname(os.path.abspath(__file__))
        src_dir = os.path.dirname(current_dir)
        wire_dir = os.path.join(src_dir, "data_MAP")  # <-- เปลี่ยนตรงนี้

        if not os.path.exists(wire_dir):
            return None

        for filename in os.listdir(wire_dir):
            if (filename.lower().endswith(('.xlsx', '.xls')) and 
                ('wire' in filename.lower() or 'book' in filename.lower())):
                return {
                    'filename': filename,
                    'filepath': os.path.join(wire_dir, filename)
                }

        return None

    except Exception as e:
        print(f"Error getting Wire data file: {e}")
        return None

def run_wb_auto_uph_web_multiple(selected_uph_files, output_filename=None, output_dir=None):
    """
    รันการวิเคราะห์ WB_AUTO_UPH สำหรับหลายไฟล์ UPH และรวมผลลัพธ์
    
    Args:
        selected_uph_files (list): รายชื่อไฟล์ UPH ที่เลือกจากเว็บ
        output_filename (str, optional): ชื่อไฟล์ output ที่ต้องการ
        output_dir (str, optional): โฟลเดอร์สำหรับไฟล์ output
    
    Returns:
        dict: ผลลัพธ์การวิเคราะห์
    """
    try:
        print(f"🚀 Starting WB_AUTO_UPH Multiple Files Analysis...")
        print(f"📁 Processing {len(selected_uph_files)} UPH files...")
        
        # หาไฟล์ Wire Data อัตโนมัติ
        wire_data = get_wire_data_file()
        if not wire_data:
            return {
                'success': False,
                'error': 'ไม่พบไฟล์ Wire Data ในโฟลเดอร์ data_wireWB'
            }
        
        current_dir = os.path.dirname(os.path.abspath(__file__))
        src_dir = os.path.dirname(current_dir)
        
        # ตรวจสอบว่าไฟล์ UPH ทั้งหมดมีอยู่
        uph_paths = []
        for selected_file in selected_uph_files:
            uph_path = os.path.join(src_dir, "data_WB", selected_file)
            if not os.path.exists(uph_path):
                return {
                    'success': False,
                    'error': f'ไม่พบไฟล์ UPH: {selected_file}'
                }
            uph_paths.append(uph_path)
        
        print(f"📁 Files to process:")
        print(f"   Wire Data: {wire_data['filename']}")
        for i, file in enumerate(selected_uph_files):
            print(f"   UPH Data {i+1}: {file}")
        
        # เก็บผลลัพธ์จากทุกไฟล์
        all_results = []
        total_groups_all = 0
        total_outliers_removed_all = 0
        total_original_data_all = 0
        total_data_points_all = 0
        file_summary = []
        
        # ประมวลผลแต่ละไฟล์
        for i, (uph_path, selected_file) in enumerate(zip(uph_paths, selected_uph_files)):
            print(f"\n🔄 Processing file {i+1}/{len(selected_uph_files)}: {selected_file}")
            
            # สร้าง analyzer ใหม่สำหรับแต่ละไฟล์
            analyzer = WireBondingAnalyzer()
            
            # โหลดข้อมูล
            if not analyzer.load_data(uph_path, wire_data['filepath']):
                print(f"⚠️ Warning: Could not load data from {selected_file}, skipping...")
                continue
            
            # คำนวณประสิทธิภาพ
            efficiency_df = analyzer.calculate_efficiency()
            
            if efficiency_df is None or efficiency_df.empty:
                print(f"⚠️ Warning: No results from {selected_file}, skipping...")
                continue
            
            # เพิ่มคอลัมน์ระบุแหล่งที่มาของไฟล์
            efficiency_df['Source_File'] = selected_file
            
            # รวมผลลัพธ์
            all_results.append(efficiency_df)
            
            # สรุปสถิติของไฟล์นี้
            file_groups = len(efficiency_df)
            file_outliers = efficiency_df['Outliers_Removed'].sum()
            file_original = efficiency_df['Original_Count'].sum()
            file_data_points = efficiency_df['Data_Points'].sum()
            
            total_groups_all += file_groups
            total_outliers_removed_all += file_outliers
            total_original_data_all += file_original
            total_data_points_all += file_data_points
            
            file_summary.append({
                'file': selected_file,
                'groups': file_groups,
                'outliers_removed': file_outliers,
                'original_data': file_original,
                'data_points': file_data_points
            })
            
            print(f"✅ File {i+1} processed: {file_groups} groups, {file_data_points} data points")
        
        if not all_results:
            return {
                'success': False,
                'error': 'ไม่สามารถประมวลผลไฟล์ใดๆ ได้'
            }
        
        # รวมผลลัพธ์จากทุกไฟล์
        print(f"\n📊 Combining results from {len(all_results)} files...")
        combined_df = pd.concat(all_results, ignore_index=True)
        
        # สร้างชื่อไฟล์ output
        if not output_filename:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"WB_Analysis_Combined_{timestamp}"
        
        # สร้างโฟลเดอร์ output (ใช้ output_dir ถ้าระบุ, ไม่เช่นนั้นใช้ temp)
        if output_dir is None:
            # fallback: ใช้ temp ใน project root
            current_dir = os.path.dirname(os.path.abspath(__file__))
            project_root = os.path.dirname(os.path.dirname(current_dir))
            output_dir = os.path.join(project_root, "temp")
        os.makedirs(output_dir, exist_ok=True)

        # เพิ่ม .xlsx ถ้ายังไม่มี
        if not output_filename.endswith('.xlsx'):
            output_filename += '.xlsx'

        output_path = os.path.join(output_dir, output_filename)
        
        # Export ไฟล์รวม
        print(f"💾 Exporting combined results...")
        
        try:
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                # Sheet 1: ผลลัพธ์รวมทั้งหมด
                combined_df.to_excel(writer, sheet_name='Combined_Results', index=False)
                
                # Sheet 2: สรุปตามรุ่นเครื่อง
                if len(combined_df) > 0:
                    try:
                        model_summary = combined_df.groupby('Model').agg({
                            'UPH': ['mean', 'std', 'count', 'min', 'max'],
                            'Wire Per Hour': 'mean',
                            'Wire_Per_Unit': 'mean'
                        }).round(3)
                        model_summary.to_excel(writer, sheet_name='Model_Summary')
                    except Exception as model_error:
                        print(f"⚠️ Warning: Could not create Model_Summary sheet: {model_error}")
                
                # Sheet 3: สรุปตามไฟล์
                try:
                    file_summary_df = pd.DataFrame(file_summary)
                    file_summary_df.to_excel(writer, sheet_name='File_Summary', index=False)
                except Exception as file_error:
                    print(f"⚠️ Warning: Could not create File_Summary sheet: {file_error}")
                
                # Sheet 4: สรุปภาพรวม
                try:
                    overall_stats = {
                        'Total_Files_Processed': len(all_results),
                        'Total_Groups': total_groups_all,
                        'Average_UPH': round(combined_df['UPH'].mean(), 3),
                        'Average_WPH': round(combined_df['Wire Per Hour'].mean(), 2),
                        'Total_Data_Points': total_data_points_all,
                        'Total_Outliers_Removed': total_outliers_removed_all,
                        'Overall_Data_Quality': round((1 - total_outliers_removed_all/total_original_data_all) * 100, 2) if total_original_data_all > 0 else 0
                    }
                    overall_df = pd.DataFrame.from_dict(
                        overall_stats, orient='index', columns=['Value'])
                    overall_df.to_excel(writer, sheet_name='Overall_Summary')
                except Exception as overall_error:
                    print(f"⚠️ Warning: Could not create Overall_Summary sheet: {overall_error}")
        
        except Exception as export_error:
            return {
                'success': False,
                'error': f'ไม่สามารถส่งออกไฟล์ผลลัพธ์ได้: {str(export_error)}'
            }
        
        # ตรวจสอบไฟล์ที่สร้าง
        if not os.path.exists(output_path):
            return {
                'success': False,
                'error': 'ไฟล์ผลลัพธ์ไม่ถูกสร้าง'
            }
        
        # คำนวณสถิติรวม
        avg_efficiency = combined_df['UPH'].mean() if not combined_df.empty else 0
        
        print(f"✅ WB_AUTO_UPH Multiple Files Analysis completed successfully!")
        
        return {
            'success': True,
            'message': f'วิเคราะห์ข้อมูล Wire Bond จาก {len(selected_uph_files)} ไฟล์สำเร็จ',
            'output_file': output_filename,
            'output_path': output_path,
            'summary': {
                'files_processed': len(all_results),
                'total_groups': total_groups_all,
                'average_efficiency': round(avg_efficiency, 3),
                'outliers_removed': total_outliers_removed_all,
                'total_original_data': total_original_data_all,
                'data_quality': round((1 - total_outliers_removed_all/total_original_data_all) * 100, 2) if total_original_data_all > 0 else 0,
                'total_data_points': total_data_points_all
            },
            'wire_data_file': wire_data['filename'],
            'uph_data_files': selected_uph_files,
            'file_details': file_summary
        }
        
    except Exception as e:
        print(f"❌ Error in WB_AUTO_UPH Multiple Files Analysis: {e}")
        import traceback
        print(f"🔍 Full traceback:")
        print(traceback.format_exc())
        return {
            'success': False,
            'error': f'เกิดข้อผิดพลาด: {str(e)}'
        }

def run_wb_auto_uph_web(selected_uph_file, output_filename=None):
    """
    รันการวิเคราะห์ WB_AUTO_UPH สำหรับเว็บ
    
    Args:
        selected_uph_file (str): ชื่อไฟล์ UPH ที่เลือกจากเว็บ
        output_filename (str, optional): ชื่อไฟล์ output ที่ต้องการ
    
    Returns:
        dict: ผลลัพธ์การวิเคราะห์
    """
    try:
        print(f"🚀 Starting WB_AUTO_UPH Web Analysis...")
        
        # หาไฟล์ Wire Data อัตโนมัติ
        wire_data = get_wire_data_file()
        if not wire_data:
            return {
                'success': False,
                'error': 'ไม่พบไฟล์ Wire Data ในโฟลเดอร์ data_wireWB'
            }
        
        # หา path ของไฟล์ UPH ที่เลือก
        current_dir = os.path.dirname(os.path.abspath(__file__))
        src_dir = os.path.dirname(current_dir)
        uph_path = os.path.join(src_dir, "data_WB", selected_uph_file)
        
        if not os.path.exists(uph_path):
            return {
                'success': False,
                'error': f'ไม่พบไฟล์ UPH: {selected_uph_file}'
            }
        
        print(f"📁 Files to process:")
        print(f"   Wire Data: {wire_data['filename']}")
        print(f"   UPH Data: {selected_uph_file}")
        
        # สร้าง analyzer
        analyzer = WireBondingAnalyzer()
        
        # โหลดข้อมูล
        print(f"📊 Loading data...")
        if not analyzer.load_data(uph_path, wire_data['filepath']):
            return {
                'success': False,
                'error': 'ไม่สามารถโหลดข้อมูลได้'
            }
        
        # คำนวณประสิทธิภาพ
        print(f"⚡ Calculating efficiency...")
        efficiency_df = analyzer.calculate_efficiency()
        
        if efficiency_df is None or efficiency_df.empty:
            return {
                'success': False,
                'error': 'ไม่สามารถคำนวณประสิทธิภาพได้ หรือไม่มีข้อมูลหลังจากประมวลผล'
            }
        
        # สร้างชื่อไฟล์ output
        if not output_filename:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"WB_Analysis_{timestamp}"
        
        # สร้างโฟลเดอร์ output
        upload_dir = os.path.join(src_dir, "Upload")
        os.makedirs(upload_dir, exist_ok=True)
        
        # เพิ่ม .xlsx ถ้ายังไม่มี
        if not output_filename.endswith('.xlsx'):
            output_filename += '.xlsx'
        
        output_path = os.path.join(upload_dir, output_filename)
        
        # Export ไฟล์
        print(f"💾 Exporting results...")
        if not analyzer.export_to_excel(output_path):
            return {
                'success': False,
                'error': 'ไม่สามารถส่งออกไฟล์ผลลัพธ์ได้'
            }
        
        # ตรวจสอบไฟล์ที่สร้าง
        if not os.path.exists(output_path):
            return {
                'success': False,
                'error': 'ไฟล์ผลลัพธ์ไม่ถูกสร้าง'
            }
        
        # สรุปผลลัพธ์
        total_groups = len(efficiency_df)
        avg_efficiency = efficiency_df['UPH'].mean() if not efficiency_df.empty else 0
        total_data_points = efficiency_df['Data_Points'].sum()
        total_outliers_removed = efficiency_df['Outliers_Removed'].sum()
        total_original_data = efficiency_df['Original_Count'].sum()
        
        print(f"✅ WB_AUTO_UPH Web Analysis completed successfully!")
        
        return {
            'success': True,
            'message': 'วิเคราะห์ข้อมูล Wire Bond สำเร็จ',
            'output_file': output_filename,
            'output_path': output_path,
            'summary': {
                'total_groups': total_groups,
                'average_efficiency': round(avg_efficiency, 3),
                'outliers_removed': total_outliers_removed,
                'total_original_data': total_original_data,
                'data_quality': round((1 - total_outliers_removed/total_original_data) * 100, 2) if total_original_data > 0 else 0,
                'total_data_points': total_data_points
            },
            'wire_data_file': wire_data['filename'],
            'uph_data_file': selected_uph_file
        }
        
    except Exception as e:
        print(f"❌ Error in WB_AUTO_UPH Web Analysis: {e}")
        import traceback
        print(f"🔍 Full traceback:")
        print(traceback.format_exc())
        return {
            'success': False,
            'error': f'เกิดข้อผิดพลาด: {str(e)}'
        }
        
def run(input_dir, output_dir, uph_filename=None, wire_filename=None, **kwargs):
    print(f"🚀 Starting WB_AUTO_UPH execution...")
    
    analyzer = WireBondingAnalyzer()
    
    # Debug: แสดงข้อมูล input
    print(f"🔍 WB_AUTO_UPH Debug Info:")
    print(f"   Input Dir: {input_dir}")
    print(f"   Output Dir: {output_dir}")
    print(f"   UPH Filename: {uph_filename}")
    print(f"   Wire Filename: {wire_filename}")
    print(f"   Input Dir exists: {os.path.exists(input_dir)}")
    
    try:
        if os.path.exists(input_dir):
            files_in_input = os.listdir(input_dir)
            print(f"   Files in input_dir: {files_in_input}")
        else:
            raise Exception(f"Input directory does not exist: {input_dir}")
        
        # ใช้ input_dir ที่ส่งมาจากระบบเว็บ (temporary directory)
        if uph_filename and wire_filename:
            uph_file = os.path.join(input_dir, uph_filename)
            wire_file = os.path.join(input_dir, wire_filename)
            print(f"   UPH File Path: {uph_file}")
            print(f"   Wire File Path: {wire_file}")
        else:
            uph_file = None
            wire_file = None
            
            # หาไฟล์ในโฟลเดอร์ input_dir แบบอัตโนมัติ
            for fname in files_in_input:
                print(f"   Checking file: {fname}")
                fname_lower = fname.lower()
                
                # ตรวจสอบไฟล์ UPH (WB, UTL, UPH, Data)
                if (('wb' in fname_lower or 'utl' in fname_lower or 'uph' in fname_lower or 'data' in fname_lower) 
                    and fname_lower.endswith(('.xlsx', '.xls', '.csv')) 
                    and 'wire' not in fname_lower and 'book' not in fname_lower):
                    uph_file = os.path.join(input_dir, fname)
                    print(f"   ✅ Found UPH file: {uph_file}")
                
                # ตรวจสอบไฟล์ Wire (Wire, Book)
                elif (('wire' in fname_lower or 'book' in fname_lower) 
                      and fname_lower.endswith(('.xlsx', '.xls'))):
                    wire_file = os.path.join(input_dir, fname)
                    print(f"   ✅ Found Wire file: {wire_file}")
            
            # ถ้าหาไม่เจอแบบอัตโนมัติ ให้ใช้ไฟล์แรกที่เจอ
            if not uph_file or not wire_file:
                files_in_dir = [f for f in files_in_input if f.endswith(('.xlsx', '.xls', '.csv'))]
                print(f"   Available files: {files_in_dir}")
                
                if len(files_in_dir) >= 2:
                    # เอาไฟล์ที่มีขนาดใหญ่กว่าเป็น UPH file
                    files_with_size = []
                    for f in files_in_dir:
                        file_path = os.path.join(input_dir, f)
                        try:
                            size = os.path.getsize(file_path)
                            files_with_size.append((f, size))
                        except:
                            files_with_size.append((f, 0))
                    
                    # เรียงตามขนาดไฟล์
                    files_with_size.sort(key=lambda x: x[1], reverse=True)
                    
                    if not uph_file:
                        uph_file = os.path.join(input_dir, files_with_size[0][0])
                        print(f"   📊 Auto-selected UPH file (largest): {files_with_size[0][0]}")
                    
                    if not wire_file:
                        # หาไฟล์ที่ไม่ใช่ UPH file
                        for fname, size in files_with_size:
                            if fname != os.path.basename(uph_file):
                                wire_file = os.path.join(input_dir, fname)
                                print(f"   📊 Auto-selected Wire file: {fname}")
                                break
        
        # ตรวจสอบว่าพบไฟล์ครบหรือไม่
        if not uph_file or not wire_file:
            missing_files = []
            if not uph_file:
                missing_files.append("UPH data file")
            if not wire_file:
                missing_files.append("Wire data file")
            
            available_files = [f for f in files_in_input if f.endswith(('.xlsx', '.xls', '.csv'))]
            error_msg = f"ไม่พบไฟล์ที่จำเป็น: {', '.join(missing_files)}\nไฟล์ที่มีในโฟลเดอร์: {', '.join(available_files)}\nกรุณาตรวจสอบให้แน่ใจว่าอัปโหลดไฟล์ครบ 2 ไฟล์ (.xlsx หรือ .xls)"
            print(f"❌ {error_msg}")
            raise Exception(error_msg)
        
        # ตรวจสอบว่าไฟล์มีอยู่จริง
        if not os.path.exists(uph_file):
            raise Exception(f"ไม่พบไฟล์ UPH: {uph_file}")
        if not os.path.exists(wire_file):
            raise Exception(f"ไม่พบไฟล์ Wire Data: {wire_file}")
        
        print(f"✅ Files validated successfully")
        
        # โหลดข้อมูล
        print(f"📁 Loading data...")
        if not analyzer.load_data(uph_file, wire_file):
            raise Exception("โหลดข้อมูลไม่สำเร็จ")
        
        print(f"📊 Data loaded successfully")
        
        # คำนวณประสิทธิภาพ
        print(f"⚡ Calculating efficiency...")
        efficiency_df = analyzer.calculate_efficiency()
        if efficiency_df is None or efficiency_df.empty:
            raise Exception("คำนวณประสิทธิภาพไม่สำเร็จ หรือไม่มีข้อมูลหลังจากประมวลผล")
        
        print(f"✅ Efficiency calculation completed")
        
        # สร้างโฟลเดอร์ output
        print(f"📁 Creating output directory...")
        os.makedirs(output_dir, exist_ok=True)
        output_path = os.path.join(output_dir, "WB_AUTO_UPH_RESULT.xlsx")
        
        print(f"📄 Output path: {output_path}")
        
        # ส่งออกไฟล์
        print(f"💾 Exporting to Excel...")
        if not analyzer.export_to_excel(output_path):
            raise Exception("ส่งออกไฟล์ผลลัพธ์ไม่สำเร็จ - กรุณาตรวจสอบสิทธิ์การเขียนไฟล์หรือพื้นที่ดิสก์")
        
        # ตรวจสอบไฟล์ผลลัพธ์
        if not os.path.exists(output_path):
            raise Exception(f"ไฟล์ผลลัพธ์ไม่ถูกสร้าง: {output_path}")
        
        file_size = os.path.getsize(output_path)
        if file_size == 0:
            raise Exception(f"ไฟล์ผลลัพธ์ว่างเปล่า: {output_path}")
        
        print(f"✅ WB_AUTO_UPH completed successfully!")
        print(f"📄 Output file: {output_path} (size: {file_size} bytes)")
        return output_path
        
    except Exception as e:
        print(f"❌ WB_AUTO_UPH failed: {str(e)}")
        import traceback
        print(f"🔍 Full traceback:")
        print(traceback.format_exc())
        raise e

def WB_AUTO_UPH(input_path, output_dir, start_date=None, end_date=None):
    """
    WB_AUTO_UPH function สำหรับเรียกใช้ผ่าน workflow ปกติ
    รองรับการรับไฟล์จากโฟลเดอร์ data_WB, data_MAP หรือ list ของไฟล์ UPH
    """
    try:
        print(f"🚀 Starting WB_AUTO_UPH workflow...")
        print(f"📁 Input: {input_path}")
        print(f"📁 Output: {output_dir}")

        # ถ้า input_path เป็น list ให้ใช้ run_wb_auto_uph_web_multiple
        if isinstance(input_path, list):
            result = run_wb_auto_uph_web_multiple(input_path, output_dir=output_dir)
            if result.get("success"):
                print(f"✅ WB_AUTO_UPH Multiple Files completed: {result['output_path']}")
                return result["output_path"]
            else:
                raise Exception(result.get("error", "Unknown error"))
        else:
            # กรณีปกติ (โฟลเดอร์หรือไฟล์เดียว)
            result_path = run(input_path, output_dir)
            return result_path

    except Exception as e:
        print(f"❌ WB_AUTO_UPH workflow failed: {str(e)}")
        raise e