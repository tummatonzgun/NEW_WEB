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
                wire_dir = os.path.join(src_dir, "data_wireWB")
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
            # ถ้าไม่ระบุ wire_data_path ให้หาในโฟลเดอร์ data_wireWB เสมอ
            if wire_data_path is None:
                # หา Wire data จากโฟลเดอร์ data_wireWB แทนโฟลเดอร์เดียวกับ UPH
                wire_data_path = self.find_wire_data_file(None)  # ส่ง None เพื่อใช้ path ของ data_wireWB
                
                if wire_data_path is None:
                    print("Wire data file not found in data_wireWB folder. Please check the folder exists and contains wire data files.")
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
        
        # ใช้ input_dir ที่ส่งมาจากระบบเว็บ (temporary directory) สำหรับ UPH files
        # Wire files จะอ่านจากโฟลเดอร์ data_wireWB
        # ไฟล์ใดก็ตามที่อัปโหลดในเว็บของ Wire Bond คือไฟล์ UPH files
        if uph_filename and wire_filename:
            uph_file = os.path.join(input_dir, uph_filename)
            # Wire file ไม่ใช้จาก input_dir แต่จะหาจาก data_wireWB โดยอัตโนมัติ
            wire_file = None  # จะให้ load_data หา wire file เอง
            print(f"   UPH File Path: {uph_file}")
            print(f"   Wire File: Will auto-detect from data_wireWB folder")
        else:
            uph_file = None
            wire_file = None
            
            # หาไฟล์ UPH ในโฟลเดอร์ input_dir (ที่อัปโหลดมา)
            # ไฟล์ใดก็ตามที่อัปโหลดมาคือไฟล์ UPH
            for fname in files_in_input:
                if fname.lower().endswith(('.xlsx', '.xls', '.csv')):
                    uph_file = os.path.join(input_dir, fname)
                    print(f"   ✅ Using uploaded file as UPH file: {uph_file}")
                    break  # หาเจอไฟล์แรกแล้วหยุด
            
            # Wire file จะให้ load_data หาจาก data_wireWB เอง
            wire_file = None
            
            # ตรวจสอบว่าหา UPH file เจอหรือไม่
            if not uph_file:
                files_in_dir = [f for f in files_in_input if f.endswith(('.xlsx', '.xls', '.csv'))]
                print(f"   Available files: {files_in_dir}")
                
                if len(files_in_dir) >= 1:
                    # เอาไฟล์แรกที่เจอเป็น UPH file (ไฟล์ใดก็ตามที่อัปโหลดมา)
                    uph_file = os.path.join(input_dir, files_in_dir[0])
                    print(f"   📊 Using first available file as UPH file: {files_in_dir[0]}")
        
        # ตรวจสอบว่าพบไฟล์ UPH หรือไม่
        if not uph_file:
            available_files = [f for f in files_in_input if f.endswith(('.xlsx', '.xls', '.csv'))]
            error_msg = f"ไม่พบไฟล์ UPH ที่จำเป็น\nไฟล์ที่มีในโฟลเดอร์: {', '.join(available_files)}\nกรุณาตรวจสอบให้แน่ใจว่าอัปโหลดไฟล์ UPH (.xlsx, .xls หรือ .csv)"
            print(f"❌ {error_msg}")
            raise Exception(error_msg)
        
        # ตรวจสอบว่าไฟล์ UPH มีอยู่จริง
        if not os.path.exists(uph_file):
            raise Exception(f"ไม่พบไฟล์ UPH: {uph_file}")
        
        print(f"✅ UPH file validated successfully")
        print(f"📋 Wire data will be loaded from data_wireWB folder automatically")
        
        # โหลดข้อมูล (wire_file จะให้ load_data หาจาก data_wireWB เอง)
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
    รองรับการรับไฟล์จากโฟลเดอร์ data_WB และ data_wireWB
    """
    try:
        print(f"🚀 Starting WB_AUTO_UPH workflow...")
        print(f"📁 Input: {input_path}")
        print(f"📁 Output: {output_dir}")
        
        # เรียกใช้ function หลัก
        result_path = run(input_path, output_dir)
        return result_path
        
    except Exception as e:
        print(f"❌ WB_AUTO_UPH workflow failed: {str(e)}")
        raise e