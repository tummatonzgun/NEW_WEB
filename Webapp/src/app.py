from flask import Flask, render_template, request, redirect, url_for, session, flash, jsonify, send_file
import os
import tempfile
import shutil
import socket
import json     
import pandas as pd
import requests
import sys
import time
import threading

# เพิ่ม path สำหรับ functions เพื่อให้ importlib หา module เจอ
FUNCTIONS_PATH = os.path.join(os.getcwd(), "functions")
if FUNCTIONS_PATH not in sys.path:
    sys.path.append(FUNCTIONS_PATH)

app = Flask(__name__)
app.secret_key = "your_secret_key"
app.api_base_url = "http://th3sroeeeng4/RTMSAPI/ApiAutoUph/api"

# Progress tracking system
progress_data = {}

class ProgressTracker:
    def __init__(self, task_id):
        self.task_id = task_id
        self.progress = 0
        self.status = "เริ่มต้น..."
        self.completed = False
        self.error = None
        progress_data[task_id] = self
    
    def update(self, progress, status):
        self.progress = progress
        self.status = status
        progress_data[self.task_id] = self
    
    def complete(self, status="เสร็จสิ้น"):
        self.progress = 100
        self.status = status
        self.completed = True
        progress_data[self.task_id] = self
    
    def error_occurred(self, error_msg):
        self.error = error_msg
        self.status = f"เกิดข้อผิดพลาด: {error_msg}"
        progress_data[self.task_id] = self

# Mapping operation -> function list
OPERATION_FUNCTIONS = {
    "Singulation": ["LOGVIEW"],
    "Pick & Place": ["PNP_CHANGE_TYPE", "PNP_AUTO_UPH"],
    "Die Attach": ["DA_AUTO_UPH"],
    "Wire Bond": ["WB_AUTO_UPH"],
}

@app.route("/", methods=["GET"])
def operation():
    return render_template("operation.html")

# --- เพิ่มฟีเจอร์เลือกไฟล์ในโฟลเดอร์และรองรับหลายไฟล์ ---
@app.route("/method", methods=["GET", "POST"])
def method():
    if request.method == "POST":
        operation = request.form.get("operation") or request.args.get("operation")
        session["operation"] = operation
        input_method = request.form.get("inputMethod")
        session["input_method"] = input_method

        temp_root = os.path.join(os.getcwd(), "temp")
        os.makedirs(temp_root, exist_ok=True)

        # รับไฟล์ (อัปโหลดหลายไฟล์)
        if input_method == "upload":
            files = request.files.getlist("file")
            uploaded_files = []
            if files and any(f.filename for f in files):
                for file in files:
                    if file and file.filename:
                        file_path = os.path.join(temp_root, file.filename)
                        file.save(file_path)
                        uploaded_files.append(file_path)
                session["uploaded_file_path"] = uploaded_files  # เก็บเป็น list
            else:
                flash("กรุณาเลือกไฟล์ก่อน", "error")
                return redirect(url_for("method", operation=operation))

        # รับโฟลเดอร์ (เลือกไฟล์ในโฟลเดอร์ได้)
        elif input_method == "folder":
            selected_folder = request.form.get("selected_folder")
            selected_files = request.form.getlist("selected_files")  # รับไฟล์ที่เลือกจากโฟลเดอร์
            if not selected_folder:
                flash("กรุณาเลือกโฟลเดอร์ก่อน", "error")
                return redirect(url_for("method", operation=operation))
            
            # แก้ไข: ใช้ path ที่ถูกต้องในการเข้าถึงโฟลเดอร์ข้อมูล
            src_dir = os.path.dirname(os.path.abspath(__file__))
            src_folder = os.path.join(src_dir, selected_folder)

            if not os.path.exists(src_folder):
                flash(f"ไม่พบโฟลเดอร์ที่เลือก: {src_folder}", "error")
                return redirect(url_for("method", operation=operation))
            
            temp_folder = tempfile.mkdtemp(dir=temp_root)
            copied_files = []
            if selected_files:
                # copy เฉพาะไฟล์ที่เลือก
                for fname in selected_files:
                    src_file = os.path.join(src_folder, fname)
                    if os.path.isfile(src_file):
                        shutil.copy2(src_file, temp_folder)
                        copied_files.append(os.path.join(temp_folder, fname))
            else:
                # ถ้าไม่ได้เลือกไฟล์ ให้ copy ทุกไฟล์ในโฟลเดอร์
                for fname in os.listdir(src_folder):
                    src_file = os.path.join(src_folder, fname)
                    if os.path.isfile(src_file):
                        shutil.copy2(src_file, temp_folder)
                        copied_files.append(os.path.join(temp_folder, fname))
            session["selected_folder"] = copied_files  # เก็บเป็น list

        # รับ API params (เหมือนเดิม)
        elif input_method == "api":
            endpoint = request.form.get("endpoint")
            plant = request.form.get("plant")
            year_quarter = request.form.get("year_quarter")
            api_operation = request.form.get("api_operation")
            bom_no = request.form.get("bom_no")
            session["endpoint"] = endpoint
            session["plant"] = plant
            session["year_quarter"] = year_quarter
            session["api_operation"] = api_operation
            session["bom_no"] = bom_no
            # สร้าง URL จาก base + endpoint
            api_url = f"{app.api_base_url}/{endpoint}"
            params = {}
            if plant: params["plant"] = plant
            if year_quarter: params["year_quarter"] = year_quarter
            if api_operation: params["operation"] = api_operation
            if bom_no: params["bom_no"] = bom_no
            try:
                response = requests.get(api_url, params=params)
                if response.status_code == 200:
                    content_type = response.headers.get('Content-Type', '')
                    if 'application/json' in content_type and response.text.strip():
                        try:
                            json_data = response.json()
                        except Exception as e:
                            flash(f"API ได้รับข้อมูลที่ไม่ใช่ JSON: {e} | ตัวอย่างข้อมูล: {response.text[:300]}", "error")
                            return redirect(url_for("method", operation=operation))
                        json_filename = f"api_{plant}_{year_quarter}_{api_operation}_{bom_no or 'none'}.json"
                        json_path = os.path.join(temp_root, json_filename)
                        with open(json_path, "w", encoding="utf-8") as f:
                            json.dump(json_data, f, ensure_ascii=False)
                        session["api_json_path"] = json_path
                    else:
                        if 'text/html' in content_type:
                            flash(f"API ไม่ได้ส่งข้อมูล JSON แต่ส่ง HTML (Content-Type: text/html)\n\nกรุณาตรวจสอบ URL endpoint ว่าเป็น API จริง ไม่ใช่ Swagger UI หรือหน้าเว็บ และตรวจสอบสิทธิ์การเข้าถึง API ปลายทาง\n\nตัวอย่างข้อมูล: {response.text[:300]}", "error")
                        else:
                            flash(f"API ไม่ได้ส่งข้อมูล JSON หรือข้อมูลว่างเปล่า | Content-Type: {content_type} | ตัวอย่างข้อมูล: {response.text[:300]}", "error")
                        return redirect(url_for("method", operation=operation))
                else:
                    flash(f"API ดึงข้อมูลไม่สำเร็จ: {response.status_code}", "error")
                    return redirect(url_for("method", operation=operation))
            except Exception as e:
                flash(f"API error: {e}", "error")
                return redirect(url_for("method", operation=operation))

        return redirect(url_for("function"))
    # GET: render หน้าเลือก method
    operation = request.args.get("operation", "")
    # --- เพิ่มสำหรับเลือกไฟล์ในโฟลเดอร์ ---
    folder_files = []
    selected_folder = request.args.get("selected_folder")
    if selected_folder:
        src_folder = os.path.join(os.getcwd(), selected_folder)
        if os.path.exists(src_folder):
            folder_files = [f for f in os.listdir(src_folder) if os.path.isfile(os.path.join(src_folder, f))]
    
    # เพิ่มรายชื่อโฟลเดอร์ข้อมูลที่มี
    data_folders = []
    # ใช้ path ของไฟล์ปัจจุบัน เพื่อให้ค้นหาโฟลเดอร์ข้อมูลได้ถูกตำแหน่ง
    src_dir = os.path.dirname(os.path.abspath(__file__))

    if os.path.exists(src_dir):
        all_items = os.listdir(src_dir)
        
        for item in all_items:
            item_path = os.path.join(src_dir, item)
            if os.path.isdir(item_path) and item.startswith('data_'):
                # นับจำนวนไฟล์ในโฟลเดอร์
                file_count = len([f for f in os.listdir(item_path) if os.path.isfile(os.path.join(item_path, f))])
                data_folders.append({
                    'name': item,
                    'path': item,  # แก้ path ให้ถูกต้อง
                    'file_count': file_count
                })
    
    return render_template("method.html", operation=operation, folder_files=folder_files, data_folders=data_folders)

@app.route("/function", methods=["GET", "POST"])
def function():
    if request.method == "POST":
        func_name = request.form.get("func_name")
        start_date = request.form.get("start_date")
        end_date = request.form.get("end_date")
        if not start_date or not end_date:
            start_date = None
            end_date = None
        input_method = session.get("input_method")
        operation = session.get("operation")
        result = None
        current_file = None

        # เตรียม path ไฟล์ที่ต้องประมวลผล
        file_path = None
        if input_method == "upload":
            file_path = session.get("uploaded_file_path")
        elif input_method == "folder":
            file_path = session.get("selected_folder")
        elif input_method == "api":
            file_path = session.get("api_json_path")

        # สร้าง task ID สำหรับ progress tracking
        task_id = f"task_{int(time.time())}_{func_name}"
        session["current_task_id"] = task_id

        # เริ่มประมวลผลในเธรดแยก
        def background_processing():
            try:
                temp_root = os.path.join(os.getcwd(), "temp")
                result = process_with_progress(func_name, file_path, temp_root, start_date, end_date, task_id)
                
                # --- สร้างไฟล์ผลลัพธ์สำหรับดาวน์โหลด ---
                export_file_path = None
                if isinstance(result, pd.DataFrame):
                    export_file_path = os.path.join(temp_root, f"result_{operation}_{func_name}.xlsx")
                    result.to_excel(export_file_path, index=False)
                elif isinstance(result, list):
                    # ถ้าเป็น list ของ path ให้ใช้ตัวแรกที่เป็นไฟล์จริง
                    for r in result:
                        if isinstance(r, str) and os.path.exists(r):
                            export_file_path = r
                            break
                elif isinstance(result, str) and os.path.exists(result):
                    export_file_path = result
                
                session["export_file_path"] = export_file_path
                session["current_file"] = file_path
                session["operation"] = operation
                session["func_name"] = func_name
                
            except Exception as e:
                tracker = progress_data.get(task_id)
                if tracker:
                    tracker.error_occurred(str(e))

        # เริ่มเธรดประมวลผล
        thread = threading.Thread(target=background_processing)
        thread.daemon = True
        thread.start()

        # ส่งไปหน้า processing
        return render_template("processing.html", task_id=task_id, func_name=func_name, operation=operation)

    # GET: render หน้าเลือกฟังก์ชัน (เพิ่ม preview date range)
    input_method = session.get("input_method")
    current_file = None
    file_path = None
    if input_method == "upload":
        file_path = session.get("uploaded_file_path")
        if file_path:
            if isinstance(file_path, list):
                current_file = [os.path.basename(f) for f in file_path]
                file_path_preview = file_path[0]
            else:
                current_file = os.path.basename(file_path)
                file_path_preview = file_path
    elif input_method == "folder":
        folder = session.get("selected_folder")
        if folder:
            if isinstance(folder, list):
                current_file = [os.path.basename(f) for f in folder]
                file_path_preview = folder[0]
            else:
                current_file = folder
                file_path_preview = folder
    elif input_method == "api":
        json_path = session.get("api_json_path")
        if json_path:
            current_file = os.path.basename(json_path)
            file_path_preview = json_path
    else:
        file_path_preview = None

    # Preview date range
    date_info = None
    if file_path_preview and os.path.exists(file_path_preview):
        try:
            from functions.da_auto_uph import preview_date_range
            date_info = preview_date_range(file_path_preview)
        except Exception as e:
            date_info = None

    operation = session.get("operation")
    functions = OPERATION_FUNCTIONS.get(operation, [])
    return render_template("function.html", functions=functions, current_file=current_file, operation=operation, date_info=date_info)

@app.route("/result", methods=["GET"])
def result():
    import pandas as pd
    export_file_path = session.get("export_file_path")
    current_file = session.get("current_file")
    operation = session.get("operation")
    func_name = session.get("func_name")
    table_html = None
    result_data = None
    error_message = None
    if not export_file_path:
        error_message = "export_file_path ไม่ถูกสร้าง กรุณาตรวจสอบการประมวลผลหรือฟังก์ชันที่เลือก"
    elif not os.path.exists(export_file_path):
        error_message = f"ไม่พบไฟล์ผลลัพธ์: {export_file_path} กรุณาตรวจสอบว่าไฟล์ถูกสร้างจริงหลังประมวลผล"
    if error_message:
        table_html = f"<pre>{error_message}</pre>"
    else:
        try:
            if export_file_path.endswith(".xlsx"):
                df = pd.read_excel(export_file_path)
            elif export_file_path.endswith(".csv"):
                df = pd.read_csv(export_file_path)
            else:
                df = None
            if df is not None:
                table_html = df.to_html(classes="table", border=0, index=False)
                result_data = df.to_dict(orient="records")
        except Exception as e:
            table_html = f"<pre>เกิดข้อผิดพลาดในการอ่านไฟล์ผลลัพธ์: {e}</pre>"
    return render_template("result.html", result=result_data, current_file=current_file, operation=operation, func_name=func_name, table_html=table_html)

@app.route("/api/", methods=["GET"])
def get_api_data():
    endpoint = request.args.get("endpoint")
    plant = request.args.get("plant")
    year_quarter = request.args.get("year_quarter")
    operation = request.args.get("operation")
    bom_no = request.args.get("bom_no")

    if not endpoint:
        return jsonify({"error": "No endpoint selected"}), 400

    url = f"{app.api_base_url}/{endpoint}"
    params = {}
    if plant: params["plant"] = plant
    if year_quarter: params["year_quarter"] = year_quarter
    if operation: params["operation"] = operation
    if bom_no: params["bom_no"] = bom_no

    try:
        response = requests.get(url, params=params)
        response.raise_for_status()
        content_type = response.headers.get('Content-Type', '')
        if 'application/json' in content_type and response.text.strip():
            try:
                data = response.json()
            except Exception as e:
                return jsonify({
                    "error": f"API ได้รับข้อมูลที่ไม่ใช่ JSON: {e}",
                    "example": response.text[:300]
                }), 500
            return jsonify({
                "request_url": response.url,
                "data": data
            })
        else:
            if 'text/html' in content_type:
                return jsonify({
                    "error": "API ไม่ได้ส่งข้อมูล JSON แต่ส่ง HTML (Content-Type: text/html). กรุณาตรวจสอบ URL endpoint ว่าเป็น API จริง ไม่ใช่ Swagger UI หรือหน้าเว็บ และตรวจสอบสิทธิ์การเข้าถึง API ปลายทาง",
                    "example": response.text[:300]
                }), 500
            else:
                return jsonify({
                    "error": f"API ไม่ได้ส่งข้อมูล JSON หรือข้อมูลว่างเปล่า | Content-Type: {content_type}",
                    "example": response.text[:300]
                }), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# === API สำหรับดึงไฟล์ในโฟลเดอร์ ===
@app.route("/api/folder_files", methods=["GET"])
def get_folder_files():
    """ดึงรายชื่อไฟล์ในโฟลเดอร์ที่เลือก"""
    try:
        folder_path = request.args.get("folder_path")
        if not folder_path:
            return jsonify({"error": "No folder path provided"}), 400
        
        full_path = os.path.join(os.getcwd(), folder_path)
        if not os.path.exists(full_path):
            return jsonify({"error": f"Folder not found: {folder_path}"}), 404
        
        files = []
        for filename in os.listdir(full_path):
            file_path = os.path.join(full_path, filename)
            if os.path.isfile(file_path):
                file_size = os.path.getsize(file_path)
                files.append({
                    'name': filename,
                    'size': file_size,
                    'size_mb': round(file_size / (1024 * 1024), 2)
                })
        
        # เรียงตามชื่อไฟล์
        files.sort(key=lambda x: x['name'])
        
        return jsonify({
            "folder": folder_path,
            "files": files,
            "file_count": len(files)
        })
        
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# === API สำหรับ Progress Tracking ===
@app.route("/api/progress/<task_id>", methods=["GET"])
def get_progress(task_id):
    """ดึงความคืบหน้าของงาน"""
    try:
        if task_id not in progress_data:
            return jsonify({"error": "Task not found"}), 404
        
        tracker = progress_data[task_id]
        return jsonify({
            "task_id": task_id,
            "progress": tracker.progress,
            "status": tracker.status,
            "completed": tracker.completed,
            "error": tracker.error
        })
        
    except Exception as e:
        return jsonify({"error": str(e)}), 500

def process_with_progress(func_name, file_path, temp_root, start_date=None, end_date=None, task_id=None):
    """ประมวลผลพร้อม progress tracking"""
    try:
        tracker = ProgressTracker(task_id) if task_id else None
        
        if tracker:
            tracker.update(10, "เริ่มโหลดข้อมูล...")
            time.sleep(0.5)  # เพื่อให้เห็น progress
        
        # Import และเรียกใช้ฟังก์ชัน
        import importlib
        func_module = importlib.import_module(f"functions.{func_name.lower()}")
        func = getattr(func_module, func_name)
        
        if tracker:
            tracker.update(30, "กำลังประมวลผลข้อมูล...")
            time.sleep(0.5)
        
        # เรียกใช้ฟังก์ชัน
        if isinstance(file_path, list):
            if tracker:
                tracker.update(50, f"ประมวลผลไฟล์ทั้งหมด {len(file_path)} ไฟล์...")
            result = [func(f, temp_root) for f in file_path]
        else:
            if func_name in ["DA_AUTO_UPH", "PNP_AUTO_UPH", "WB_AUTO_UPH"]:
                if tracker:
                    tracker.update(60, "คำนวณข้อมูล UPH...")
                result = func(file_path, temp_root, start_date, end_date)
            else:
                result = func(file_path, temp_root)
        
        if tracker:
            tracker.update(90, "กำลังสร้างไฟล์ผลลัพธ์...")
            time.sleep(0.5)
            tracker.complete("ประมวลผลเสร็จสิ้น!")
        
        return result
        
    except Exception as e:
        if tracker:
            tracker.error_occurred(str(e))
        raise e

# === WB_AUTO_UPH Function (ใช้ผ่าน /function แบบปกติ) ===

@app.route("/download_result")
def download_result():
    # สมมติไฟล์ผลลัพธ์ถูกสร้างไว้ใน session["export_file_path"]
    export_file_path = session.get("export_file_path")
    if not export_file_path or not os.path.exists(export_file_path):
        flash("ไม่พบไฟล์สำหรับดาวน์โหลด", "error")
        return redirect(url_for("result"))
    return send_file(export_file_path, as_attachment=True)

if __name__ == "__main__":
    ip = socket.gethostbyname(socket.gethostname())
    print(f"\n✅ Flask app is running on: http://{ip}:8080\n(เปิดจากเครื่องอื่นในเครือข่ายได้ด้วย IP นี้)\n")
    app.run(debug=True, host='0.0.0.0', port=8080)

# ===== Version Information =====
# version 3.0 - Fully Refactored with Service Classes and Modern Architecture

