
## คู่มือการใช้งานและการเพิ่มฟังก์ชันใหม่ 

### 1. โครงสร้างโปรเจกต์ที่ควรมี

```
NEW_WEB/
│
├── Webapp/
│   ├── src/
│   │   ├── app.py                # ไฟล์หลักของเว็บแอป
│   │   ├── functions/            # โฟลเดอร์เก็บฟังก์ชันย่อย (แต่ละฟังก์ชัน 1 ไฟล์ .py)
│   │   │   ├── DA_AUTO_UPH.py
│   │   │   ├── PNP_AUTO_UPH.py
│   │   │   └── ...อื่นๆ
│   │   ├── templates/            # โฟลเดอร์เก็บไฟล์ HTML
│   │   │   ├── method.html
│   │   │   ├── function.html
│   │   │   └── ...อื่นๆ
│   │   ├── data_Da/              # โฟลเดอร์เก็บไฟล์ข้อมูล (แต่ละ operation)
│   │   ├── data_PNP/
│   │   ├── data_WB/
│   │   └── data_logview/
│   └── ...
└── README.md
```

**หมายเหตุ:**

- ชื่อโฟลเดอร์และไฟล์ต้องตรงกับที่ระบุในโค้ด เช่น `data_Da`, `data_PNP` ฯลฯ
- ไฟล์ข้อมูลที่ต้องการประมวลผลให้นำไปวางในโฟลเดอร์ที่ตรงกับ operation

---

### 2. การเพิ่มฟังก์ชันใหม่ให้กับระบบ

#### 2.1 สร้างไฟล์ฟังก์ชันใหม่

1. ไปที่โฟลเดอร์ `Webapp/src/functions/`
2. สร้างไฟล์ใหม่ เช่น `MY_FUNCTION.py`
3. ในไฟล์นี้ ต้องมีฟังก์ชันหลักที่ชื่อเดียวกับไฟล์ (ตัวพิมพ์ใหญ่) เช่น
   ```python
   # filepath: Webapp/src/functions/MY_FUNCTION.py
   def MY_FUNCTION(file_path, temp_root, start_date=None, end_date=None):
       # เขียนโค้ดประมวลผลที่ต้องการ
       # ต้อง return path ของไฟล์ผลลัพธ์ (เช่น Excel/CSV) ที่สร้างขึ้น
       return export_file_path
   ```

#### 2.2 เพิ่มชื่อฟังก์ชันใน `app.py`

1. เปิดไฟล์ `Webapp/src/app.py`
2. หา dictionary ที่ชื่อ `OPERATION_FUNCTIONS` (หรือโครงสร้างที่ใช้เก็บ mapping operation กับชื่อฟังก์ชัน)
3. เพิ่มชื่อฟังก์ชันใหม่ใน operation ที่ต้องการ เช่น
   ```python
   OPERATION_FUNCTIONS = {
       "Die Attach": ["DA_AUTO_UPH", "MY_FUNCTION"],
       "Pick & Place": ["PNP_AUTO_UPH"],
       # ...อื่นๆ
   }
   ```

#### 2.3 ไม่ต้องแก้ไขส่วนอื่นใน `app.py` ถ้าฟังก์ชันรับ argument ตามรูปแบบ

---

### 3. การจัดการโฟลเดอร์ข้อมูล

- ไฟล์ข้อมูลแต่ละประเภท (เช่น Die Attach, Pick & Place) ให้นำไปวางในโฟลเดอร์ที่ตรงกับ operation
- ตัวอย่าง: ถ้าจะประมวลผล Die Attach ให้นำไฟล์ไปวางใน `Webapp/src/data_Da/`
- ถ้าเพิ่ม operation ใหม่ ให้สร้างโฟลเดอร์ใหม่ใน `src/` และเพิ่ม mapping ใน `app.py` ด้วย

---

### 4. การรันแอป

1. เปิด terminal ไปที่โฟลเดอร์ `Webapp/src/`
2. ติดตั้ง Python packages ที่จำเป็น (ถ้ายังไม่ติดตั้ง)
   ```sh
   pip install flask pandas openpyxl
   ```
3. รันแอป
   ```sh
   python app.py
   ```
4. เปิดเว็บเบราว์เซอร์ ไปที่ `http://localhost:8080` (หรือ IP ที่แสดงใน terminal)

---

### 5. ข้อควรระวัง

- ฟังก์ชันใหม่ต้อง return path ของไฟล์ผลลัพธ์ที่สร้างขึ้น
- ถ้าเกิด error ให้ดูข้อความใน terminal เพื่อ debug
- ชื่อไฟล์และโฟลเดอร์ต้องตรงกับที่ระบุในโค้ด
- หากเพิ่ม operation ใหม่ ต้องสร้างโฟลเดอร์ข้อมูลและเพิ่ม mapping ใน `app.py` ด้วย

---


### 7. สิ่งที่ต้องรู้และข้อควรระวังเชิงลึก (สำคัญมาก)

1. **ฟังก์ชันใหม่ต้อง return “path ของไฟล์” จริง**
   - ฟังก์ชันใน `src/functions/` ที่จะให้เว็บเรียก ต้อง return path ของไฟล์ที่สร้างจริง (เช่น .xlsx, .csv)
   - ถ้า return เป็น DataFrame หรือ dict ต้องปรับ backend ให้รองรับ หรือแปลงเป็นไฟล์ก่อน

2. **ชื่อฟังก์ชันและชื่อไฟล์ต้องตรงกัน (case-sensitive)**
   - ชื่อฟังก์ชันในไฟล์ .py ต้องตรงกับชื่อที่ใส่ใน `OPERATION_FUNCTIONS` (เช่น DA_AUTO_UPH)
   - ชื่อไฟล์ .py ต้องตรงกับชื่อฟังก์ชัน (เช่น DA_AUTO_UPH.py)

3. **โฟลเดอร์ข้อมูลต้องมีจริง และมีไฟล์อยู่**
   - ถ้าเลือก operation แล้ว dropdown ไม่ขึ้นไฟล์ ให้เช็ค path และชื่อโฟลเดอร์
   - ต้องมีไฟล์ในโฟลเดอร์ เช่น `data_Da/` ไม่ใช่แค่โฟลเดอร์เปล่า

4. **การ mapping operation → โฟลเดอร์**
   - ใน app.py มี mapping ว่า operation ไหนใช้โฟลเดอร์ไหน เช่น
     ```python
     operation_folder_map = {
         "Die Attach": "data_Da",
         "Pick & Place": "data_PNP",
         # ...
     }
     ```
   - ถ้าเพิ่ม operation ใหม่ ต้องเพิ่ม mapping นี้ด้วย

5. **การ debug**
   - ถ้าเกิด error หรือ dropdown ไม่ขึ้น ให้ print log ดูค่าตัวแปร เช่น path, folder_list
   - ดู error ใน terminal จะช่วยบอกปัญหาได้ตรงจุด

6. **การอัปเดต requirements**
   - ถ้าเพิ่มฟังก์ชันที่ใช้ไลบรารีใหม่ (เช่น numpy, matplotlib) ต้อง `pip install` และอัปเดต requirements.txt ด้วย

7. **การใช้งานบน Windows**
   - path ต้องใช้ `os.path.join` เสมอ หลีกเลี่ยง hardcode `/` หรือ `\\`
   - ถ้า path ผิดจะหาไฟล์ไม่เจอ

8. **การใช้งาน session**
   - ข้อมูลไฟล์ที่เลือก, operation, ผลลัพธ์ ฯลฯ จะเก็บใน session ของ Flask
   - ถ้า session หาย (เช่น browser ปิด) อาจต้องเลือกไฟล์ใหม่

9. **การ deploy จริง**
   - ถ้าจะ deploy จริง (ไม่ใช่ dev) ควรใช้ production server เช่น gunicorn + nginx
   - อย่าใช้ debug=True ใน production

10. **การ backup ข้อมูล**
    - ผลลัพธ์จะถูกเก็บในโฟลเดอร์ temp/ ควร backup หรือเคลียร์ไฟล์เก่าเป็นระยะ

---

**อ่านหัวข้อนี้ให้ครบ จะช่วยให้ดูแลและต่อยอดระบบนี้ได้อย่างมั่นใจ**
