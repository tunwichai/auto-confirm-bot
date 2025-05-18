import gspread
from google.oauth2.service_account import Credentials
import os # สำหรับการจัดการ path ของไฟล์ credentials

# --- การตั้งค่า ---
# ขอบเขตการเข้าถึง API (scope)
# หากต้องการอ่านและเขียน ให้ใช้ทั้งสอง scope นี้
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive.file' # จำเป็นสำหรับ gspread บางฟังก์ชัน
]

# ชื่อไฟล์ JSON ของ service account credentials ที่ดาวน์โหลดมา
# **สำคัญ:** แก้ไขเป็น path ที่ถูกต้องไปยังไฟล์ JSON ของคุณ
# ตัวอย่าง: 'path/to/your/credentials.json' หรือ 'credentials.json' (ถ้าอยู่ในโฟลเดอร์เดียวกับสคริปต์)
CREDENTIALS_FILE = 'credentials.json' # <<<--- แก้ไขตรงนี้

# ชื่อของ Google Sheet ที่คุณต้องการเชื่อมต่อ
# หรือจะใช้ URL ของ Sheet ก็ได้
SHEET_NAME = 'https://docs.google.com/spreadsheets/d/1XJFeDVYTaNHl4Tse4LyC5kq4JqHnW3YOJVjTd22IKgI/edit?gid=0#gid=0' # <<<--- แก้ไขตรงนี้
# SHEET_URL = 'https://docs.google.com/spreadsheets/d/your_sheet_id/edit#gid=0' # ตัวอย่างการใช้ URL

# ชื่อของ Worksheet (แท็บ) ภายใน Spreadsheet (ถ้าไม่ระบุ จะใช้แท็บแรก)
WORKSHEET_NAME = 'Sheet1' # <<<--- แก้ไขตรงนี้ (ถ้าต้องการระบุแท็บ)

def authenticate_google_sheets():
    """
    ยืนยันตัวตนกับ Google Sheets API โดยใช้ service account credentials
    และคืนค่า client object สำหรับการทำงานกับ gspread
    """
    try:
        # ตรวจสอบว่าไฟล์ credentials มีอยู่จริง
        if not os.path.exists(CREDENTIALS_FILE):
            print(f"ข้อผิดพลาด: ไม่พบไฟล์ credentials '{CREDENTIALS_FILE}'")
            print("โปรดตรวจสอบว่าคุณได้ใส่ path ที่ถูกต้อง และไฟล์มีอยู่จริง")
            return None

        creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
        client = gspread.authorize(creds)
        print("เชื่อมต่อ Google Sheets API สำเร็จ!")
        return client
    except Exception as e:
        print(f"เกิดข้อผิดพลาดระหว่างการยืนยันตัวตน: {e}")
        print("ตรวจสอบให้แน่ใจว่า:")
        print("1. ไฟล์ credentials.json ถูกต้องและอยู่ใน path ที่ระบุ")
        print("2. Google Drive API และ Google Sheets API ถูกเปิดใช้งานใน Google Cloud Project")
        print(f"3. Service account ({creds.service_account_email if 'creds' in locals() else 'ไม่ทราบอีเมล'}) ได้รับสิทธิ์เข้าถึง Google Sheet นี้แล้ว")
        return None

def read_data_from_sheet(client, sheet_name_or_url, worksheet_name=None):
    """
    อ่านข้อมูลจาก Google Sheet ที่ระบุ
    """
    if not client:
        return

    try:
        # เปิด Spreadsheet ด้วยชื่อ หรือ URL
        if "docs.google.com/spreadsheets" in sheet_name_or_url:
            spreadsheet = client.open_by_url(sheet_name_or_url)
        else:
            spreadsheet = client.open(sheet_name_or_url)
        print(f"เปิด Spreadsheet '{spreadsheet.title}' สำเร็จ")

        # เลือก Worksheet
        if worksheet_name:
            worksheet = spreadsheet.worksheet(worksheet_name)
            print(f"เลือก Worksheet '{worksheet.title}'")
        else:
            worksheet = spreadsheet.sheet1 # เลือก worksheet แรกโดยอัตโนมัติ
            print(f"เลือก Worksheet แรก '{worksheet.title}' โดยอัตโนมัติ")

        # ตัวอย่างการอ่านข้อมูล
        # 1. อ่านข้อมูลทั้งหมดใน worksheet
        # all_data = worksheet.get_all_records() # คืนค่าเป็น list of dictionaries (ใช้ header row เป็น keys)
        # all_values = worksheet.get_all_values() # คืนค่าเป็น list of lists
        # print("\n--- ข้อมูลทั้งหมด (get_all_records) ---")
        # for row in all_data[:5]: # แสดง 5 แถวแรกเป็นตัวอย่าง
            # print(row)
        # print("\n--- ข้อมูลทั้งหมด (get_all_values) ---")
        # for row in all_values[:5]: # แสดง 5 แถวแรกเป็นตัวอย่าง
        #     print(row)

        # 2. อ่านข้อมูลจาก cell ที่ระบุ (เช่น A1)
        # cell_value = worksheet.acell('A1').value
        # print(f"\n--- ค่าใน Cell A1 --- \n{cell_value}")

        # 3. อ่านข้อมูลจาก range ที่ระบุ (เช่น A2:B3)
        # range_values = worksheet.get('A2:B3')
        # print("\n--- ค่าใน Range A2:B3 ---")
        # for row in range_values:
            # print(row)

        return worksheet # คืนค่า worksheet object เพื่อใช้ต่อ

    except gspread.exceptions.SpreadsheetNotFound:
        print(f"ข้อผิดพลาด: ไม่พบ Spreadsheet ชื่อหรือ URL '{sheet_name_or_url}'")
        print("โปรดตรวจสอบว่าชื่อหรือ URL ถูกต้อง และ service account ได้รับสิทธิ์เข้าถึง")
    except gspread.exceptions.WorksheetNotFound:
        print(f"ข้อผิดพลาด: ไม่พบ Worksheet ชื่อ '{worksheet_name}' ใน Spreadsheet นี้")
    except Exception as e:
        import traceback
        print(f"เกิดข้อผิดพลาดระหว่างการอ่านข้อมูล: {e}")
        traceback.print_exc()

def write_data_to_sheet(worksheet):
    """
    เขียนข้อมูลลงใน Google Sheet (worksheet ที่ระบุ)
    """
    if not worksheet:
        print("ไม่สามารถเขียนข้อมูลได้เนื่องจาก worksheet ไม่พร้อมใช้งาน")
        return

    try:
        # ตัวอย่างการเขียนข้อมูล
        # 1. อัปเดตค่าใน cell ที่ระบุ (เช่น C1)
        worksheet.update_acell('C1', "เขียนจาก Python!")
        print(f"\nอัปเดตค่าใน Cell C1 เป็น 'เขียนจาก Python!'")

        # 2. อัปเดตค่าใน range ที่ระบุ (เช่น D1:E1)
        worksheet.update('D1:E1', [['ข้อมูลใหม่1', 'ข้อมูลใหม่2']])
        print(f"อัปเดตค่าใน Range D1:E1")

        # 3. เพิ่มแถวใหม่ (append row)
        new_row = ['แถวที่1_คอลัมน์A', 'แถวที่1_คอลัมน์B', 'แถวที่1_คอลัมน์C_จากPython']
        worksheet.append_row(new_row)
        print(f"เพิ่มแถวใหม่: {new_row}")

        # 4. เพิ่มหลายแถว (append rows)
        new_rows_data = [
            ['แถวที่2_A', 'แถวที่2_B', 'แถวที่2_C'],
            ['แถวที่3_A', 'แถวที่3_B', 'แถวที่3_C']
        ]
        worksheet.append_rows(new_rows_data)
        print(f"เพิ่มหลายแถวใหม่: {new_rows_data}")

        # 5. ล้างข้อมูลใน worksheet (ใช้ด้วยความระมัดระวัง!)
        # worksheet.clear()
        # print("ล้างข้อมูลทั้งหมดใน worksheet แล้ว")

        print("\nเขียนข้อมูลลง Google Sheet สำเร็จ!")

    except Exception as e:
        import traceback
        print(f"เกิดข้อผิดพลาดระหว่างการอ่านข้อมูล: {e}")
        traceback.print_exc()

# --- ส่วนหลักของโปรแกรม ---
if __name__ == "__main__":
    # 1. ยืนยันตัวตนและสร้าง client
    gspread_client = authenticate_google_sheets()

    if gspread_client:
        # 2. อ่านข้อมูลจาก Sheet
        # คุณสามารถใช้ SHEET_NAME หรือ SHEET_URL
        current_worksheet = read_data_from_sheet(gspread_client, SHEET_NAME, WORKSHEET_NAME)
        # current_worksheet = read_data_from_sheet(gspread_client, SHEET_URL, WORKSHEET_NAME)

        # 3. เขียนข้อมูลลง Sheet (ถ้าการอ่านสำเร็จและได้ worksheet object)
        # if current_worksheet:
        #     write_data_to_sheet(current_worksheet)

        print("\n--- การทำงานเสร็จสิ้น ---")
