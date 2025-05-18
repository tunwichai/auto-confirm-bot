from playwright.sync_api import Playwright, sync_playwright, expect
from google_sheets_integration import authenticate_google_sheets, read_data_from_sheet, SHEET_NAME, WORKSHEET_NAME
from my_selectors import (
    LOGIN_BUTTON_XPATH,
    USERNAME_INPUT_XPATH,
    PASSWORD_INPUT_XPATH,
    MENU_ITEM_XPATH,
    TABLE_SELECTOR,
    TABLE_ROW_SELECTOR,
    CELL_SELECTOR,
    GRAB_SINGLE_SELECTOR,
    MODAL_CLOSE_SELECTOR,
)

def run(playwright: Playwright) -> None:
    # อ่านค่าจาก Google Sheet (cell A1)
    gspread_client = authenticate_google_sheets()
    worksheet = None
    if gspread_client:
        worksheet = read_data_from_sheet(gspread_client, SHEET_NAME, WORKSHEET_NAME)

    # เปิดเบราว์เซอร์ Chromium (สามารถเปลี่ยนเป็น firefox หรือ webkit ได้)
    browser = playwright.chromium.launch(
        channel="chrome",
        headless=False,
        args=["--start-maximized", "--window-size=1920,1080"]
    )
    context = browser.new_context(ignore_https_errors=True, viewport={"width":1920, "height":1080})

    # สร้างหน้าเว็บใหม่
    page = context.new_page()

    # ไปยัง URL ของเว็บแอปพลิเคชัน
    page.goto("https://kkmediatech.github.io/logistics-app/")
    print("เปิดหน้าเว็บ https://kkmediatech.github.io/logistics-app/ เรียบร้อยแล้ว")

    # รอให้หน้าเว็บโหลดสมบูรณ์ (เผื่อกรีมี element ที่โหลดช้า)
    page.wait_for_load_state("networkidle")

    # ค้นหาและคลิกปุ่ม "Login"
    try:
        login_button = page.locator(LOGIN_BUTTON_XPATH)
        # กรอก Username
        username_input = page.locator(USERNAME_INPUT_XPATH)
        username_input.fill("0652674703")
        print("กรอก Username: 0652674703")

        # กรอก Password
        password_input = page.locator(PASSWORD_INPUT_XPATH)
        password_input.fill("Fleet2KL@2025")
        print("กรอก Password: Fleet2KL@2025")

        # คลิกปุ่ม Login
        login_button.click()
        print("คลิกปุ่ม Login แล้ว")

        # รอให้หน้าเปลี่ยนหลังล็อกอิน
        page.wait_for_load_state("networkidle")

        # คลิกเมนูที่ต้องการ
        menu_item = page.locator(MENU_ITEM_XPATH)
        menu_item.click()
        print("คลิกเมนูที่ตำแหน่ง li[4] แล้ว")

        # รอให้ตารางโหลด
        page.wait_for_selector(TABLE_SELECTOR, timeout=10000)
        page.wait_for_timeout(2000)  # รอเพิ่มอีก 2 วินาที

        table_rows = page.locator(TABLE_ROW_SELECTOR)
        print(f"จำนวนแถวที่พบ: {table_rows.count()}")
        all_rows = table_rows.all_inner_texts()
        print(f"พบทั้งหมด {len(all_rows)} แถวในตาราง:")
        for i, row_text in enumerate(all_rows, 1):
            print(f"แถวที่ {i}: {row_text}")

        # เตรียมลิสต์คำค้นจาก Google Sheet (A1, A2, A3, A4)
        search_values = []
        if worksheet:
            for row in range(1, 6):  # A1 ถึง A4
                value = worksheet.acell(f'A{row}').value
                if value:
                    search_values.append(value)
        if not search_values:
            print("ไม่พบคำค้นใน A1-A4")
            return

        search_index = 0
        while True:
            cell_value = search_values[search_index]
            print(f"เริ่มค้นหาด้วยค่า: {cell_value}")
            found = False
            for i in range(table_rows.count()):
                row = table_rows.nth(i)
                cells = row.locator(CELL_SELECTOR)
                for j in range(cells.count()):
                    cell_text = cells.nth(j).inner_text().strip()
                    if cell_text == cell_value:
                        print(f"พบข้อมูลในตารางที่ตรงกับค่า: {cell_value} (แถวที่ {i+1}, คอลัมน์ที่ {j+1})")
                        # คลิกปุ่ม grab-single ในแถวเดียวกัน
                        grab_button = row.locator(GRAB_SINGLE_SELECTOR)
                        if grab_button.count() > 0:
                            grab_button.first.click()
                            print(f"คลิกปุ่ม grab-single ในแถวที่ {i+1} เรียบร้อยแล้ว")
                            # รอ modal ปรากฏแล้วค้างไว้ 5 วินาทีก่อนปิด
                            page.wait_for_selector(MODAL_CLOSE_SELECTOR, timeout=3000)
                            page.wait_for_timeout(3000)  # ค้าง modal ไว้ 5 วินาที
                            close_button = page.locator(MODAL_CLOSE_SELECTOR)
                            if close_button.count() > 0:
                                close_button.first.click()
                                print("คลิกปุ่มปิด modal (el-dialog__close) เรียบร้อยแล้ว")
                            else:
                                print("ไม่พบปุ่มปิด modal (el-dialog__close)")
                        else:
                            print(f"ไม่พบปุ่ม grab-single ในแถวที่ {i+1}")
                        found = True
                        break
                if found:
                    break
            if found:
                # หลัง modal ปิดแล้ว ให้หยุดค้นหาชั่วคราว (รอ modal ปิดจริง)
                page.wait_for_timeout(1000)
                # ไปยังคำค้นถัดไป (หรือวนกลับไป A1)
                search_index = (search_index + 1) % len(search_values)
            else:
                # ถ้าไม่เจอ ให้ไปยังคำค้นถัดไป (หรือวนกลับไป A1)
                search_index = (search_index + 1) % len(search_values)
            # หมุนวนไปเรื่อยๆ (ถ้าต้องการหยุดหลังครบ 1 รอบ ให้เพิ่ม break ได้)

    except Exception as e:
        print(f"ไม่พบปุ่ม Login หรือเกิดข้อผิดพลาด: {e}")
        return

    input("กด Enter เพื่อปิดโปรแกรม...")  # เพิ่มบรรทัดนี้

with sync_playwright() as playwright:
    run(playwright)
