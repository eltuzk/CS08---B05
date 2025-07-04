import openpyxl
from datetime import datetime
import re

def calculate_age(birth_date):
    today = datetime.today()
    age = today.year - birth_date.year - ((today.month, today.day) < (birth_date.month, birth_date.day))
    return age

def validate_email(email):
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return bool(re.match(pattern, email))

def validate_phone(phone):
    pattern = r'^\d{10,11}$'
    return bool(re.match(pattern, phone))

def get_status_code(age, job, is_student):
    if is_student:
        return 1
    if age >= 18:
        return 3 if job else 2
    return None

#main
wb = openpyxl.Workbook()
ws = wb.active

headers = ["Ho ten", "Ngay sinh", "Email", "So dien thoai", "Cong viec", "Tinh trang hon nhan", "Tuoi", "Ma trang thai"]
ws.append(headers)

while True:
    name = input("Nhap ho ten (bam enter de thoat): ").strip()
    if not name:
        break

    birth = input("Nhap ngay sinh (dd/mm/yyyy): ").strip()
    try:
        birth = datetime.strptime(birth, "%d/%m/%Y")
        tuoi = calculate_age(birth)
    except ValueError:
        print("Ngay sinh khong hop le. Vui long nhap lai.")
        continue

    email = input("Nhap email: ").strip()
    if not validate_email(email):
        print("Email khong hop le. Vui long nhap lai.")
        continue

    number = input("Nhap so dien thoai: ").strip()
    if not validate_phone(number):
        print("So dien thoai khong hop le. Vui long nhap lai.")
        continue

    job = input("Ban co dang lam viec khong? (y/n): ").strip().lower() == 'y'
    status = input("Tinh trang hon nhan (Doc than/Ket hon): ").strip()
    is_student = input("Ban co phai la sinh vien khong? (y/n): ").strip().lower() == 'y'
    ma_trang_thai = get_status_code(tuoi, job, is_student)

    ws.append([
        name,
        birth.strftime("%d/%m/%Y"),
        email,
        number,
        "Co" if job else "Khong",
        status,
        tuoi,
        ma_trang_thai
    ])

file_name = "users_info.xlsx"
wb.save(file_name)

    