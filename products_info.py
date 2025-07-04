import openpyxl
import os

# Hàm kiểm tra và tải/lưu file Excel
def initialize_workbook(excel_file):
    if os.path.exists(excel_file):
        wb = openpyxl.load_workbook(excel_file)
        ws = wb.active
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Product Information"
        headers = ["ID sản phẩm", "Tên sản phẩm", "Phí sản xuất", "Phí vận chuyển",
                   "Tỷ lệ lợi nhuận", "Giá bán sản phẩm"]
        ws.append(headers)
    return wb, ws

# Hàm nhập và xác thực ID sản phẩm
def get_product_id():
    while True:
        product_id = input("Nhập ID sản phẩm: ").strip()
        if product_id:
            return product_id
        print("ID sản phẩm không được để trống! Vui lòng nhập lại.")

# Hàm nhập và xác thực thông tin sản phẩm
def get_product_details():
    product_name = input("Nhập tên sản phẩm: ").strip()
    while True:
        try:
            production_cost = float(input("Nhập phí sản xuất: "))
            shipping_cost = float(input("Nhập phí vận chuyển: "))
            if production_cost >= 0 and shipping_cost >= 0:
                break
            print("Phí phải là số dương! Vui lòng nhập lại.")
        except ValueError:
            print("Vui lòng nhập số hợp lệ!")
    return product_name, production_cost, shipping_cost

# Hàm tính giá bán sản phẩm
def calculate_selling_price(production_cost, shipping_cost, profit_margin, discount, voucher=0):
    base_price = (production_cost + shipping_cost) / (1 - profit_margin)
    if discount == 1:
        selling_price = base_price * (1 - voucher)
    else:
        selling_price = base_price
    return round(selling_price, 2)

# Hàm nhập và xác thực phân loại mặt hàng
def get_discount_info():
    while True:
        try:
            discount = int(input("Phân loại mặt hàng (0: Không khuyến mãi, 1: Có khuyến mãi): "))
            if discount in [0, 1]:
                break
            print("Vui lòng nhập 0 hoặc 1!")
        except ValueError:
            print("Vui lòng nhập số hợp lệ!")
    voucher = 0
    if discount == 1:
        while True:
            try:
                voucher = int(input("Nhập % khuyến mãi (ví dụ: 30, 40, ...): ")) / 100
                if 0 < voucher < 1:
                    break
                print("Phần trăm khuyến mãi phải từ 0 đến 100! Vui lòng nhập lại.")
            except ValueError:
                print("Vui lòng nhập số hợp lệ!")
    return discount, voucher

# Hàm thêm dữ liệu vào worksheet
def add_product_data(ws, product_id, product_name, production_cost, shipping_cost, profit_margin, selling_price):
    ws.append([
        product_id,
        product_name,
        production_cost,
        shipping_cost,
        profit_margin,
        selling_price
    ])

excel_file = "products_info.xlsx"
wb, ws = initialize_workbook (excel_file)

while True: 
    product_id = get_product_id()
    product_name, production_cost, shipping_cost = get_product_details()
    
    while True:
        try:
            profit_margin = float(input("Nhap ty le loi nhuan (%): ")) / 100
            if 0 < profit_margin < 1:
                break
            print("Ty le loi nhuan phai nam trong khoang 0-100%! Vui long nhap lai.")
        except ValueError:
            print("Vui long nhap so hop le!")

    discount, voucher = get_discount_info()

    selling_price = calculate_selling_price(production_cost, shipping_cost, profit_margin, discount, voucher)
    print(f"Gia ban san pham: {selling_price} VND")

    add_product_data(ws, product_id, product_name, production_cost, shipping_cost, profit_margin, selling_price)
    wb.save(excel_file)

    cont = input("Ban co muon nhap san pham khac khong? (y/n): ").strip().lower()
    if cont not in ["y"]:
        break
