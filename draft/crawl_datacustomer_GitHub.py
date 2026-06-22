from playwright.sync_api import sync_playwright
from datetime import datetime
from zoneinfo import ZoneInfo
import time
import pathlib
import pandas as pd
import re

# URL của trang Infocom
BASE_URL = "https://infocom.vn/ma-nganh-nghe/5610/trang-1"

# Danh sách tỉnh/thành phố ở Việt Nam
tinh_thanh_list = [
    "Hà Nội", "Hồ Chí Minh", "Đà Nẵng", "Hải Phòng", "Cần Thơ", "An Giang", "Bà Rịa - Vũng Tàu", "Bắc Giang", "Bắc Kạn", "Bạc Liêu", "Bắc Ninh", "Bến Tre", "Bình Định", "Bình Dương", "Bình Phước", "Bình Thuận", "Cà Mau", "Cao Bằng", "Đắk Lắk", "Đắk Nông", "Điện Biên", "Đồng Nai", "Đồng Tháp", "Gia Lai", "Hà Giang", "Hà Nam", "Hà Tĩnh", "Hải Dương", "Hậu Giang", "Hòa Bình", "Hưng Yên", "Khánh Hòa", "Kiên Giang", "Kon Tum", "Lai Châu", "Lâm Đồng", "Lạng Sơn", "Lào Cai", "Long An", "Nam Định", "Nghệ An", "Ninh Bình", "Ninh Thuận", "Phú Thọ", "Phú Yên", "Quảng Bình", "Quảng Nam", "Quảng Ngãi", "Quảng Ninh", "Quảng Trị", "Sóc Trăng", "Sơn La", "Tây Ninh", "Thái Bình", "Thái Nguyên", "Thanh Hóa", "Thừa Thiên Huế", "Tiền Giang", "Trà Vinh", "Tuyên Quang", "Vĩnh Long", "Vĩnh Phúc", "Yên Bái"
]

# Danh sách quận đặc biệt giữ nguyên chữ "Quận"
quan_dac_biet = [f"Quận {i}" for i in range(1, 13)]


# Danh sách cột theo Account_Template.xlsx
template_columns = pd.read_excel("./Account_Template.xlsx", engine="openpyxl").columns.tolist()

# Hàm làm sạch mã số thuế
def clean_tax_code(tax_code):
    return tax_code.replace("Copy", "").strip() if tax_code else ""

# Hàm làm sạch tên công ty
def clean_company_name(name):
    return re.sub(r'^\d+\s+', '', name).strip() if name else ""

# Hàm lấy thời gian hiện tại theo múi giờ Việt Nam
def get_vietnam_time():
    return datetime.now(ZoneInfo("Asia/Bangkok"))

# Hàm tạo tên file với ngày hiện tại
def generate_filename():
    now = get_vietnam_time()
    date_str = now.strftime("%d-%m-%Y")
    base_filename = f"data_customer_{date_str}"
    
    result_dir = pathlib.Path("result/infocom_data")
    result_dir.mkdir(parents=True, exist_ok=True)

    counter = 1
    while (result_dir / f"{base_filename}_{counter}.xlsx").exists():
        counter += 1

    return result_dir / f"{base_filename}_{counter}.xlsx"

# Hàm phân tích địa chỉ thành các phần riêng biệt
def extract_address_components(address):
    if not address:
        return "Việt Nam", "", "", "", ""

    country = "Việt Nam"
    city, district, ward, street = "", "", "", ""

    for tinh in tinh_thanh_list:
        if tinh in address:
            city = tinh
            break

    address_parts = address.split(", ")
    address_parts = [part.strip() for part in address_parts]

    for i in range(len(address_parts) - 1, -1, -1):
        part = address_parts[i]
        if "Quận" in part or "Huyện" in part or "Thành phố" in part:
            if part in quan_dac_biet:
                district = part  # Giữ nguyên chữ "Quận" nếu thuộc danh sách đặc biệt
            else:
                district = part.replace("Quận ", "").replace("Huyện ", "").strip()
        elif "Phường" in part or "Xã" in part:
            ward = part.replace("Phường ", "").replace("Xã ", "").strip()
        elif city and part == city:
            break
        else:
            street = ", ".join(address_parts[:i + 1])

    return country, city, district, ward, street

# Hàm lưu dữ liệu vào file Excel
def save_to_excel(data):
    if not data:
        print("Không có dữ liệu để lưu!")
        return None

    filepath = generate_filename()
    df = pd.DataFrame(data)
    
    # Sắp xếp theo thứ tự cột trong template
    df = df.reindex(columns=template_columns, fill_value="")
    
    df.to_excel(filepath, index=False, engine="openpyxl")
    print(f"Dữ liệu đã được lưu vào: {filepath}")

# Hàm cào dữ liệu từ trang Infocom
def scrape_infocom(num_pages=1):
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context()
        page = context.new_page()

        all_companies = []
        
        for i in range(1, num_pages + 1):
            url = BASE_URL + str(i)
            print(f"Đang truy cập: {url}")
            page.goto(url, timeout=120000)
            page.wait_for_load_state("networkidle")
            
            companies = page.query_selector_all(".main-content-paging")
            
            for company in companies:
                name_element = company.query_selector(".title-lg-content-paging a")
                name = clean_company_name(name_element.text_content().strip()) if name_element else ""
                tax_code_element = company.query_selector(".fa-id-card + span")
                tax_code = clean_tax_code(tax_code_element.text_content().strip()) if tax_code_element else ""
                phone_element = company.query_selector(".fa-phone-alt + a")
                phone = phone_element.text_content().strip() if phone_element else ""
                email_element = company.query_selector(".fa-envelope + a")
                email = email_element.text_content().strip() if email_element else ""
                
                address_element = company.query_selector(".fa-map-marker-alt + span")
                address = address_element.text_content().strip() if address_element else ""
                country, city, district, ward, street = extract_address_components(address)
                
                # Kiểm tra nếu cả Mã số thuế và Tên công ty đều rỗng, bỏ qua công ty này
                if not name and not tax_code:
                    continue
                
                company_data = {col: "" for col in template_columns}
                company_data.update({
                    "Mã số thuế (*)": tax_code,
                    "Tên khách hàng (*)": name,
                    "Điện thoại": phone,
                    "Email": email,
                    "Địa chỉ (Hóa đơn)": address,
                    "Quốc gia (Hóa đơn)": country,
                    "Tỉnh/Thành phố (Hóa đơn)": city,
                    "Quận/Huyện (Hóa đơn)": district,
                    "Phường/Xã (Hóa đơn)": ward,
                    "Số nhà, Đường phố (Hóa đơn)": street
                })
                all_companies.append(company_data)
            
            time.sleep(2)
        
        browser.close()
        save_to_excel(all_companies)

# Hàm main để chạy chương trình
def main():
    num_pages = 10
    scrape_infocom(num_pages)
    
if __name__ == "__main__":
    main()