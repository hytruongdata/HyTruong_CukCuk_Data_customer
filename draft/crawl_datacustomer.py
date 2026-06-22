from playwright.sync_api import sync_playwright
from datetime import datetime
from zoneinfo import ZoneInfo
import time
import pathlib
import pandas as pd
import re

# ================= CONFIG =================
BASE_URL = "https://infocom.vn/nganh-nghe/du-lich-khach-san-nha-hang-giai-tri"

# Danh sách tỉnh/thành phố
tinh_thanh_list = [
    "Hà Nội", "Hồ Chí Minh", "Đà Nẵng", "Hải Phòng", "Cần Thơ", "An Giang",
    "Bà Rịa - Vũng Tàu", "Bắc Giang", "Bắc Kạn", "Bạc Liêu", "Bắc Ninh",
    "Bến Tre", "Bình Định", "Bình Dương", "Bình Phước", "Bình Thuận",
    "Cà Mau", "Cao Bằng", "Đắk Lắk", "Đắk Nông", "Điện Biên", "Đồng Nai",
    "Đồng Tháp", "Gia Lai", "Hà Giang", "Hà Nam", "Hà Tĩnh", "Hải Dương",
    "Hậu Giang", "Hòa Bình", "Hưng Yên", "Khánh Hòa", "Kiên Giang",
    "Kon Tum", "Lai Châu", "Lâm Đồng", "Lạng Sơn", "Lào Cai", "Long An",
    "Nam Định", "Nghệ An", "Ninh Bình", "Ninh Thuận", "Phú Thọ", "Phú Yên",
    "Quảng Bình", "Quảng Nam", "Quảng Ngãi", "Quảng Ninh", "Quảng Trị",
    "Sóc Trăng", "Sơn La", "Tây Ninh", "Thái Bình", "Thái Nguyên",
    "Thanh Hóa", "Thừa Thiên Huế", "Tiền Giang", "Trà Vinh", "Tuyên Quang",
    "Vĩnh Long", "Vĩnh Phúc", "Yên Bái"
]

quan_dac_biet = [f"Quận {i}" for i in range(1, 13)]

# Load template Excel
template_columns = pd.read_excel("./Account_Template.xlsx", engine="openpyxl").columns.tolist()

# ================= CLEANING =================
def clean_tax_code(tax_code):
    return tax_code.replace("Copy", "").strip() if tax_code else ""

def clean_company_name(name):
    return re.sub(r'^\d+\s+', '', name).strip() if name else ""

def get_vietnam_time():
    return datetime.now(ZoneInfo("Asia/Bangkok"))

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

# ================= ADDRESS PARSER =================
def extract_address_components(address):
    if not address:
        return "Việt Nam", "", "", "", ""

    country = "Việt Nam"
    city, district, ward, street = "", "", "", ""

    for tinh in tinh_thanh_list:
        if tinh in address:
            city = tinh
            break

    address_parts = [part.strip() for part in address.split(",")]

    for i in range(len(address_parts) - 1, -1, -1):
        part = address_parts[i]

        if "Quận" in part or "Huyện" in part or "Thành phố" in part:
            if part in quan_dac_biet:
                district = part
            else:
                district = part.replace("Quận ", "").replace("Huyện ", "").strip()

        elif "Phường" in part or "Xã" in part:
            ward = part.replace("Phường ", "").replace("Xã ", "").strip()

        elif city and part == city:
            break
        else:
            street = ", ".join(address_parts[:i + 1])

    return country, city, district, ward, street

# ================= SAVE =================
def save_to_excel(data):
    if not data:
        print("Không có dữ liệu để lưu!")
        return None

    filepath = generate_filename()
    df = pd.DataFrame(data)

    df = df.reindex(columns=template_columns, fill_value="")

    df.to_excel(filepath, index=False, engine="openpyxl")
    print(f"Đã lưu: {filepath}")

# ================= SCRAPER =================
def scrape_infocom(start_page=1, end_page=1):
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()

        all_companies = []

        for i in range(start_page, end_page + 1):
            url = BASE_URL if i == 1 else f"{BASE_URL}/trang-{i}"

            print(f"Đang crawl trang {i}: {url}")
            page.goto(url, timeout=120000)
            page.wait_for_load_state("networkidle")
            time.sleep(1)

            companies = page.query_selector_all("div.mb-4.bg-white.shadow")

            # Nếu trang không có data → dừng sớm
            if not companies:
                print(f"Trang {i} không có dữ liệu → dừng")
                break

            for company in companies:
                try:
                    name_el = company.query_selector("h2 a")
                    name = clean_company_name(name_el.inner_text().strip()) if name_el else ""

                    tax_el = company.query_selector("li:has(i.fa-id-card) span")
                    tax_code = clean_tax_code(tax_el.inner_text().strip()) if tax_el else ""

                    address_el = company.query_selector("li:has(i.fa-map-marker-alt) span")
                    address = address_el.inner_text().strip() if address_el else ""

                    phone = ""
                    email = ""

                    country, city, district, ward, street = extract_address_components(address)

                    if not name and not tax_code:
                        continue

                    row = {col: "" for col in template_columns}
                    row.update({
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

                    all_companies.append(row)

                except Exception as e:
                    print("Lỗi khi parse company:", e)

        browser.close()
        save_to_excel(all_companies)


# ================= MAIN =================
if __name__ == "__main__":
    scrape_infocom(start_page=51, end_page=60)