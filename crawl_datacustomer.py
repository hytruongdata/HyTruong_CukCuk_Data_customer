from playwright.sync_api import sync_playwright
from datetime import datetime
from zoneinfo import ZoneInfo
import time
import pathlib
import pandas as pd
import re

BASE_URL = "https://infocom.vn/nganh-nghe/quan-ca-phe-do-uong"

# Danh sách ánh xạ để nhận diện tỉnh thành
tinh_thanh_mapping = {
    "Hà Nội": "Thành phố Hà Nội",
    "Hồ Chí Minh": "Thành phố Hồ Chí Minh",
    "HCM": "Thành phố Hồ Chí Minh",
    "Đà Nẵng": "Thành phố Đà Nẵng",
    "Hải Phòng": "Thành phố Hải Phòng",
    "Cần Thơ": "Thành phố Cần Thơ",
    "Thừa Thiên Huế": "Tỉnh Thừa Thiên Huế",
    "Huế": "Tỉnh Thừa Thiên Huế", 
    "Bà Rịa - Vũng Tàu": "Tỉnh Bà Rịa - Vũng Tàu",
    "Vũng Tàu": "Tỉnh Bà Rịa - Vũng Tàu",
    "An Giang": "Tỉnh An Giang", "Bắc Giang": "Tỉnh Bắc Giang", "Bắc Kạn": "Tỉnh Bắc Kạn",
    "Bạc Liêu": "Tỉnh Bạc Liêu", "Bắc Ninh": "Tỉnh Bắc Ninh", "Bến Tre": "Tỉnh Bến Tre",
    "Bình Định": "Tỉnh Bình Định", "Bình Dương": "Tỉnh Bình Dương", "Bình Phước": "Tỉnh Bình Phước",
    "Bình Thuận": "Tỉnh Bình Thuận", "Cà Mau": "Tỉnh Cà Mau", "Cao Bằng": "Tỉnh Cao Bằng",
    "Đắk Lắk": "Tỉnh Đắk Lắk", "Đắk Nông": "Tỉnh Đắk Nông", "Điện Biên": "Tỉnh Điện Biên",
    "Đồng Nai": "Tỉnh Đồng Nai", "Đồng Tháp": "Tỉnh Đồng Tháp", "Gia Lai": "Tỉnh Gia Lai",
    "Hà Giang": "Tỉnh Hà Giang", "Hà Nam": "Tỉnh Hà Nam", "Hà Tĩnh": "Tỉnh Hà Tĩnh",
    "Hải Dương": "Tỉnh Hải Dương", "Hậu Giang": "Tỉnh Hậu Giang", "Hòa Bình": "Tỉnh Hòa Bình",
    "Hưng Yên": "Tỉnh Hưng Yên", "Khánh Hòa": "Tỉnh Khánh Hòa", "Kiên Giang": "Tỉnh Kiên Giang",
    "Kon Tum": "Tỉnh Kon Tum", "Lai Châu": "Tỉnh Lai Châu", "Lâm Đồng": "Tỉnh Lâm Đồng",
    "Lạng Sơn": "Tỉnh Lạng Sơn", "Lào Cai": "Tỉnh Lào Cai", "Long An": "Tỉnh Long An",
    "Nam Định": "Tỉnh Nam Định", "Nghệ An": "Tỉnh Nghệ An", "Ninh Bình": "Tỉnh Ninh Bình",
    "Ninh Thuận": "Tỉnh Ninh Thuận", "Phú Thọ": "Tỉnh Phú Thọ", "Phú Yên": "Tỉnh Phú Yên",
    "Quảng Bình": "Tỉnh Quảng Bình", "Quảng Nam": "Tỉnh Quảng Nam", "Quảng Ngãi": "Tỉnh Quảng Ngãi",
    "Quảng Ninh": "Tỉnh Quảng Ninh", "Quảng Trị": "Tỉnh Quảng Trị", "Sóc Trăng": "Tỉnh Sóc Trăng",
    "Sơn La": "Tỉnh Sơn La", "Tây Ninh": "Tỉnh Tây Ninh", "Thái Bình": "Tỉnh Thái Bình",
    "Thái Nguyên": "Tỉnh Thái Nguyên", "Thanh Hóa": "Tỉnh Thanh Hóa", "Tiền Giang": "Tỉnh Tiền Giang",
    "Trà Vinh": "Tỉnh Trà Vinh", "Tuyên Quang": "Tỉnh Tuyên Quang", "Vĩnh Long": "Tỉnh Vĩnh Long",
    "Vĩnh Phúc": "Tỉnh Vĩnh Phúc", "Yên Bái": "Tỉnh Yên Bái"
}

# Load template Excel
template_columns = pd.read_excel("./Account_Template.xlsx", engine="openpyxl").columns.tolist()

# ====== CLEANING =======
def clean_tax_code(tax_code):
    if not tax_code:
        return ""
    return tax_code.replace("Copy", "").strip()

def clean_company_name(name):
    if not name:
        return ""
    return re.sub(r'^\d+\s*[\.\-]?\s*', '', name).strip()

def get_vietnam_time():
    return datetime.now(ZoneInfo("Asia/Bangkok"))

def generate_filename():
    now = get_vietnam_time()
    date_str = now.strftime("%d-%m-%Y")
    base_filename = f"data_customer_{date_str}"
    
    result_dir = pathlib.Path("result/infocom_datanew")
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

    # 1. Nhận diện Tỉnh/Thành phố thông qua mapping từ khóa đầy đủ/vắn tắt
    for key, value in tinh_thanh_mapping.items():
        if key in address:
            city = value
            break

    # Tách các thành phần bằng dấu phẩy
    address_parts = [part.strip() for part in address.split(",")]

    # Quét ngược từ phải qua trái để phân tách các cấp hành chính
    for i in range(len(address_parts) - 1, -1, -1):
        part = address_parts[i]
        
        # Bỏ qua nếu phần tử trùng với Tỉnh/Thành phố lớn đã nhận diện ở trên
        if city and (part in city or any(k in part for k in tinh_thanh_mapping.keys())):
            continue

        # 2. Xử lý cấp Quận/Huyện/Thị xã/Thành phố thuộc tỉnh (Hỗ trợ viết tắt Q., H., Tx., TP.)
        match_district = re.search(r'^(Quận|Q\.?|Huyện|H\.?|Thị xã|Tx\.?|Thành phố|TP\.?)\s*(.*)$', part, re.IGNORECASE)
        if match_district and not district:
            prefix = match_district.group(1).lower()
            name_part = match_district.group(2).strip()
            
            if 'q' in prefix:
                district = f"Quận {name_part}"
            elif 'h' in prefix:
                district = f"Huyện {name_part}"
            elif 't' in prefix:
                if 'x' in prefix:
                    district = f"Thị xã {name_part}"
                else:
                    district = f"Thành phố {name_part}" # Sẽ chuyển "TP. Huế" thành "Thành phố Huế" chuẩn template
            continue

        # 3. Xử lý cấp Phường/Xã/Thị trấn (Hỗ trợ viết tắt P., X., Tt.)
        match_ward = re.search(r'^(Phường|P\.?|Xã|X\.?|Thị trấn|Tt\.?)\s*(.*)$', part, re.IGNORECASE)
        if match_ward and not ward:
            prefix = match_ward.group(1).lower()
            name_part = match_ward.group(2).strip()
            
            if 'p' in prefix:
                ward = f"Phường {name_part}"
            elif 'x' in prefix:
                ward = f"Xã {name_part}"
            elif 't' in prefix:
                ward = f"Thị trấn {name_part}"
            continue

        # 4. Gom phần còn lại ở phía trước làm Số nhà, Đường phố
        street = ", ".join(address_parts[:i + 1])
        break

    return country, city, district, ward, street

# ===== SAVE =========
def save_to_excel(data):
    if not data:
        print("Không có dữ liệu để lưu!")
        return None

    filepath = generate_filename()
    df = pd.DataFrame(data)

    df = df.reindex(columns=template_columns, fill_value="")
    df.to_excel(filepath, index=False, engine="openpyxl")
    print(f"Đã lưu thành công: {filepath}")

# ================= SCRAPER =================
def scrape_infocom(start_page=1, end_page=1):
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        page = context.new_page()

        all_companies = []

        for i in range(start_page, end_page + 1):
            url = BASE_URL if i == 1 else f"{BASE_URL}/trang-{i}"

            print(f"Đang crawl trang {i}: {url}")
            try:
                page.goto(url, timeout=120000)
                page.wait_for_load_state("networkidle")
                time.sleep(1.5)
            except Exception as e:
                print(f"Lỗi khi tải trang {i}: {e}")
                continue

            companies = page.query_selector_all("div.mb-4.bg-white.shadow.rounded-lg.p-4")

            if not companies:
                print(f"Trang {i} không có dữ liệu hoặc sai selector → dừng")
                break

            for company in companies:
                try:
                    # 1. TÊN CÔNG TY
                    name_el = company.query_selector("h2 a")
                    name = clean_company_name(name_el.inner_text()) if name_el else ""

                    # 2. MÃ SỐ THUẾ
                    tax_el = company.query_selector("li:has(i.fa-id-card) span.font-medium")
                    tax_code = clean_tax_code(tax_el.inner_text()) if tax_el else ""

                    # 3. ĐỊA CHỈ
                    address_el = company.query_selector("li:has(i.fa-map-marker-alt) span.flex-1")
                    address = address_el.inner_text().strip() if address_el else ""

                    if not name and not tax_code:
                        continue

                    # 4. ĐIỆN THOẠI
                    phone_el = company.query_selector("li:has(i.fa-phone-alt)")
                    phone = phone_el.inner_text().strip() if phone_el else ""
                    phone = re.sub(r"^(Điện thoại|Phone|Hotline|SĐT)\s*:\s*", "", phone, flags=re.IGNORECASE)

                    # 5. EMAIL
                    email_el = company.query_selector("li:has(i.fa-envelope)")
                    email = email_el.inner_text().strip() if email_el else ""
                    email = re.sub(r"^(Email|Thư điện tử)\s*:\s*", "", email, flags=re.IGNORECASE)

                    # Phân bổ địa chỉ dựa vào hàm xử lý chữ viết tắt mới
                    country, city, district, ward, street = extract_address_components(address)

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
                        "Số nhà, Đường phố (Hóa đơn)": street,
                    })

                    all_companies.append(row)

                except Exception as e:
                    print("Lỗi khi parse dữ liệu doanh nghiệp:", e)

        browser.close()
        
        if all_companies:
            df_final = pd.DataFrame(all_companies)
            df_final.drop_duplicates(subset=["Mã số thuế (*)"], keep="first", inplace=True)
            save_to_excel(df_final.to_dict(orient="records"))
        else:
            print("Không thu thập được dữ liệu hợp lệ nào!")


# ================= MAIN =================
if __name__ == "__main__":
    scrape_infocom(start_page=1, end_page=10)