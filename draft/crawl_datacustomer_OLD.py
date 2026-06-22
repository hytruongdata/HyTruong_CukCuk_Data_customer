from playwright.sync_api import sync_playwright
from datetime import datetime
from zoneinfo import ZoneInfo
import time
import pathlib
import pandas as pd
import re

# ================= CONFIG =================
# Đã sửa về đúng ngành Quán cà phê, đồ uống
BASE_URL = "https://infocom.vn/nganh-nghe/quan-ca-phe-do-uong"

# Danh sách tỉnh/thành phố để nhận diện nhanh
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

# Load template Excel để lấy danh sách cột chính xác
template_columns = pd.read_excel("./Account_Template.xlsx", engine="openpyxl").columns.tolist()

# ================= CLEANING =================
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

    # Xác định Tỉnh/Thành phố chuẩn định dạng template mẫu
    for tinh in tinh_thanh_list:
        if tinh in address:
            if tinh in ["Hà Nội", "Hồ Chí Minh", "Đà Nẵng", "Hải Phòng", "Cần Thơ"]:
                city = f"Thành phố {tinh}"
            else:
                city = f"Tỉnh {tinh}"
            break

    # Tách chuỗi địa chỉ dựa theo dấu phẩy từ phải qua trái
    address_parts = [part.strip() for part in address.split(",")]

    for i in range(len(address_parts) - 1, -1, -1):
        part = address_parts[i]

        # Giữ nguyên tiền tố Quận, Huyện, Thị xã, Thành phố trực thuộc tỉnh
        if any(keyword in part for keyword in ["Quận", "Huyện", "Thị xã", "Thành phố"]):
            if city and part in city:
                continue
            district = part

        # Giữ nguyên tiền tố Phường, Xã, Thị trấn
        elif any(keyword in part for keyword in ["Phường", "Xã", "Thị trấn"]):
            ward = part

        elif city and (part in city or any(t in part for t in tinh_thanh_list)):
            continue
        else:
            # Đoạn text còn lại ở đầu chuỗi sẽ gom vào trường Số nhà, Đường phố
            street = ", ".join(address_parts[:i + 1])
            break

    return country, city, district, ward, street

# ================= SAVE =================
def save_to_excel(data):
    if not data:
        print("Không có dữ liệu để lưu!")
        return None

    filepath = generate_filename()
    df = pd.DataFrame(data)

    # Đảm bảo giữ đúng chuẩn cấu trúc cột của file Account_Template.xlsx ban đầu
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

                    # 4. ĐIỆN THOẠI (Lấy dữ liệu thực tế thay vì để trống)
                    phone_el = company.query_selector("li:has(i.fa-phone-alt)")
                    phone = phone_el.inner_text().strip() if phone_el else ""
                    phone = re.sub(r"^(Điện thoại|Phone|Hotline|SĐT)\s*:\s*", "", phone, flags=re.IGNORECASE)

                    # 5. EMAIL (Lấy dữ liệu thực tế thay vì để trống)
                    email_el = company.query_selector("li:has(i.fa-envelope)")
                    email = email_el.inner_text().strip() if email_el else ""
                    email = re.sub(r"^(Email|Thư điện tử)\s*:\s*", "", email, flags=re.IGNORECASE)

                    # Phân bổ địa chỉ dựa vào hàm chuẩn hóa template mẫu
                    country, city, district, ward, street = extract_address_components(address)

                    # Gán cấu trúc rỗng theo template
                    row = {col: "" for col in template_columns}
                    
                    # Điền các trường thông tin cào được vào đúng cột tương ứng
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
                    print("Lỗi khi parse dữ liệu doanh nghiệp:", e)

        browser.close()
        
        # Xóa trùng lặp theo Mã số thuế (*) trước khi ghi file
        if all_companies:
            df_final = pd.DataFrame(all_companies)
            df_final.drop_duplicates(subset=["Mã số thuế (*)"], keep="first", inplace=True)
            save_to_excel(df_final.to_dict(orient="records"))
        else:
            print("Không thu thập được dữ liệu hợp lệ nào!")


# ================= MAIN =================
if __name__ == "__main__":
    scrape_infocom(start_page=1, end_page=2)