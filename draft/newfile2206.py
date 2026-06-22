from playwright.sync_api import sync_playwright
import pandas as pd
import re
import time

BASE_URL = "https://infocom.vn/nganh-nghe/quan-ca-phe-do-uong"

def clean_name(text):
    if not text:
        return ""

    text = re.sub(r"^\d+\s*", "", text)
    return text.strip()

def scrape_infocom(start_page=1, end_page=100):

    all_data = []

    with sync_playwright() as p:

        browser = p.chromium.launch(
            headless=True
        )

        page = browser.new_page()

        for page_num in range(start_page, end_page + 1):

            url = f"{BASE_URL}/trang-{page_num}"

            print(f"\nTrang {page_num}")

            try:

                page.goto(
                    url,
                    wait_until="domcontentloaded",
                    timeout=120000
                )

                page.wait_for_timeout(2000)

            except Exception as e:

                print("Lỗi mở trang:", e)
                continue

            cards = page.locator(
                "div.mb-4.bg-white.shadow.rounded-lg.p-4"
            )

            count = cards.count()

            print("Số doanh nghiệp:", count)

            for i in range(count):

                try:

                    card = cards.nth(i)

                    # =====================
                    # TÊN
                    # =====================

                    name = ""

                    name_el = card.locator("h2 a")

                    if name_el.count():
                        name = clean_name(
                            name_el.inner_text()
                        )

                    # =====================
                    # MST
                    # =====================

                    tax_code = ""

                    tax_el = card.locator(
                        "li:has(i.fa-id-card) span.font-medium"
                    )

                    if tax_el.count():
                        tax_code = tax_el.inner_text().strip()

                    # =====================
                    # ĐỊA CHỈ
                    # =====================

                    address = ""

                    addr_el = card.locator(
                        "li:has(i.fa-map-marker-alt) span.flex-1"
                    )

                    if addr_el.count():
                        address = addr_el.inner_text().strip()

                    # =====================
                    # ĐIỆN THOẠI
                    # =====================

                    phone = ""

                    phone_el = card.locator(
                        "li:has(i.fa-phone-alt)"
                    )

                    if phone_el.count():
                        phone = phone_el.inner_text().strip()

                    # =====================
                    # EMAIL
                    # =====================

                    email = ""

                    email_el = card.locator(
                        "li:has(i.fa-envelope)"
                    )

                    if email_el.count():
                        email = email_el.inner_text().strip()

                    # =====================
                    # LINK CHI TIẾT
                    # =====================

                    detail_url = ""

                    if name_el.count():
                        href = name_el.get_attribute("href")

                        if href:
                            detail_url = (
                                "https://infocom.vn" + href
                            )

                    all_data.append({
                        "Tên công ty": name,
                        "Mã số thuế": tax_code,
                        "Địa chỉ": address,
                        "Điện thoại": phone,
                        "Email": email,
                        "Link": detail_url
                    })

                except Exception as e:

                    print("Lỗi:", e)

        browser.close()

    df = pd.DataFrame(all_data)

    df.drop_duplicates(
        subset=["Mã số thuế"],
        inplace=True
    )

    output = "infocom_quan_cafe.xlsx"

    df.to_excel(
        output,
        index=False
    )

    print(f"\nĐã lưu {len(df)} dòng")
    print(output)

if __name__ == "__main__":

    scrape_infocom(
        start_page=1,
        end_page=2
    )