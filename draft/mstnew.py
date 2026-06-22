import re
import time
import pandas as pd
from playwright.sync_api import sync_playwright

BASE_URL = "https://masothue.com/tra-cuu-ma-so-thue-theo-nganh-nghe/nha-hang-va-cac-dich-vu-an-uong-phuc-vu-luu-dong-5610"

OUTPUT_FILE = "restaurants_hcm_new.xlsx"


def is_active_status(status: str):

    status = status.lower()

    bad_keywords = [
        "ngừng",
        "không hoạt động",
        "tạm ngừng",
        "chấm dứt"
    ]

    return not any(x in status for x in bad_keywords)


def extract_email(body):

    emails = re.findall(
        r'[\w\.-]+@[\w\.-]+\.\w+',
        body
    )

    emails = [
        x for x in emails
        if "masothue.com" not in x.lower()
    ]

    return emails[0] if emails else ""


def get_detail(page, url):

    try:

        page.goto(
            url,
            wait_until="networkidle",
            timeout=120000
        )

        result = {
            "active_date": "",
            "status": "",
            "phone": "",
            "email": ""
        }

        # Phone
        try:
            result["phone"] = page.locator(
                "#tel-full"
            ).inner_text().strip()
        except:
            pass

        # Status
        try:
            result["status"] = page.locator(
                "#tax-status-html"
            ).inner_text().strip()
        except:
            pass

        body = page.locator(
            "body"
        ).inner_text()

        # Ngày hoạt động
        m = re.search(
            r"Ngày hoạt động\s+(\d{4}-\d{2}-\d{2})",
            body
        )

        if m:
            result["active_date"] = m.group(1)

        result["email"] = extract_email(body)

        return result

    except Exception as e:

        print("DETAIL ERROR:", e)

        return {
            "active_date": "",
            "status": "",
            "phone": "",
            "email": ""
        }


rows = []

with sync_playwright() as p:

    browser = p.chromium.launch(
        headless=False
    )

    page = browser.new_page()

    detail_page = browser.new_page()

    page_num = 1

    while True:

        url = f"{BASE_URL}?page={page_num}"

        print("\n" + "=" * 80)
        print(f"PAGE {page_num}")
        print(url)

        try:

            page.goto(
                url,
                wait_until="networkidle",
                timeout=120000
            )

        except Exception as e:

            print("STOP:", e)
            break

        companies = page.locator(
            "div[data-prefetch]"
        )

        total = companies.count()

        print("Companies:", total)

        if total == 0:
            break

        for i in range(total):

            try:

                item = companies.nth(i)

                name = item.locator(
                    "h3 a"
                ).inner_text().strip()

                href = item.locator(
                    "h3 a"
                ).get_attribute("href")

                address = item.locator(
                    "address"
                ).inner_text().strip()

                # Chỉ TP.HCM
                if "Hồ Chí Minh" not in address:
                    continue

                # Loại chi nhánh
                blacklist = [
                    "CHI NHÁNH",
                    "VĂN PHÒNG ĐẠI DIỆN",
                    "ĐỊA ĐIỂM KINH DOANH"
                ]

                if any(
                    x in name.upper()
                    for x in blacklist
                ):
                    continue

                full_url = (
                    "https://masothue.com"
                    + href
                )

                print("Checking:", name)

                detail = get_detail(
                    detail_page,
                    full_url
                )

                active_date = detail[
                    "active_date"
                ]

                status = detail[
                    "status"
                ]

                # Chưa lấy được ngày
                if not active_date:
                    continue

                # Chỉ lấy từ năm 2026
                if active_date < "2026-01-01":
                    continue

                # Chỉ lấy đang hoạt động
                if not is_active_status(
                    status
                ):
                    continue

                rows.append({

                    "Tên công ty":
                        name,

                    "Địa chỉ":
                        address,

                    "Ngày hoạt động":
                        active_date,

                    "Tình trạng":
                        status,

                    "Điện thoại":
                        detail["phone"],

                    "Email":
                        detail["email"],

                    "URL":
                        full_url
                })

                print(
                    "Added:",
                    len(rows)
                )

                time.sleep(0.3)

            except Exception as e:

                print("ROW ERROR:", e)

        page_num += 1

    browser.close()


df = pd.DataFrame(rows)

df.to_excel(
    OUTPUT_FILE,
    index=False
)

print("\nDONE")
print("TOTAL:", len(df))
print("FILE:", OUTPUT_FILE)