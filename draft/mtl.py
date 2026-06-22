import re
import time
import random
import pandas as pd
from playwright.sync_api import sync_playwright

BASE_URL = "https://masothue.com/tra-cuu-ma-so-thue-doanh-nghiep-moi-thanh-lap/"

OUTPUT_FILE = "restaurant_5610_hcm.xlsx"


# =====================================
# SAFE GOTO
# =====================================

def safe_goto(page, url, retry=3):

    for i in range(retry):

        try:

            page.goto(
                url,
                wait_until="domcontentloaded",
                timeout=60000
            )

            return True

        except Exception as e:

            print(
                f"Retry {i+1}/{retry}: {e}"
            )

            time.sleep(5)

    return False


# =====================================
# STATUS
# =====================================

def is_active_status(status):

    status = status.lower()

    blacklist = [
        "ngừng",
        "không hoạt động",
        "tạm ngừng",
        "chấm dứt"
    ]

    return not any(
        x in status
        for x in blacklist
    )


# =====================================
# EMAIL
# =====================================

def extract_email(body):

    emails = re.findall(
        r'[\w\.-]+@[\w\.-]+\.\w+',
        body
    )

    emails = [
        e for e in emails
        if "masothue.com"
        not in e.lower()
    ]

    return emails[0] if emails else ""


# =====================================
# DETAIL
# =====================================

def get_detail(page, url):

    if not safe_goto(page, url):
        return None

    result = {

        "active_date": "",

        "status": "",

        "phone": "",

        "email": "",

        "is_5610": False
    }

    try:

        body = page.locator(
            "body"
        ).inner_text()

        # ==================
        # ngành 5610
        # ==================

        if (
            "5610" in body
            and
            "Nhà hàng và các dịch vụ ăn uống phục vụ lưu động"
            in body
        ):
            result["is_5610"] = True

        # ==================
        # phone
        # ==================

        try:

            result["phone"] = page.locator(
                "#tel-full"
            ).inner_text().strip()

        except:
            pass

        # ==================
        # status
        # ==================

        try:

            result["status"] = page.locator(
                "#tax-status-html"
            ).inner_text().strip()

        except:
            pass

        # ==================
        # active date
        # ==================

        m = re.search(
            r"Ngày hoạt động\s+(\d{4}-\d{2}-\d{2})",
            body
        )

        if m:

            result["active_date"] = (
                m.group(1)
            )

        result["email"] = (
            extract_email(body)
        )

        return result

    except Exception as e:

        print(
            "DETAIL ERROR:",
            e
        )

        return None


# =====================================
# MAIN
# =====================================

rows = []

with sync_playwright() as p:

    browser = p.chromium.launch(
        headless=False
    )

    page = browser.new_page()

    detail_page = browser.new_page()

    page_num = 1

    while True:

        url = (
            f"{BASE_URL}?page={page_num}"
        )

        print(
            "\n"
            + "=" * 80
        )

        print(
            f"PAGE {page_num}"
        )

        print(url)

        if not safe_goto(
            page,
            url
        ):
            break

        companies = page.locator(
            "div[data-prefetch]"
        )

        total = companies.count()

        print(
            "Companies:",
            total
        )

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
                ).get_attribute(
                    "href"
                )

                address = item.locator(
                    "address"
                ).inner_text().strip()

                # ==================
                # TPHCM
                # ==================

                if (
                    "Hồ Chí Minh"
                    not in address
                ):
                    continue

                full_url = (
                    "https://masothue.com"
                    + href
                )

                print(
                    "Checking:",
                    name
                )

                detail = get_detail(
                    detail_page,
                    full_url
                )

                if detail is None:
                    continue

                # ==================
                # ngành 5610
                # ==================

                if not detail[
                    "is_5610"
                ]:
                    continue

                # ==================
                # trạng thái
                # ==================

                if not is_active_status(
                    detail["status"]
                ):
                    continue

                rows.append({

                    "Tên công ty":
                        name,

                    "Địa chỉ":
                        address,

                    "Ngày hoạt động":
                        detail[
                            "active_date"
                        ],

                    "Tình trạng":
                        detail[
                            "status"
                        ],

                    "Điện thoại":
                        detail[
                            "phone"
                        ],

                    "Email":
                        detail[
                            "email"
                        ],

                    "URL":
                        full_url
                })

                print(
                    "FOUND:",
                    len(rows)
                )

                # save realtime

                pd.DataFrame(
                    rows
                ).to_excel(
                    OUTPUT_FILE,
                    index=False
                )

                time.sleep(
                    random.uniform(
                        2,
                        5
                    )
                )

            except Exception as e:

                print(
                    "ROW ERROR:",
                    e
                )

        page_num += 1

    browser.close()


pd.DataFrame(
    rows
).to_excel(
    OUTPUT_FILE,
    index=False
)

print("\nDONE")

print(
    "TOTAL:",
    len(rows)
)

print(
    "FILE:",
    OUTPUT_FILE
)