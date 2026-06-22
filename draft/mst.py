from playwright.sync_api import sync_playwright

URL = "https://masothue.com/tra-cuu-ma-so-thue-theo-nganh-nghe/nha-hang-va-cac-dich-vu-an-uong-phuc-vu-luu-dong-5610"

with sync_playwright() as p:

    browser = p.chromium.launch(
        headless=False
    )

    page = browser.new_page()

    print("Loading...")

    page.goto(
        URL,
        timeout=120000,
        wait_until="networkidle"
    )

    page.wait_for_timeout(5000)

    print("=" * 50)
    print("TITLE")
    print("=" * 50)

    print(page.title())

    print("=" * 50)
    print("URL")
    print("=" * 50)

    print(page.url)

    print("=" * 50)
    print("ALL LINKS")
    print("=" * 50)

    links = page.locator("a").evaluate_all("""
        els => els.map(e => ({
            text: e.innerText,
            href: e.href
        }))
    """)

    print(f"TOTAL LINKS: {len(links)}")

    for item in links[:100]:
        print(item)

    print("=" * 50)
    print("BODY SAMPLE")
    print("=" * 50)

    body = page.locator("body").inner_text()

    print(body[:5000])

    with open(
        "masothue_page.html",
        "w",
        encoding="utf-8"
    ) as f:
        f.write(page.content())

    print("\nSaved HTML -> masothue_page.html")

    input("\nPress Enter to close browser...")

    browser.close()