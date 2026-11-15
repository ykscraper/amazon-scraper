
import asyncio
from playwright.async_api import async_playwright
from urllib.parse import urljoin
from openpyxl import Workbook
from openpyxl.styles import Font
import csv

async def main():

    all_products = []
    seen_asins = set()

    async with async_playwright() as p:

        browser = await p.chromium.launch(
            headless=True
        )

        page = await browser.new_page(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36"
        )

        url = "https://www.amazon.com/s?k=apple+watches"
        page_number = 1

        while True:

            print(f"\nPAGE: {page_number}")

            await page.goto(
                url,
                wait_until="domcontentloaded",
                timeout=60000
            )

            await page.wait_for_timeout(3000)

            products = page.locator(
                "div[data-component-type='s-search-result']"
            )

            count = await products.count()
            print("Products Found:", count)

            for i in range(count):

                product = products.nth(i)

                title_locator = product.locator(
                    "div[data-cy='title-recipe'] a h2 span"
                )

                if await title_locator.count() == 0:
                    continue

                title = (await title_locator.first.inner_text()).strip()

                if "watch" not in title.lower():
                    continue

                asin = await product.get_attribute("data-asin")

                if not asin or asin in seen_asins:
                    continue

                seen_asins.add(asin)

                price_locator = product.locator(
                    "span.a-price span.a-offscreen"
                )

                if await price_locator.count() > 0:
                    price = await price_locator.first.inner_text()
                else:
                    price = "Not Available"

                rating_locator = product.locator(
                    "div.a-row.a-size-small span.a-size-small.a-color-base"
                )

                if await rating_locator.count() > 0:
                    rating = await rating_locator.first.inner_text()
                else:
                    rating = "Not Available"

                product_url = f"https://www.amazon.com/dp/{asin}"

                product_data = {
                    "Title": title,
                    "Price": price,
                    "Rating": rating,
                    "Product URL": product_url
                }

                all_products.append(product_data)

                print("Title:", title)
                print("Price:", price)
                print("Rating:", rating)
                print("Product URL:", product_url)
                print("-" * 40)

            next_button = page.locator("a.s-pagination-next")

            if await next_button.count() == 0:
                print("No More Pages!")
                break

            next_url = await next_button.get_attribute("href")

            if not next_url:
                print("No More Pages!")
                break

            url = urljoin(page.url, next_url)
            page_number += 1

        await browser.close()

    with open(
        "watches.csv",
        "w",
        newline="",
        encoding="utf-8-sig"
    ) as file:

        fieldnames = [
            "Title",
            "Price",
            "Rating",
            "Product URL"
        ]

        writer = csv.DictWriter(
            file,
            fieldnames=fieldnames
        )

        writer.writeheader()
        writer.writerows(all_products)

    print("\nCSV File Saved Successfully!")

    wb = Workbook()
    ws = wb.active
    ws.title = "Watches"

    headers = [
        "Title",
        "Price",
        "Rating",
        "Product URL"
    ]

    ws.append(headers)

    for product in all_products:

        ws.append([
            product["Title"],
            product["Price"],
            product["Rating"],
            product["Product URL"]
        ])

        cell = ws.cell(
            row=ws.max_row,
            column=4
        )

        cell.hyperlink = product["Product URL"]

        cell.font = Font(
            color="0563C1",
            underline="single"
        )

    ws.column_dimensions["A"].width = 55
    ws.column_dimensions["B"].width = 15
    ws.column_dimensions["C"].width = 15
    ws.column_dimensions["D"].width = 60

    wb.save("watches.xlsx")

    print("Clickable Excel Saved Successfully!")
    print("\nTotal Unique Products:", len(all_products))


if __name__ == "__main__":
    asyncio.run(main())

