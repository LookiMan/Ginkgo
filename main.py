"""
Your challenge is to automate the process of extracting data from itdashboard.gov.
The bot should get a list of agencies and the amount of spending from the main page
Click "DIVE IN" on the homepage to reveal the spend amounts for each agency
Write the amounts to an excel file and call the sheet "Agencies".
Then the bot should select one of the agencies, for example, 
National Science Foundation (this should be configured in a file or on a Robocloud)
Going to the agency page scrape a table with all "Individual Investments" and write it to a new sheet in excel.
If the "UII" column contains a link, open it and download PDF with Business Case (button "Download Business Case PDF")
Your solution should be submitted and tested on Robocloud. 
-> contact Andrew to get an access after the code is ready or you can create your own account for free!
Store downloaded files and Excel sheet to the root of the output folder
"""

import os
import time

from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.HTTP import HTTP


browser = Selenium()
http = HTTP()


def download_file(src, filename):
    headers = {
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.95 Safari/537.36",
        "Cookie": "SSESS04ae24068a2b6b9bb1975f7ad3e4d1c2=RgFK3bynIDSPQ8SIpjFz6CHTbFX94c7u-TvTB4wspe8;has_js=1;wstact=b50a07eb65d8937f411969c0c47e320b3408887d7966a684bde8ce9beda89c2b",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
        "Host": "itdashboard.gov",
        "Accept-Language": "ru-RU,ru;q=0.8,en-US;q=0.5,en;q=0.3",
        "Accept-Encoding": "gzip, deflate, br",
        "Connection": "keep-alive",
        "Referer": f"https://itdashboard.gov/drupal/summary/422/{src}",
        "Upgrade-Insecure-Requests": "1",
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
        "Sec-Fetch-Site": "same-origin",
        "Sec-Fetch-User": "?1",
        "DNT": "1",
        "Sec-GPC": "1",
        "Pragma": "no-cache",
        "Cache-Control": "no-cache",
    }

    url = f"https://itdashboard.gov/api/v1/ITDB2/businesscase/pdf/generate/uii/{src}"

    directory = os.path.dirname(filename)

    if not os.path.exists(directory):
        os.mkdir(directory)

    http.download(url, filename, headers=headers)


def write_exel_worksheet(filename, sheet_name, data):
    app = Files()

    if not os.path.exists(filename):
        app.create_workbook(filename)
    else:
        app.open_workbook(filename)

    if not app.worksheet_exists(sheet_name):
        app.create_worksheet(sheet_name)

    app.set_active_worksheet(sheet_name)

    for index, value in enumerate(data, 1):
        app.set_cell_value(row=index, column=1, value=value[0])
        app.set_cell_value(row=index, column=2, value=value[1])

    app.save_workbook()


def select_agencies():
    agencies = []

    agency_tiles_widget = browser.get_webelement('css:div[id="agency-tiles-widget"]')

    for element in agency_tiles_widget.find_elements_by_css_selector(
        'div[class="col-sm-12"]'
    ):
        title = element.find_element_by_css_selector("div:nth-child(2) span.h4").text
        spending = element.find_element_by_css_selector("div:nth-child(2) span.h1").text

        agencies.append((title, spending))

    return agencies


def select_investments_data():
    investments = []

    investments_table = browser.get_webelement(
        'css:table[id="investments-table-object"]'
    )

    while True:
        for table_row in investments_table.find_elements_by_tag_name("tr"):
            uii = table_row.find_element_by_css_selector(":first-child")

            try:
                uii.find_element_by_tag_name("a")
            except:
                pass
            else:
                download_file(uii.text, f"./downloads/{uii.text}.pdf")

            uii = uii.text
            investment = table_row.find_element_by_css_selector(":nth-child(4)").text

            if uii and investment:
                investments.append((uii, investment))

        try:
            browser.get_webelement('css:a[class="paginate_button next disabled"]')
        except:
            pass
        else:
            break

        browser.click_link('css:a[class="paginate_button next"]')
        time.sleep(10)

    return investments


def store_web_page_content():
    browser.open_available_browser("https://itdashboard.gov")

    button_selector = 'css:a[href="#home-dive-in"]'
    browser.wait_until_page_contains_element(button_selector)
    browser.get_webelement(button_selector).click()
    browser.wait_until_element_is_visible('css:div[id="agency-tiles-widget"]')

    agencies_data = select_agencies()
    write_exel_worksheet("./Agencies.xlsx", "agencies", agencies_data)

    browser.go_to("https://itdashboard.gov/drupal/summary/422")

    browser.wait_until_page_contains_element(
        'css:table[id="investments-table-object"]', timeout=15
    )

    investments_data = select_investments_data()
    write_exel_worksheet("./Agencies.xlsx", "Individual Investments", investments_data)


def main():
    try:
        store_web_page_content()
    finally:
        browser.close_browser()


if __name__ == "__main__":
    main()
