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

TODO:

We are looking for people that like going the extra mile if time allows or if your curiosity gets the best of you :солнцезащитные_очки: 
Extract data from PDF. You need to get the data from Section A in each PDF. 
Then compare the value "Name of this Investment" with the column "Investment Title", 
and the value "Unique Investment Identifier (UII)" with the column "UII".

"""

import time
import os

from RPA.Browser.Selenium import Selenium
from selenium.webdriver.common.by import By

import utils


browser = Selenium()
browser.set_download_directory(utils.output_direcrory)


def write_agencies(filename: str, data: dict, sheet_name: str = None) -> None:
    app = utils.get_exel_app(filename, sheet_name)

    for index, item in enumerate(data, 1):
        app.set_cell_value(row=index, column=1, value=item["agency_name"])
        app.set_cell_value(row=index, column=2, value=item["spending"])

    app.save_workbook()


def write_investments_data(filename: str, data: dict, sheet_name: str = None) -> None:
    app = utils.get_exel_app(filename, sheet_name)

    for index, item in enumerate(data, 1):
        app.set_cell_value(row=index, column=1, value=item["uii"])
        app.set_cell_value(row=index, column=2, value=item["investment"])

    app.save_workbook()


def select_agencies():
    agencies = []

    for element in browser.find_elements("css:div#agency-tiles-widget .col-sm-12"):

        title = element.find_element(By.CSS_SELECTOR, "div:nth-child(2) span.h4").text
        spending = element.find_element(
            By.CSS_SELECTOR, "div:nth-child(2) span.h1"
        ).text

        agencies.append({"agency_name": title, "spending": spending})

    return agencies


def download_pdf(links):
    for link in links:
        browser.go_to(link)
        browser.wait_until_element_is_visible("css:div#business-case-pdf")
        browser.click_element("css:div#business-case-pdf")

        time.sleep(15)


def select_investments_data():
    investments = []
    pdf_links = []

    while True:
        for table_row in browser.find_elements("css:table#investments-table-object tr"):
            uii = table_row.find_element(By.CSS_SELECTOR, ":first-child").text

            try:
                a = table_row.find_element(By.CSS_SELECTOR, ":first-child td a")
                pdf_links.append(a.get_attribute("href"))
            except:
                pass

            investment = table_row.find_element(By.CSS_SELECTOR, ":nth-child(4)").text

            if uii and investment:
                investments.append({"uii": uii, "investment": investment})

        if browser.does_page_contain_element("css:a.paginate_button.next.disabled"):
            break
        else:
            browser.click_link("css:a.paginate_button.next")

        while browser.is_element_visible("css:div.loading"):
            time.sleep(0.5)

        time.sleep(0.5)

    download_pdf(pdf_links)

    return investments


def parse_main_page(output_file):
    browser.open_available_browser("https://itdashboard.gov")
    browser.click_element_when_visible('css:a[href="#home-dive-in"]')
    browser.wait_until_element_is_visible("css:div#agency-tiles-widget")

    agencies_data = select_agencies()
    write_agencies(output_file, agencies_data, "Agencies")


def parse_sub_pages(output_file):
    for link in utils.load_links():
        browser.go_to(link)

        browser.wait_until_page_contains_element(
            'css:table[id="investments-table-object"]', timeout=15
        )

        investments_data = select_investments_data()
        write_investments_data(output_file, investments_data, "Individual Investments")


def store_web_page_content():
    output_file = os.path.join(utils.output_direcrory, "Agencies.xlsx")

    if not os.path.exists(utils.output_direcrory):
        os.mkdir(utils.output_direcrory)

    parse_main_page(output_file)
    parse_sub_pages(output_file)


def main():
    try:
        store_web_page_content()
    finally:
        browser.close_browser()


if __name__ == "__main__":
    main()
