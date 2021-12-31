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
 
Extract data from PDF. You need to get the data from Section A in each PDF.
Then compare the value "Name of this Investment" with the column "Investment Title", 
and the value "Unique Investment Identifier (UII)" with the column "UII".

"""

import time
import os

from RPA.Browser.Selenium import Selenium

import utils


browser = Selenium()
browser.set_download_directory(utils.output_direcrory)


def write_data(filename: str, data: list, sheet_name: str = None) -> None:
    app = utils.get_exel_app(filename, sheet_name)

    for row, item in enumerate(data, 1):
        for column, value in enumerate(item, 1):
            app.set_cell_value(row=row, column=column, value=value)

    app.save_workbook()


def select_agencies() -> list:
    agencies = []

    for element in browser.find_elements("css:div#agency-tiles-widget .col-sm-4"):

        title = browser.find_element("css:div.col-sm-12 span.h4", parent=element).text
        spending = browser.find_element(
            "css:div.col-sm-12 span.h1", parent=element
        ).text

        agencies.append((title, spending))

    return agencies


def select_agencies_links(agencies: list) -> dict:
    links = {}

    container = browser.find_element("css:div#agency-tiles-widget")

    for div in browser.find_elements("css:div.col-sm-12", parent=container):

        title = browser.find_element("css:div:nth-child(2) span.h4", parent=div).text

        if title in agencies:
            href = browser.find_element("tag:a", parent=div).get_attribute("href")
            links[title] = href

    return links


def download_pdf(links: list) -> None:
    for link in links:
        browser.go_to(link)
        browser.wait_until_element_is_visible("css:div#business-case-pdf")
        browser.click_element("css:div#business-case-pdf")

        time.sleep(15)


def select_investments_data() -> list:
    investments = []
    header = browser.find_elements('css:table.datasource-table thead tr[role="row"] th')

    investments.append(
        (
            header[0].text,
            header[1].text,
            header[2].text,
            header[3].text,
            header[4].text,
            header[5].text,
            header[6].text,
        )
    )

    browser.click_element('css:option[value="-1"]')
    browser.wait_until_element_is_not_visible("css:div.loading", timeout=15)

    for table_row in browser.find_elements(
        "css:table#investments-table-object tbody tr"
    ):
        td = browser.find_elements("tag:td", parent=table_row)

        investments.append(
            (
                td[0].text,
                td[1].text,
                td[2].text,
                td[3].text,
                td[4].text,
                td[5].text,
                td[6].text,
            )
        )

    return investments


def download_files() -> None:
    pdf_links = []

    for table_row in browser.find_elements(
        "css:table#investments-table-object tbody tr"
    ):
        td = browser.find_elements("tag:td", parent=table_row)

        try:
            a = browser.find_element("tag:a", parent=td[0])
            pdf_links.append(a.get_attribute("href"))
        except:
            pass

    download_pdf(pdf_links)


def compare_data(investments_data: list) -> None:
    output = []
    files_data = utils.parse_pdf_files()

    for row in investments_data:
        for file_data in files_data:
            if file_data["unique_investment_identifier"] == row[0]:
                output.append(
                    [
                        file_data["unique_investment_identifier"],
                        row[0],
                        file_data["name_of_this_investment"],
                        row[2],
                        "Yes"
                        if file_data["name_of_this_investment"] == row[2]
                        else "No",
                    ]
                )

    output.insert(
        0,
        [
            "Unique investment identifier in file",
            "Unique investment identifier in column",
            "Name of this investment in file",
            "Name of this investment in column",
            "Is same",
        ],
    )

    return output


def store_web_page_content() -> None:
    output_file = os.path.join(utils.output_direcrory, "Agencies.xlsx")

    if not os.path.exists(utils.output_direcrory):
        os.mkdir(utils.output_direcrory)

    browser.open_available_browser("https://itdashboard.gov")
    browser.click_element_when_visible('css:a[href="#home-dive-in"]')
    browser.wait_until_element_is_visible("css:div#agency-tiles-widget")

    agencies_data = select_agencies()
    write_data(output_file, agencies_data, "Agencies")

    agencies = utils.load_agency_names()
    links = select_agencies_links(agencies)

    for agency_name, link in links.items():
        browser.go_to(link)

        browser.wait_until_page_contains_element(
            'css:table[id="investments-table-object"]', timeout=15
        )

        investments_data = select_investments_data()
        write_data(output_file, investments_data, agency_name)

        download_files()

    results = compare_data(investments_data)
    write_data(output_file, results, "Comparison results")


def main() -> None:
    try:
        store_web_page_content()
    finally:
        browser.close_browser()


if __name__ == "__main__":
    main()
