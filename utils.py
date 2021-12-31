import os
import re
import glob

from RPA.Excel.Files import Files
from RPA.PDF import PDF


pdf = PDF()

root = os.path.dirname(__file__)
output_direcrory = os.path.join(root, "output")


def load_agency_names():
    with open(os.path.join(root, "agencies.txt")) as file:
        return file.read().split("\n")


def get_exel_app(filename, sheet_name):
    app = Files()

    if not os.path.exists(filename):
        app.create_workbook(filename)
    else:
        app.open_workbook(filename)

    if sheet_name:
        if not app.worksheet_exists(sheet_name):
            app.create_worksheet(sheet_name)

        app.set_active_worksheet(sheet_name)

    return app


def get_document_data(filename: str):
    first_page = pdf.get_text_from_pdf(filename)[1]

    date_investment_first_submitted = re.compile(
        "Date Investment First Submitted:\s*(\d{4}-\d{2}-\d{2})"
    )
    date_of_last_change_to_activities = re.compile(
        "Date of Last Change to Activities:\s*(\d{4}-\d{2}-\d{2})"
    )
    date_of_last_investment_detail_update = re.compile(
        "Date of Last Investment Detail Update:\s*(\d{4}-\d{2}-\d{2})"
    )
    date_of_last_business_case_update = re.compile(
        "Date of Last Business Case Update:\s*(\d{4}-\d{2}-\d{2})"
    )
    date_of_last_revision = re.compile("Date of Last Revision:\s*(\d{4}-\d{2}-\d{2})")

    agency = re.compile("Agency:\s?(\d{3}\s+-\s+\w*\s*\w*\s*\w*)")
    bureau = re.compile("Bureau:\s?(\d{2}\s+-\s+\w*-?\w*\s?[A-Za-z]*)")
    name_of_this_investment = re.compile("Name of this Investment:\s?([\s A-Za-z]*)")
    unique_investment_identifier = re.compile(
        "Unique Investment Identifier \(UII\):\s+(\d{3}-\d{9})"
    )

    return {
        "date_investment_first_submitted": re.search(
            date_investment_first_submitted, first_page
        ).group(1),
        "date_of_last_change_to_activities": re.search(
            date_of_last_change_to_activities, first_page
        ).group(1),
        "date_of_last_investment_detail_update": re.search(
            date_of_last_investment_detail_update, first_page
        ).group(1),
        "date_of_last_business_case_update": re.search(
            date_of_last_business_case_update, first_page
        ).group(1),
        "date_of_last_revision": re.search(date_of_last_revision, first_page).group(),
        "agency": re.search(agency, first_page).group(1),
        "bureau": re.search(bureau, first_page).group(1),
        "name_of_this_investment": re.search(name_of_this_investment, first_page).group(
            1
        ),
        "unique_investment_identifier": re.search(
            unique_investment_identifier, first_page
        ).group(1),
    }


def get_dowloaded_pdf_files():
    return glob.glob("./output/*.pdf")


def parse_pdf_files():
    files_data = []

    for filename in get_dowloaded_pdf_files():
        files_data.append(get_document_data(filename))

    return files_data
