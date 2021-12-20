import os

from RPA.Excel.Files import Files


root = os.path.dirname(__file__)
output_direcrory = os.path.join(root, "output")


def load_links():
    with open(os.path.join(root, "links.txt")) as file:
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
