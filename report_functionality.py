import fnmatch
import os, holidays
import re
from typing import Dict, Union, List, Optional

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QMessageBox, QTableWidget, QSpacerItem, QSizePolicy

from other.database import Database
from sub.DB_Table_Export import REPORT_TYPES
from sub.DB_Table_Export.DBExport import DatabaseExport
from sub.DB_Table_Export.ReportPopUp import ReportPopup


def report_functionality(
    parent_object: object,
    table: QTableWidget,
    report_name: str,
    report_type: REPORT_TYPES,
    scale: float = None,
    is_landscape: bool = None,
    **kwargs,
):
    template = None
    pdf_filename = ""
    weekdays, year = None, None

    if report_type == REPORT_TYPES.REPORT_TABLE:
        template = "report_template_files/TEMPLATE_TABLE_REPORT.html"

    elif report_type == REPORT_TYPES.REPORT_WEEKPLAN:
        template = "report_template_files/TEMPLATE_WEEKPLAN_REPORT.html"
        is_landscape = True

        # check if the expected kwargs are present
        keys = {"weekdays", "year"}  # your set of keys
        if not keys <= kwargs.keys():  # check if keys is a subset of kwargs.keys()
            # raise an exception if they are not
            raise ValueError("Missing required parameters for the weekplan report type")

        # get the weekdays and colnames from kwargs
        weekdays = kwargs.get("weekdays")
        year = kwargs.get("year")

    # Create and show a popup window for the report options
    popup = ReportPopup()
    popup.exec_()

    # Get the result dictionary from the popup window
    result = popup.result

    # return if result is None which is the case if the cancel button is clicked
    if not result:
        return

    # Set the download path for the report files
    download_path = os.path.join(os.path.expanduser("~"), "Downloads")

    # Create a DatabaseExport object with the template, title and file names
    dbExp = DatabaseExport(template, report_name, download_path, download_path)

    # Get the table headers and rows from the table widget
    headers = __get_headers_from_table_widget__(table)
    rows = __get_rows_from_table_widget__(table, report_type)

    # Colour table
    color_dict = __get_color_dict__(parent_object)
    colors_list = __create_color_list__(table, color_dict)

    # Mark holidays on WEEKPLAN report type
    if report_type == REPORT_TYPES.REPORT_WEEKPLAN:
        __mark_AT_holidays__(rows, colors_list, weekdays, year)

    # Create an HTML file from the template, headers and rows
    html_filename = dbExp.create_html(
        headers,
        rows,
        colors_list,
        open_file=result["html"],
        save_file=(result["html"] and result["save"]),
    )
    # Convert the HTML file to a PDF file with a given scale factor

    if result["pdf"]:
        pdf_filename = dbExp.convert_html_to_pdf(
            is_landscape=is_landscape,
            scale=scale,
            open_file=result["pdf"],
            save_file=(result["pdf"] and result["save"]),
        )
    if result["save"]:
        __success_msgbox__(result, html_filename, pdf_filename)


# Austrian holidays are determined and then marked in red in the weekly plan
def __mark_AT_holidays__(
    rows: List[List[List[str]]],
    color_list: List[List[Optional[str]]],
    weekdays: List,
    year: int,
):
    austria_holidays = holidays.AT(years=int(year))  # get holidays for the given year
    color = "background-color: #D3D3D3;"

    # Check if each weekday is a holiday and if so, make it red and the background gray
    for j, day in enumerate(weekdays):
        if (
            day in austria_holidays
        ):  # use membership test instead of iterating over items
            for i in range(len(rows)):
                color_list[i][j] = color


def __get_color_dict__(parent_object: object):
    db = Database.get_instance(parent_object)
    sql = "SELECT Familienname, Farbe from kat_ausbilder"
    cursor = db.select(sql)
    fetch = cursor.fetchall()
    colors_dict = dict((key, value) for key, value in fetch)
    return colors_dict


def __create_color_list__(table: Union[QTableWidget, QTableWidget], color_dict: dict):
    rows = []
    for r in range(table.rowCount()):
        row = []
        for c in range(table.columnCount()):
            color = ""
            cell_colors = []
            if table.item(r, c) is not None:
                text = table.item(r, c).text().lower()
                # regex for 3 or 6 digit hex color
                found_hex_values = re.findall(r"(#[\da-fA-F]{3,6})", text)
                # search for hex values in text
                for hex_val in found_hex_values:
                    cell_colors.append(hex_val)
                    continue

                # if no hex values found search for instructor names
                if not cell_colors:
                    for color_key in color_dict.keys():
                        tex = text.replace("\n", " ")
                        pattern = re.compile(rf"[\s/\\]{color_key.lower()}[\s/\\]")
                        if pattern.search(tex):
                            text = text.replace(color_key, "")
                            print("uff: ", color_dict.get(color_key, "#663399"))
                            cell_colors.append(
                                color_dict.get(color_key, "#663399")
                            )  # DEBUG color violet: #663399

                # evaluate formatted background color
                if len(cell_colors) < 1:
                    color = f"background-color: #FFFFFF;"
                elif len(cell_colors) < 2:
                    color = f"background-color: {cell_colors[0]};"
                else:
                    color = f"background-image: linear-gradient(to bottom right, {','.join(cell_colors)});"

            row.append(color)
        if any(row):
            rows.append(row)
    return rows


def __get_headers_from_table_widget__(
    table: Union[QTableWidget, QTableWidget]
) -> List[str]:
    headers = [table.horizontalHeaderItem(i).text() for i in range(table.columnCount())]
    return headers


def __get_rows_from_table_widget__(
    table: Union[QTableWidget, QTableWidget], report_type: REPORT_TYPES
) -> List[List[List[str]]]:
    rows = []
    # If the report type is REPORT_TABLE
    if report_type == REPORT_TYPES.REPORT_TABLE:
        # Iterate over each row in the table
        for r in range(table.rowCount()):
            row = []
            # Iterate over each column in the table
            for c in range(table.columnCount()):
                item = table.item(r, c)
                # If item is None
                if not item:
                    row.append("")
                    continue
                # If the item text is not empty
                if item.text() != "":
                    row.append(item.text())
                    continue

                # If the item text is empty
                data = item.data(Qt.ItemDataRole.CheckStateRole)
                # Check if cell is actually empty or a type of checkbox
                row.append(
                    "☐" if data == 0 else "▣" if data == 1 else "☑" if data == 2 else ""
                )
            # If there is any data in the row, append it to rows
            if any(row):
                rows.append(row)

    # If the report type is REPORT_WEEKPLAN
    elif report_type == REPORT_TYPES.REPORT_WEEKPLAN:
        rows = []
        # Iterate over each row in the table
        for r in range(table.rowCount()):
            row = []  # Iterate over each column in the table
            for c in range(table.columnCount()):
                if table.item(r, c) is not None:
                    col = table.item(r, c).text()
                else:
                    col = ""
                # Split the cell content into individual lines on the linebreak and store them
                lines = col.splitlines()
                lines_stripped = [line.strip() for line in lines]
                row.append(lines_stripped)
            # If there is any data in the row, append it to rows
            if any(row):
                rows.append(row)

    return rows


def __success_msgbox__(
    result: Dict[str, bool], html_filename: str, pdf_filename: str
) -> None:
    msg = QMessageBox()
    msg.setWindowTitle("Report Successful")
    msg.setIcon(QMessageBox.Information)
    horizontalSpacer = QSpacerItem(600, 0, QSizePolicy.Minimum, QSizePolicy.Expanding)

    if result["save"]:
        msg.setText("Successfully created and saved following reports:")
    else:
        msg.setText(f"Successfully created following reports: ")
    report = "<html><ul>"
    if result["html"]:
        report += f"<li>HTML: <br>{html_filename}</li>"
    if result["pdf"]:
        report += f"<li>PDF: <br>{pdf_filename}</li>"
    report += "</ul></html>"
    msg.setInformativeText(report)

    layout = msg.layout()
    layout.addItem(horizontalSpacer, layout.rowCount(), 0, 1, layout.columnCount())

    x = msg.exec_()
