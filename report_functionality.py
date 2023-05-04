import os

from PyQt5.QtWidgets import QMessageBox, QTableWidget, QSpacerItem, QSizePolicy, QGridLayout

from sub.DB_Table_Export import REPORT_TYPES
from sub.DB_Table_Export.DBExport import DatabaseExport
from sub.DB_Table_Export.ReportPopUp import ReportPopup


def report_functionality(table: QTableWidget, export_name: str, report_type: REPORT_TYPES):
    template = None
    if report_type == REPORT_TYPES.REPORT_TABLE:
        template = "report_template_files/TEMPLATE_TABLE_REPORT.html"
    elif report_type == REPORT_TYPES.REPORT_WEEKPLAN:
        template = "report_template_files/TEMPLATE_WEEKPLAN_REPORT.html"

    # Create and show a popup window for the report options
    popup = ReportPopup()
    popup.exec_()
    # Get the result dictionary from the popup window
    result = popup.result

    # Set the download path for the report files
    download_path = os.path.join(os.path.expanduser('~'), "Downloads", export_name)

    # Create a DatabaseExport object with the template, title and file names
    dbExp = DatabaseExport(template, export_name,
                           download_path + ".html", download_path + ".pdf")
    # Get the table headers and rows from the table widget
    headers = [table.horizontalHeaderItem(i).text() for i in range(table.columnCount())]
    rows = [[table.item(row, col).text() for col in range(table.columnCount())] for row in
            range(table.rowCount())]

    # Create an HTML file from the template, headers and rows
    html_filename = dbExp.create_html(headers, rows, open_file=result['html'], save_file=result['save'])
    # Convert the HTML file to a PDF file with a given scale factor
    if result['pdf']:
        pdf_filename = dbExp.convert_html_to_pdf(scale=0.7, open_file=result['pdf'], save_file=result['save'])

    msg = QMessageBox()
    msg.setWindowTitle("Report Successful")
    msg.setIcon(QMessageBox.Information)
    horizontalSpacer = QSpacerItem(600, 0, QSizePolicy.Minimum, QSizePolicy.Expanding)

    if result['save']:
        msg.setText("Successfully created and saved following reports:")
    else:
        msg.setText(f"Successfully created following reports: ")
    report = "<html><ul>"
    if result['html']:
        report += f"<li>HTML: <br>{html_filename}</li>"
    if result['pdf']:
        report += f"<li>PDF: <br>{pdf_filename}</li>"
    report += "</ul></html>"
    msg.setInformativeText(report)

    layout = msg.layout()
    layout.addItem(horizontalSpacer, layout.rowCount(), 0, 1, layout.columnCount())

    x = msg.exec_()


