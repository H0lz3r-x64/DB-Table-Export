import base64, os, jinja2, mysql.connector
import datetime

from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from webdrivermanager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options


class DatabaseExport:
    def __init__(self, template: str, host: str, user: str, pwd: str, database: str, table: str,
                 output_html="", output_pdf=""):
        # Database Data
        self.host = host
        self.user = user
        self.pwd = pwd
        self.database = database
        self.table = table

        # Filenames
        self.template = template
        if output_html == "":
            self.output_html = os.path.abspath(f"{table}_Export_{datetime.date.today()}.html")
        else:
            self.output_html = os.path.abspath(output_html)
        if output_pdf == "":
            self.output_pdf = os.path.abspath(f"{table}_export_{datetime.date.today()}.pdf")
        else:
            self.output_pdf = os.path.abspath(output_pdf)

        # SQL Variables
        self.mydb = None
        self.mycursor = None

        # Formats
        self.format_dict = {"Legal": (8.5, 14), "legal": (8.5, 14),
                            "Letter": (8.5, 11), "letter": (8.5, 11),
                            "A5": (5.8, 8.3), "a5": (5.8, 8.3),
                            "A4": (8.3, 11.7), "a4": (8.3, 11.7),
                            "A3": (11.7, 16.5), "a3": (11.7, 16.5)}

    def establish_connection(self):
        DB_Dict = {
            'host': self.host,
            'user': self.user,
            'password': self.pwd,
            'database': self.database
        }
        # Trying to establish SQL Connection
        try:
            self.mydb = mysql.connector.connect(**DB_Dict)
            self.mycursor = self.mydb.cursor()
        except mysql.connector.errors.Error:
            print("Error\nConnection to database failed")
            exit(0)

    def close_connection(self):
        if self.mydb.is_connected():
            self.mycursor.close()
            self.mydb.close()
            print("MySQL Verbindung wurde geschlossen")

    @staticmethod
    def render_without_request(template_name, **template_vars):
        # Usage is the same as flask.render_template:
        env = jinja2.Environment(
            loader=jinja2.FileSystemLoader("./")
        )
        template = env.get_template(template_name)
        return template.render(**template_vars)

    def get_headers(self):
        self.establish_connection()
        sql = (f"SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = '{self.database}'"
               f"AND TABLE_NAME = '{self.table}'")
        self.mycursor.execute(sql)
        header = self.mycursor.fetchall()
        # unpack list
        i = 0
        for item in header:
            header[i] = item[0]
            i += 1
        return header

    def get_format_in_inches(self, paper_format: str):
        return self.format_dict[paper_format]

    def create_html(self, display_headers=None, filter_column_name=None, filter_value=None,
                    style='light', open_file=True):
        # mutable handling. Needed to archive optional filter arguments as lists
        if filter_value is None:
            filter_value = []
        if filter_column_name is None:
            filter_column_name = []
        if display_headers is None:
            display_headers = []

        self.establish_connection()

        # --------------- Header --------------
        sql = (f"SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = '{self.database}'"
               f"AND TABLE_NAME = '{self.table}'")
        self.mycursor.execute(sql)
        header = self.mycursor.fetchall()

        # Unpack list
        i = 0
        for item in header:
            header[i] = item[0]
            i += 1

        # List to string and apply header filter
        if len(display_headers) > 0:  # If header filter
            display_headers_string = ', '.join(map(str, display_headers))
        else:                       # If no header filter
            display_headers_string = "*"
            display_headers = header

        # ------------------------- Row Filters -------------------------
        # List to string and apply header filter
        row_filters = str()
        if len(filter_column_name) > 0:  # If header filter
            row_filters += "WHERE "
            for i in range(len(filter_column_name)):

                valueString = ', '.join(map(str, filter_value[i]))

                row_filters += f"{filter_column_name[i]} in ({valueString})"
                if i < len(filter_column_name)-1:
                    row_filters += " AND "
            i += 1

        # --------------------------- SQL Data ---------------------------
        sql = f"Select {display_headers_string} from {self.table} {row_filters}"
        print("sql: ", sql)
        self.mycursor.execute(sql)
        data = self.mycursor.fetchall()

        # Get the right format
        rows = [[str(item) for item in row] for row in data]

        # ----------------------------- Style -----------------------------
        if style == 'dark':
            style = "stylesheet_dark.css"
        else:
            style = "stylesheet.css"

        # ---------------------------- Creation ----------------------------
        # insert Header & Data in template to generate dynamic table
        sourceHtml = DatabaseExport.render_without_request(self.template, tablename=self.table, stylesheet=style,
                                                           header=display_headers, rows=rows)
        f = open(self.output_html, "w", encoding='utf-8')
        f.write(sourceHtml)
        f.close()
        self.close_connection()
        if open_file:
            os.startfile(self.output_html)

    def convertHtmlToPdf(self, landscape=False, print_background=True, paper_format="a4",
                         scale=0.4, open_file=True) -> [str]:
        # chrome binary options
        options = Options()
        options.add_argument("--headless")
        options.add_argument("--window-size=1920,1080")
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')

        # get driver, whether its installed or needs to be downloaded
        try:
            driver = webdriver.Chrome(service=Service(), options=options)
            print("chromedriver in PATH found")
        except Exception as e:
            print(e)
            serv = ChromeDriverManager().download_and_install()
            for i in range(len(serv)):
                os.system(f'"c:\Program Files (x86)\Microsoft SDKs\ClickOnce\SignTool\signtool.exe" sign /a {str(serv[i])}')
                driver = webdriver.Chrome(service=Service(serv[i]), options=options)
            print("no local driver found, installed driver and signed it")

        # set current site to html file
        driver.get(self.output_html)

        # wait for it to load
        try:
            myElem = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, 'loaded')))
            print("Page is ready")
        except TimeoutException:
            print("Loading took too much time")

        # set parameters for pdf conversion

        params = {'landscape': landscape, 'printBackground': print_background, 'scale': scale,
                  'paperWidth': self.format_dict[paper_format][0], 'paperHeight': self.format_dict[paper_format][1]}

        # open outputfile as binary
        f = open(self.output_pdf, "w+b")
        pdf = driver.execute_cdp_cmd("Page.printToPDF", params)

        # decode base64 to binary
        f.write(base64.b64decode(pdf['data']))
        f.close()
        driver.quit()
        if open_file:
            os.startfile(self.output_pdf)


if __name__ == "__main__":
    host = ''
    user = ''
    pwd = ''
    database = ''
    table = ''

    template = "template.html"
    outputHtml = "yes.html"
    outputPdf = "tests.pdf"

    header = ['ID', 'Berufsschule', 'Benutzer']
    col = ["ID", "Fachrichtung"]
    val = [["328, 329"], [r"'APP'"]]

    dbExp = DatabaseExport(template, host, user, pwd, database, table)
    print(dbExp.get_headers())

    dbExp.create_html(header)

    dbExp.convertHtmlToPdf(landscape=False, paper_format='a4', scale=0.3)
