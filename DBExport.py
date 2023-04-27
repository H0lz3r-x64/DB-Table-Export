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
    def __init__(self, template: str, exportName, output_html="", output_pdf=""):
        # Filenames
        self.template = template
        self.exportName = exportName
        if output_html == "":
            self.output_html = os.path.abspath(f"{exportName}_Export_{datetime.date.today()}.html")
        else:
            self.output_html = os.path.abspath(output_html)
        if output_pdf == "":
            self.output_pdf = os.path.abspath(f"{exportName}_export_{datetime.date.today()}.pdf")
        else:
            self.output_pdf = os.path.abspath(output_pdf)

        # Formats
        self.format_dict = {"Legal": (8.5, 14), "legal": (8.5, 14),
                            "Letter": (8.5, 11), "letter": (8.5, 11),
                            "A5": (5.8, 8.3), "a5": (5.8, 8.3),
                            "A4": (8.3, 11.7), "a4": (8.3, 11.7),
                            "A3": (11.7, 16.5), "a3": (11.7, 16.5)}

    @staticmethod
    def render_without_request(template_name, **template_vars):
        # Usage is the same as flask.render_template:
        env = jinja2.Environment(
            loader=jinja2.FileSystemLoader("./")
        )
        template = env.get_template(template_name)
        return template.render(**template_vars)

    def get_format_in_inches(self, paper_format: str):
        return self.format_dict[paper_format]

    def create_html(self, display_headers, rows, open_file=True):
        sourceHtml = DatabaseExport.render_without_request(self.template, tablename=self.exportName, header=display_headers, rows=rows)
        
        f = open(self.output_html, "w", encoding='utf-8')
        f.write(sourceHtml)
        f.close()
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
