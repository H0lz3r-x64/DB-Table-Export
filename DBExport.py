import base64, os, jinja2, datetime, bs4
from typing import Union, Collection

from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.by import By
from webdrivermanager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options


class DatabaseExport:
    def __init__(self, template: str, export_name: str, path_to_output_html="", path_to_output_pdf=""):
        # Filenames
        self.template = template
        splitup = export_name.split("<split>")
        self.title = splitup[0]
        self.extratitle = ""
        if len(splitup) > 1:
            self.extratitle = splitup[1]
        self.escaped_export_name = export_name.replace("<split>", "_")

        if path_to_output_html == "":
            self.output_html = os.path.abspath(f"{export_name}_Export_{datetime.date.today()}.html")
        else:
            self.output_html = os.path.abspath(os.path.join(path_to_output_html, f"{self.escaped_export_name}.html"))
        if path_to_output_pdf == "":
            self.output_pdf = os.path.abspath(f"{export_name}_export_{datetime.date.today()}.pdf")
        else:
            self.output_pdf = os.path.abspath(os.path.join(path_to_output_pdf, f"{self.escaped_export_name}.pdf"))

        # Formats
        self.format_dict = {"Legal": (8.5, 14), "legal": (8.5, 14),
                            "Letter": (8.5, 11), "letter": (8.5, 11),
                            "A5": (5.8, 8.3), "a5": (5.8, 8.3),
                            "A4": (8.3, 11.7), "a4": (8.3, 11.7),
                            "A3": (11.7, 16.5), "a3": (11.7, 16.5)}

        self.tmp_html_path = "tmp_files\\tmp_report.html"
        self.tmp_pdf_path = "tmp_files\\tmp_report.pdf"

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

    @staticmethod
    def __consolidate_css_html(input_html) -> str:
        """
        Consolidate the html string with its external css references and png images,
        so any externally referenced css page(s) and image file(s) are not needed.

        :param input_html: A string containing the content of the html file
        :return: A string containing the html source now with the before external css and images embedded into it
        """
        soup = bs4.BeautifulSoup(input_html, features="lxml")
        stylesheets = soup.findAll("link", {"rel": "stylesheet"})
        for s in stylesheets:
            t = soup.new_tag('style')
            c = bs4.element.NavigableString(open(s["href"]).read())
            t.insert(0, c)
            t['type'] = 'text/css'
            s.replaceWith(t)

        images = soup.findAll("img")
        for i in images:
            f = open(i["src"], "rb")
            b = base64.b64encode(f.read())
            f.close()
            i["src"] = "data:image/png;base64," + b.decode()

        return str(soup)

    def create_html(self, display_headers: Collection, rows: Collection, rows_addition_data: Collection, open_file=True,
                    save_file=False) -> str:
        print(f"{datetime.datetime.now()}: creating {self.escaped_export_name} HTML file...")

        # checking shape
        print("checking shape")
        if not self.__check_same_shape__(rows, rows_addition_data, 1):
            raise TypeError("The collections have different shapes")

        # generate source html string from template and data
        sourceHtml = DatabaseExport.render_without_request(
            self.template, title=self.title, extratitle=self.extratitle, header=display_headers,
            rows=rows, rows_addition_data=rows_addition_data, zip=zip
        )

        # consolidate the html string with its external css references,
        # so any externally referenced css page(s) are not needed.
        sourceHtml = self.__consolidate_css_html(sourceHtml)

        # save as temporary file
        print(f"{datetime.datetime.now()}: saving temporary PDF file")
        self.__save_to_file(self.tmp_html_path, sourceHtml, override_check=False)

        # open file
        if open_file:
            print(f"{datetime.datetime.now()}: opening HTML file")
            os.startfile(self.tmp_html_path)

        # save to file
        if save_file:
            print(f"{datetime.datetime.now()}: saving HTML file")
            self.output_html = self.__save_to_file(self.output_html, sourceHtml, override_check=True)
            print(f"{datetime.datetime.now()}: saved HTML file to {self.output_html}")

        # return path
        return self.output_html

    def convert_html_to_pdf(self, is_landscape=False, print_background=True, paper_format="a4",
                            scale=0.4, open_file=True, save_file=False) -> Union[str, None]:
        print(f"{datetime.datetime.now()}: converting HTML to PDF...")

        # chrome binary options
        options = Options()
        options.add_argument("--headless")
        options.add_argument("--window-size=1920,1080")
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')

        driver = WebDriver
        # get driver, search for existing one or download it
        try:
            driver = webdriver.Chrome(service=Service(), options=options)
            print("    chromedriver in PATH found")
        except Exception as e:
            print(e)
            serv = ChromeDriverManager().download_and_install()
            for i in range(len(serv)):
                driver = webdriver.Chrome(service=Service(serv[i]), options=options)
            print("    no local driver found, installed driver")

        # set current site to html file
        driver.get(os.path.abspath(self.tmp_html_path))

        # wait for it to load
        try:
            myElem = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, 'loaded')))
            print("    HTML page successfully loaded")
        except TimeoutException:
            print("    Loading took too much time")
            print("    Aborting...")
            return None

        # set parameters for pdf conversion

        params = {'landscape': is_landscape, 'printBackground': print_background, 'scale': scale,
                  'paperWidth': self.format_dict[paper_format][0], 'paperHeight': self.format_dict[paper_format][1]}

        # perform pdf conversion
        pdf = driver.execute_cdp_cmd("Page.printToPDF", params)
        driver.quit()

        # save as temporary file
        print(f"{datetime.datetime.now()}: saving temporary PDF file")
        self.__save_to_file(self.tmp_pdf_path, base64.b64decode(pdf['data']), override_check=False)

        # open file
        if open_file:
            print(f"{datetime.datetime.now()}: opening PDF file")
            os.startfile(self.tmp_pdf_path)

        # save file
        if save_file:
            print(f"{datetime.datetime.now()}: saving PDF file")
            self.output_pdf = self.__save_to_file(self.output_pdf, base64.b64decode(pdf['data']),
                                                  override_check=True)
            print(f"{datetime.datetime.now()}: saved PDF file to {self.output_pdf}")

        # return path
        return self.output_pdf

    @staticmethod
    def __save_to_file(output_path: str, data: Union[str, bytes], override_check=False) -> str:
        if override_check and os.path.exists(output_path):
            # Split the output path into the part before and after the dot
            no_suffix_path, suffix = output_path.split('.')

            # Initialize a counter for duplicate file names
            i = 1

            # Loop until there is no existing file with the same name
            while os.path.exists(f"{no_suffix_path} ({i}).{suffix}"):
                i += 1

            # If the counter is greater than zero, append it to the output path
            if i > 0:
                output_path = f"{no_suffix_path} ({i}).{suffix}"

        # Set the mode and encoding for writing to the file
        mode = 'w'
        encoding = 'utf-8'
        # If the data is bytes, use binary mode and no encoding
        if type(data) == bytes:
            mode = 'w+b'
            encoding = None

        # Open the file with the given mode and encoding
        with open(output_path, mode, encoding=encoding) as f:
            # Write the data to the file
            f.write(data)

        # Return the output path as a string
        return output_path

    @staticmethod
    def __check_same_shape__(collection1: Collection, collection2: Collection, depth: int = None) -> bool:
        """Check if two collections have the same shape up to a certain depth.

        The shape of a collection is defined by its length and
        the length of its nested collections (if any).

        Args:
            collection1: A list or tuple.
            collection2: A list or tuple.
            depth: An optional integer indicating the maximum depth to check.
                   If None, the function checks the shape of the entire collections.

        Returns:
            True if the collections have the same shape up to the given depth,
            False otherwise.

        Raises:
            ValueError: If either collection is not a list or tuple,
                        or if depth is not a positive integer or None.

        Examples:
            >>> __check_same_shape__([1, 2], [3, 4])
            True
            >>> __check_same_shape__([1, [2]], [3, [4]])
            True
            >>> __check_same_shape__([1, [2]], [3])
            False
            >>> __check_same_shape__([1], 2)
            ValueError: Invalid collections
            >>> __check_same_shape__([1, [2]], [3, [4]], depth=2)
            True
            >>> __check_same_shape__([1, [2]], [3, {4}], depth=2)
            False
            >>> __check_same_shape__([1, [2]], [3, [4]], depth=3)
            True
            >>> __check_same_shape__([1, [2]], [3], depth=1)
            False
            >>> __check_same_shape__([1, 2, [3, 4]], [5, 6, {7, 8}], depth=2)
            False
        """
        # check if the collections are valid
        if not isinstance(collection1, (list, tuple)) or not isinstance(collection2, (list, tuple)):
            raise ValueError("Invalid collections")

        # check if the depth is valid
        if depth is not None and (not isinstance(depth, int) or depth < 0):
            raise ValueError("Invalid depth")

        # check if the collections are empty
        if not collection1 or not collection2:
            return False

        # check if the collections have the same length
        if len(collection1) != len(collection2):
            return False

        # check if the depth is zero
        if depth == 0:
            return True

        # check if the collections have the same shape recursively up to the given depth
        return all(__check_same_shape__(item1, item2, depth - 1 if depth is not None else None) if isinstance(item1, (
        list, tuple)) and isinstance(item2, (list, tuple))
                   else not isinstance(item1, (list, tuple)) ^ isinstance(item2, (list, tuple))
                   for item1, item2 in zip(collection1, collection2))
        # Note: The XOR operator (^) is used to check if the items have different types.
        # This avoids raising an exception when one item is a nested collection and the other is not,
        # and returns False as expected.
