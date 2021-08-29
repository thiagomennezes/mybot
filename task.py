from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.FileSystem import FileSystem
from RPA.PDF import PDF
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

import json
import re
import time

class MyBot:

    def __init__(self):
        self.browser = Selenium()
        self.dirpath = self.__make_dir("./output")
        self.filepaths = {
            "config": "settings.json",
            "excel": f"{self.dirpath}/Agencies.xlsx",
            "pdf": []
        }
        self.url = "http://itdashboard.gov"
        self.__read_config_file(self.filepaths["config"])
        self.browser.set_download_directory(self.dirpath)

    def run(self):
        try:
            self.browser.open_available_browser(self.url)
            self.browser.click_element_when_visible("//div/a[@href='#home-dive-in']")
            self.__create_excel()
            self.__convert_tiles_to_excel()
            self.browser.click_element(f"//span[contains(text(), '{self.target}')]")
            self.__convert_table_to_excel()
            self.__download_files()
            self.__compare_data()
        finally:
            self.browser.close_all_browsers()

    def __compare_data(self):
        message = ""
        data = {"excel": None, "pdf": []}
        for filepath in self.filepaths["pdf"]:
            data["pdf"].append(self.__extract_data_from_pdf(filepath))
        data["excel"] = self.__extract_data_from_excel(self.filepaths["excel"], self.target)
        for p in data["pdf"]:
            found = False
            for e in data["excel"]:
                if e["A"] == p["uii"] and e["C"] == p["title"]:
                    message += f"{p['uii']}: Titles match to {p['title']}\n"
                    found = True
                    break
                if e["A"] == p["uii"] and e["C"] != p["title"]:
                    message += f"{p['uii']}: Titles unmatch --> {p['title']} != {e['C']}\n"
                    found = True
                    break
            if not found:
                message += f"{p['uii']}: Not found\n"
        fs = FileSystem()
        fs.create_file(f"{self.dirpath}/compare-pdf.txt", message, "utf-8", True)

    def __convert_table_to_excel(self):
        file = Files()
        try:
            data = self.__extract_data_from_table()
            file.open_workbook(self.filepaths["excel"])
            file.create_worksheet(self.target, data)
            file.save_workbook()
        finally:
            file.close_workbook()

    def __convert_tiles_to_excel(self):
        file = Files()
        try:
            data = self.__extract_data_from_tiles()            
            file.open_workbook(self.filepaths["excel"])
            file.rename_worksheet("Sheet", "Agencies")
            file.append_rows_to_worksheet(data)
            file.save_workbook()
        finally:
            file.close_workbook()

    def __create_excel(self):
        file = Files()
        file.create_workbook(self.filepaths["excel"], "xlsx")
        file.save_workbook()
        file.close_workbook()

    def __download_files(self):
        ancors = self.browser.find_elements(
            "css:#investments-table-object tbody > tr td:nth-of-type(1) a"
        )
        links = []
        for a in ancors:
            links.append(a.get_attribute("href"))
            filename = a.text
            self.filepaths["pdf"].append(f"{self.dirpath}/{filename}.pdf")
        indexes = range(0, len(links))
        for i in indexes:
            self.browser.open_available_browser(links[i]) 
            self.browser.click_element_when_visible("css:#business-case-pdf > a")
            self.__wait_download(self.filepaths["pdf"][i])

    def __extract_data_from_excel(self, filepath, worksheet):
        file = Files()
        file.open_workbook(filepath)
        try:
            return file.read_worksheet(worksheet)
        finally:
            file.close_workbook()

    def __extract_data_from_pdf(self, filepath):
        pdf = PDF()
        text = pdf.get_text_from_pdf(filepath)
        data = {
            "title": re.findall(r"Name of this Investment: (.+)2\.", text[1])[0],
            "uii": re.findall(r"Unique Investment Identifier \(UII\): (.+)Section B", text[1])[0]
        }
        return data

    def __extract_data_from_table(self):
        data = []
        table = self.__prepare_table()
        rows = self.browser.find_elements("css:tbody > tr", table)
        for row in rows:
            cols = self.browser.find_elements("css:td", row)
            datum = [col.text for col in cols]
            data.append(datum)
        return data
        
    def __extract_data_from_tiles(self):
        tiles = self.__wait(By.ID, "agency-tiles-container")
        agencies = self.browser.find_elements("css:div a span:nth-of-type(1)", tiles)
        amounts = self.browser.find_elements("css:div a span:nth-of-type(2)", tiles)
        indexes = range(0, len(agencies))
        data = []
        for i in indexes:
            data.append([agencies[i].text, amounts[i].text])
        return data

    def __make_dir(self, dirpath):
        fs = FileSystem()
        fs.create_directory(dirpath)
        return fs.absolute_path(dirpath)

    def __prepare_table(self):
        table = self.__wait(By.ID, "investments-table-object")
        self.browser.find_element(
            "css:#investments-table-object_length select > option:nth-child(4)"
        ).click()
        size = 0
        while size < 11:
            table = self.__wait(By.ID, "investments-table-object")
            rows = self.browser.find_elements("css:tbody > tr", table)
            size = len(rows)
        return table

    def __read_config_file(self, filepath):
        fs = FileSystem()
        data = json.loads(fs.read_file(filepath))
        self.target = data["target"]

    def __wait(self, type, locator):
        waiter = WebDriverWait(self.browser.driver, 60)
        return waiter.until(EC.visibility_of_element_located((type, locator)))

    def __wait_download(self, filepath):
        fs = FileSystem()
        countdown = 20
        while (
            (
                fs.does_file_not_exist(filepath)
                or (fs.does_file_exist(filepath) and fs.get_file_size(filepath) == 0)
            )
            and countdown > 0
        ):
            time.sleep(3)
            countdown -= 1

if __name__ == "__main__":
    bot = MyBot()
    bot.run()
