import openpyxl
from selenium.webdriver.support.wait import WebDriverWait
from record import Record
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
import time


def get_resource(path):
    resources = {}
    with open(path, encoding="utf-8") as file:
        for line in file.read().splitlines():
            list = line.split("=")
            resources[list[0]] = list[1]
    return resources


def read_excel_record(excel_path):
    sheet = openpyxl.load_workbook(excel_path)["Sheet1"]

    records = []
    row_size = 1
    while True:
        cell_value = sheet.cell(column=1, row=row_size).value
        if cell_value is None:
            break
        records.append(Record(cell_value))
        row_size += 1
    return records

resource_path = "resource.properties"
resources = get_resource(resource_path)
records = read_excel_record(resources["excel_path"])

form_url = resources["form_url"] + "?usp=pp_url"
form_url += "&entry.1092907065=" + resources["mail_address"]
form_url += "&entry.1319393789=" + resources["name"]
form_url += "&entry.638281816=" + resources["group"]
for r in records:
    try:
        driver = webdriver.Chrome()
        driver.get(form_url)
        time.sleep(5)
        driver.find_element_by_xpath(
            '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[1]/div/div/div[2]/div[1]/div/span/div/div[1]/label'
        ).click()
        driver.find_element_by_xpath(
            '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div[1]/div/div[1]/input'
        ).send_keys(r.url)
        driver.find_element_by_xpath(
            '//*[@id="mG61Hd"]/div[2]/div/div[3]/div[1]/div/div'
        ).click()
        driver.quit()
    except Exception as e:
        print(e)
        exit()
