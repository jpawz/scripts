#! python3
# Batch renames WTparts and CADdocuments in Windchill
# Input file (*.xlsx) should contain three columns:
#      part_number | long_name | short_name
# as processing goes by fourth column will be filled with statuses:
# ok - rename successful, fail - rename failed, empty - not yet processed.

import openpyxl
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

if len(sys.argv) != 2:
    print("use: python windchill_batch_rename.py excel_file.xlsx")
    exit()

xlsx_file = sys.argv[1]
workbook = load_workbook(filename= xlsx_file)
sheet = workbook.active
windchill_site = "http://www."
driver = webdriver.Firefox()
login = input("login: ")
password = input("password: ")
driver.get(windchill_site)
alert = driver.switch_to.alert
alert.send_keys(login + Keys.TAB + password)
alert.accept()

def rename_part(part_number, new_name):
    print("not yet implemented")

for row in sheet.iter_rows(max_col=4):
    rename_part(row[1], row[3])
    row[4] = "OK"

workbook.save(xlsx_file)
