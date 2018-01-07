#! python3
# Scraper: get short names from Products Catalogue website
# Input: xlsx file with first column containing part numbers
#        and optionally long names in second column
# Output: the same xlsx file with third column filled with
#         short names

import sys, openpyxl
from bs4 import BeautifulSoup
from openpyxl import load_workbook

if len(sys.argv) != 2:
    print("use: python catalog_names.py excel_file.xlsx")
    exit()

xlsx_file = sys.argv[1]
workbook = load_workbook(filename= xlsx_file)
sheet = workbook.active
name_library = {}

for row in sheet.iter_rows(max_col=3):
    part_number_dashed = dash(row[0])
    part_short_name = get_short_name(part_number_dashed)
    row[3] = part_short_name

def dash(number):
    if len(number) == 9:
        return number[:3] + "-" + number[3:5] + "-"
    if len(number) == 11:
        return dash(number[:9]) + "-" + number[9:]

def get_short_name(part_number):
    ''' Return short name for given number.
        It first checks if number is in name_library dictionary. If not
        tryes to find in website, also filling name_library with
        numbers and corresponding short names from the page.'''
    if name_library.get(part_number)
        return name_library.get(part_number)
    soup = BeautifulSoup(site_url)
    return short_name

workbook.save(xlsx_file)
