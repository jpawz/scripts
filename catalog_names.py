#! python3
# Scraper: get short names from Products Catalogue website
# Input: xlsx file with first column containing part numbers (continous)
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

# TODO: consider other number lengths
def dash(number):
    if len(number) == 9:
        return number[:3] + "-" + number[3:5] + "-" + number[5:]
    if len(number) == 11:
        return dash(number[:9]) + " " + number[9:]

def get_short_name(part_number):
    ''' Return short name for given number.
        It first checks if number is in name_library dictionary. If not
        tryes to find in website, also filling name_library with
        numbers and corresponding short names from the page.'''
    if name_library.get(part_number):
        return name_library.get(part_number)
    with open("file.htm",encoding="windows-1250") as fp:
        soup = BeautifulSoup(fp, "html.parser")
        naz = soup.find_all(name="td",attrs={"class": "naz"}, )
        nr = soup.find_all(name="td",attrs={"class": "nr"}, )
        for i in range(len(nr)):
            name_library[nr[i].string] = naz[i].string
    return name_library.get(part_number)

for row in sheet.iter_rows(max_col=3):
    part_number_dashed = dash(row[0].value)
    part_short_name = get_short_name(part_number_dashed)
    row[2].value = part_short_name
    workbook.save(xlsx_file)
