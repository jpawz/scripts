#! python3
# Batch adds alternative to WTparts in Windchill
# Input file (*.xlsx) with filled two columns:
#      part_number | alternative_part_number
# as processing goes by, third column will be filled with statuses:
# ok - alternative added, fail - alternative not added, empty - not yet processed.

import openpyxl
import sys
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

if len(sys.argv) != 2:
    print("use: windchill_batch_rename.py excel_file.xlsx")
    exit()

xlsx_file = sys.argv[1]
workbook = load_workbook(filename=xlsx_file)
sheet = workbook.active
windchill_url = "http://windchill_url"
browser = webdriver.Ie()
browser.set_page_load_timeout(5)
browser.get(windchill_url)
main_window = browser.current_window_handle
input("login and press [Enter] to continue...")


def find_part(part_number):
    find_field = WebDriverWait(browser, 5).until(
        EC.presence_of_element_located(By.ID), "gloabalSearchField")
    find_field.clear()
    find_field.send_keys(part_number)
    find_field.send_keys(Keys.ENTER)


def open_part_info_window(part_number):
    number_link = WebDriverWait(browser, 5).until(
        EC.presence_of_element_located((By.LINK_TEXT, part_number)))
    ac = ActionChains(browser)
    ac.click(number_link).perform()
    switch_to_related_objects_tab()
    open_find_alternate_part_window()


def open_find_alternate_part_window():
    alt_bar = browser.find_element_by_id("netmarkets.alternates.list.toolBar")
    alt_bar.location_once_scrolled_into_view
    buttons = alt_bar.find_elements_by_tag_name("button")
    add_button = buttons[0]  # "Add" button
    add_button.click()


def switch_to_related_objects_tab():
    related_object_tab = WebDriverWait(browser, 5).until(
        EC.presence_of_element_located((By.LINK_TEXT, "Related Objects")))
    related_object_tab.click()
    browser.implicitly_wait(3)


def handle_find_alternate_part_window(alternative_number):
    WebDriverWait(browser, 5).until(EC.new_window_is_opened)
    browser.switch_to.window(browser.window_handles[1])  # new window
    try:
        wait = WebDriverWait(browser, 5)
        find_number_field = wait.until(
            EC.presence_of_element_located((By.ID, "number2_SearchTextBox")))
        find_number_field.clear()
        find_number_field.send_keys(alternative_number)
        find_number_field.send_keys(Keys.ENTER)
        search_button = browser.find_element_by_name("pickerSearch")
        search_button.click()
        checker = wait.until(
            EC.presence_of_element_located((By.CLASS_NAME,
                                            "x-grid3-col-checker")))
        checker.click()
        ok_btn = browser.find_element_by_xpath(
            "//button[contains(@accesskey, 'o')]")  # c - cancel, o - OK
        ok_btn.click()
    except:
        raise
    finally:
        browser.close()
        browser.switch_to.window(main_window)


def add_alternative_part(part_number, alt_number):
    browser.get(windchill_url)
    find_part(part_number)
    open_part_info_window(part_number)
    handle_find_alternate_part_window(alt_number)


for row in sheet.iter_rows(max_col=3):
    try:
        add_alternative_part(row[0], row[1])
        row[2] = "OK"  # if success
    except Exception as ex:
        print("error at " + row[0])
        print(ex)
        row[2] = "failed"
        continue
    workbook.save(xlsx_file)
