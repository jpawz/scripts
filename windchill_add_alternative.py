#! python3
# Batch adds alternative to WTparts in Windchill
# Input file (*.xlsx) with filled two columns:
#	  part_number | alternative_part_number
# as processing goes by, third column will be filled with statuses:
# ok - alternative added, fail - alternative not added, empty - not yet processed.

import sys

import openpyxl
from openpyxl import load_workbook
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

if len(sys.argv) != 2:
    print("use: skrypt.py excel_file.xlsx")
    exit()

xlsx_file = sys.argv[1]
workbook = load_workbook(filename=xlsx_file)
sheet = workbook.active
windchill_url = ""

cap = DesiredCapabilities().INTERNETEXPLORER
cap['ignoreProtectedModeSettings'] = True
cap['IntroduceInstabilityByIgnoringProtectedModeSettings'] = True
cap['nativeEvents'] = True
cap['ignoreZoomSetting'] = True
cap['requireWindowFocus'] = True
cap['INTRODUCE_FLAKINESS_BY_IGNORING_SECURITY_DOMAINS'] = True

browser = webdriver.Ie(
    capabilities=cap,
    executable_path=r'D:\WinPython\python\IEDriverServer.exe')
browser.set_page_load_timeout(5)
browser.get(windchill_url)
main_window = browser.current_window_handle
wait_5_sec = WebDriverWait(browser, 5)
input("login and press [Enter] to continue...")
MSG_MISSING_NR = "NUMBER NOT FOUND"
MSG_MISSING_ALT_NR = "ALT NUMBER NOT FOUND"


def find_part(part_number):
    browser.get(windchill_url)
    find_field = wait_5_sec.until(
        EC.presence_of_element_located((By.ID, "gloabalSearchField")))
    find_field.clear()
    find_field.send_keys(part_number)
    find_field.send_keys(Keys.ENTER)


def open_part_info_window(part_number):
    try:
        number_links = wait_5_sec.until(
            EC.presence_of_element_located((By.PARTIAL_LINK_TEXT,
                                            part_number)))
    except TimeoutException:
        raise Exception(MSG_MISSING_NR)
    else:
        if (len(number_links) > 1):
            raise Exception("Error: not unique number link found")
        ac = ActionChains(browser)
        ac.click(number_links[0]).perform()


def switch_to_tab(link_text):
    tab = wait_5_sec.until(
        EC.presence_of_element_located((By.LINK_TEXT, link_text)))
    tab.click()


def alternative_already_exist(alternative_number):
    alternates_table = browser.find_element_by_id("ext-gen2255")
    try:
        alternates_table.find_element_by_link_text(alternative_number)
        return True
    except NoSuchElementException:
        return False


def open_find_alternate_part_window():
    alt_bar = wait_5_sec.until(
        EC.presence_of_element_located((By.ID,
                                        "netmarkets.alternates.list.toolBar")))
    alt_bar.location_once_scrolled_into_view
    buttons = alt_bar.find_elements_by_tag_name("button")
    add_button = buttons[0]  # "Add" button
    add_button.click()


def handle_find_alternative_part_window(alternative_number):
    wait_5_sec.until(EC.number_of_windows_to_be(2))
    if (main_window == browser.window_handles[0]):
        new_window = browser.window_handles[1]
    else:
        new_window = browser.window_handles[0]
    browser.switch_to.window(new_window)
    try:
        find_number_field = wait_5_sec.until(
            EC.presence_of_element_located((By.ID, "number2_SearchTextBox")))
        find_number_field.clear()
        find_number_field.send_keys(alternative_number)
        find_number_field.send_keys(Keys.ENTER)
        search_button = browser.find_element_by_name("pickerSearch")
        search_button.click()
        checker = wait_5_sec.until(
            EC.presence_of_element_located((By.CLASS_NAME,
                                            "x-grid3-col-checker")))
        tab_rows = browser.find_elements_by_class_name("x-grid3-row-table")
        if (len(tab_rows) > 1):
            raise Exception("multiple elements found")
    except TimeoutException:
        browser.switch_to_window(new_window)
        browser.close()
        raise Exception(MSG_MISSING_ALT_NR)
    else:
        checker.click()
        ok_btn = browser.find_element_by_id("ext-gen40")
        ok_btn.click()
        try:
            WebDriverWait(browser, 1).until(EC.alert_is_present())
            browser.switch_to.alert.accept()
            browser.switch_to_window(new_window)
            browser.close()
        except TimeoutException:
            pass
    finally:
        browser.switch_to.window(main_window)


def add_alternative_part(part_number, alt_number):
    browser.switch_to.window(main_window)
    find_part(part_number)
    open_part_info_window(part_number)
    switch_to_tab("Related Objects")
    if (not alternative_already_exist(alt_number)):
        open_find_alternate_part_window()
        handle_find_alternative_part_window(alt_number)


def part_is_already_checked():
    if ((str(row[3].value) == MSG_MISSING_NR)
            or (str(row[3].value) == MSG_MISSING_ALT_NR)
            or (str(row[3].value) == "OK")):
        return True
    return False


for index, row in enumerate(sheet.rows):
    try:
        part_n = str(row[0].value).strip()
        alt_n = str(row[1].value).strip()
        if ((not part_n) or (not alt_n) or (part_n == "None")
                or (alt_n == "None")):
            continue
        if (part_is_already_checked()):
            continue
        add_alternative_part(part_n, alt_n)
        sheet.cell(row=index + 1, column=4).value = "OK"
    except Exception as ex:
        sheet.cell(row=index + 1, column=4).value = str(ex)
        continue
    finally:
        print(index)
        workbook.save(xlsx_file)
workbook.save(xlsx_file)
browser.quit()
