from http.client import OK

from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

driver = webdriver.Firefox(
    executable_path='geckodriver.exe')

home_page = ""

driver.get(home_page)

input("Press Enter to continue...")

with open("lista.txt", "r") as number_list:
    count = 0
    for number in number_list:
        count += 1
        print(f'{count}: {number.strip()}')
        try:
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located(
                    (By.ID, "folderbrowser_PDM.searchInListTextBox"))
            )
        finally:
            search_bar = driver.find_element(
                By.ID, "folderbrowser_PDM.searchInListTextBox")
            search_bar.clear()
            search_bar.send_keys(number.strip())
            search_bar.send_keys(Keys.RETURN)
            search_bar.send_keys(Keys.RETURN)

        try:
            WebDriverWait(driver, 20).until(
                EC.text_to_be_present_in_element(
                    (By.LINK_TEXT, number.strip()), number.strip())
            )
        except TimeoutException:
            search_bar = driver.find_element(
                By.ID, "folderbrowser_PDM.searchInListTextBox")
            search_bar.clear()
            search_bar.send_keys(number.strip())
            search_bar.send_keys(Keys.RETURN)
            search_bar.send_keys(Keys.RETURN)
            driver.implicitly_wait(5)
        finally:
            document_link = driver.find_element(By.LINK_TEXT, number.strip())
            action = ActionChains(driver)
            action.move_to_element(document_link).perform()
            try:
                action.context_click(document_link).perform()
            except StaleElementReferenceException:
                driver.implicitly_wait(3)
                document_link = driver.find_element(By.LINK_TEXT, number.strip())
                action = ActionChains(driver)
                action.move_to_element(document_link).perform()
                action.context_click(document_link).perform()

        try:
            WebDriverWait(driver, 10).until(
                EC.text_to_be_present_in_element(
                    (By.LINK_TEXT, "Rename"), "Rename")
            )
        finally:
            rename_link = driver.find_element(By.LINK_TEXT, "Rename")
            action = ActionChains(driver)
            action.move_to_element(rename_link).perform()
            action.click(rename_link).perform()

        driver.switch_to.window(driver.window_handles[1])

        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "number"))
            )
        finally:
            number_input = driver.find_element(By.ID, "number")
            number_input.clear()
            number_input.send_keys(number.strip())
            buttons = driver.find_elements(By.TAG_NAME, "button")
            action = ActionChains(driver)
            action.move_to_element(buttons[2]).perform()
            action.click(buttons[2]).perform()
            #  2 - OK, 3 - Cancel

        driver.switch_to.window(driver.window_handles[0])
