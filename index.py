from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import load_workbook
from time import sleep
import random

# Function to wait for the visibility of an element
def wait_for_visibility(driver, locator, timeout=10):
    return WebDriverWait(driver, timeout).until(EC.visibility_of_element_located(locator))

# Function to click on an element
def click_element(driver, locator, timeout=10):
    element = wait_for_visibility(driver, locator, timeout)
    element.click()

# Configuring the Chrome service and options
service = Service(ChromeDriverManager().install())
options = webdriver.ChromeOptions()
options.add_argument("--guest")  # Run Chrome in guest mode
options.add_argument("--start-maximized")
options.add_argument("--disable-notifications")
options.add_experimental_option("prefs", {"profile.default_content_setting_values.cookies": 2})
options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.110 Safari/537.36")  # Set user-agent string

# Load the Excel workbook
workbook = load_workbook('/Users/rpecorreia/Desktop/Yoga - faturas/2024/Livro1.xlsx')

try:
    # Select a specific worksheet
    worksheet = workbook['Mar']

    # Create a list to store non-empty values from column A
    column_a_values = []

    # Iterate through cells in column A
    for cell in worksheet['A']:
        if cell.value:
            column_a_values.append(cell.value)

    # Print all non-empty values from column A
    for value in column_a_values:
        print(value)

    # Print the length of column_a_values
    print("\nLength of column A values:", len(column_a_values),"\n")

finally:
    # Close the workbook when you're done
    workbook.close()

    # --------------------RoomRaccoon ---------------------------
    # Creating an instance of the Chrome driver
    driver = webdriver.Chrome(service=service, options=options)

    try:
        # Opening RoomRaccoon
        driver.get("https://rms.roomraccoon.com/admin/reservations/")

        # Waiting for the email and password fields to be visible
        email = wait_for_visibility(driver, (By.NAME, 'user[email]'))
        pw = wait_for_visibility(driver, (By.NAME, 'user[password]'))
        login = wait_for_visibility(driver, (By.ID, "login-btn"))

        # Entering the email and password
        email.send_keys("####")
        pw.send_keys("####")

        # Clicking login
        login.click()


    finally:
        # Quitting the driver
        sleep(1000)
        driver.quit()
