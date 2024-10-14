import time
from time import sleep
import openpyxl
from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.wait import WebDriverWait

# Set up options for Chrome WebDriver
options = Options()
options.add_experimental_option("detach", True)  # Keeps browser open after script ends

# Set up the Selenium WebDriver
driver = webdriver.Chrome(options=options)

def xlsx_reader(xlsx_file):
    # Load the workbook and select the active sheet
    workbook = openpyxl.load_workbook(xlsx_file)
    sheet = workbook.active

    # Read all rows in the sheet
    data = []
    for row in sheet.iter_rows(min_row=2, values_only=True):  # Start from row 2 (skip header)
        data.append(row)

    return data

# Load the data from the Excel file
lines = xlsx_reader('Book1.xlsx')

# Iterate over each line (each row from Excel)

driver.get("https://demo.nuqare.com/")  # Open the website
driver.maximize_window()
driver.implicitly_wait(10)

    # Log in
wait = WebDriverWait(driver, 20)
driver.find_element(By.CSS_SELECTOR, "#standard-phoneno-input").send_keys("8999379626")
driver.find_element(By.CSS_SELECTOR, "button[type='submit'] span[class='MuiButton-label']").click()

    # OTP
driver.find_element(By.CSS_SELECTOR, "input[placeholder='-'][aria-label='Please enter verification code. Digit 1']").send_keys("5746")
driver.find_element(By.XPATH, "//span[normalize-space()='Login OTP']").click()

    # Select location
driver.find_element(By.XPATH, "//p[normalize-space()='Wag Cancer Clinic']").click()
driver.implicitly_wait(10)
driver.find_element(By.XPATH, "//span[normalize-space()='OK']").click()
for line in lines:
    # Navigate to the form
    driver.find_element(By.XPATH,"//button[@title='Instant EOC Form']//span[@class='MuiIconButton-label']//*[name()='svg']").click()

    # Fill the form with first name and last name from the Excel file
    firstname = driver.find_element(By.XPATH, "//input[@placeholder='Search/Enter Patient Name']")
    firstname.send_keys(line[0])  # Assuming first name is in the first column of the Excel file
    lastname = driver.find_element(By.XPATH, "//input[@placeholder='Last Name']")
    lastname.send_keys(line[1])  # Assuming last name is in the second column of the Excel file

    driver.find_element(By.XPATH, "//input[@name='dateOfBirth']").send_keys("28/12/2000")
    driver.find_element(By.XPATH, "//input[@value='Male']").click()



    driver.find_element(By.XPATH, "//input[@placeholder='Select Dose']").send_keys("1")
    action = webdriver.ActionChains(driver)
    action.send_keys(Keys.ARROW_DOWN).perform()
    action.send_keys(Keys.ENTER).perform()
    driver.find_element(By.XPATH, "//span[normalize-space()='Submit']").click()
    driver.refresh()

    # Optional: Wait a bit before processing the next entry
    sleep(2)

# You can close the driver after the loop finishes if desired
#driver.quit()
