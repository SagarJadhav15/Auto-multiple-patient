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
lines = xlsx_reader('Book2.xlsx')

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







 #create Patient

driver.find_element(By.XPATH, "//span[normalize-space()='Appointments']").click()
driver.find_element(By.XPATH, "//span[normalize-space()='My Appointments']").click()
for line in lines:
    driver.implicitly_wait(10)
    driver.find_element(By.XPATH, "//span[normalize-space()='Create Appointment']").click()
    driver.find_element(By.XPATH, "//span[normalize-space()='Add New Patient']").click()

    firstname = driver.find_element(By.XPATH, "//input[@placeholder='First Name']")
    firstname.send_keys(line[0])  # Assuming first name is in the first column of the Excel file
    lastname = driver.find_element(By.XPATH, "//input[@placeholder='Last Name']")
    lastname.send_keys(line[1])  # Assuming last name is in the second column of the Excel file


    #driver.find_element(By.XPATH, "//input[@placeholder='First Name']").send_keys("Test")
    #driver.find_element(By.XPATH, "//input[@placeholder='Last Name']").send_keys("Auto")
    driver.find_element(By.XPATH, "//input[@placeholder='Select Date']").send_keys("28/12/2000")
    driver.find_element(By.XPATH, "//p[normalize-space()='Male']").click()
    driver.find_element(By.XPATH, "//input[@placeholder='Plot No. Apartment']").send_keys("mumbai")
    driver.find_element(By.XPATH, "//input[@placeholder='City']").send_keys("Thane")
                               # ********** Change contact Number Of Patent ********************
    driver.find_element(By.XPATH, "//input[@name='contactNumber']").send_keys("")
    driver.find_element(By.XPATH, "//input[@name='alternetNumber']").send_keys("9999999999")
    driver.find_element(By.XPATH, "//input[@placeholder='Enter Email']").send_keys("Sagarjadhav1522@gmail.com")
    driver.find_element(By.XPATH, "//input[@placeholder='00-0000-0000-0000']").send_keys("98765654312342")
    driver.find_element(By.XPATH, "//input[@name='panNumber']").send_keys("DFGVF3456D")
    driver.find_element(By.XPATH, "//input[@name='aadharNumber']").send_keys("676545654321")
    driver.find_element(By.XPATH, "//input[@placeholder='Enter Hospital Id']").send_keys("ASD1234")
    driver.find_element(By.XPATH, "//input[@name='sendInvite']").click()
    driver.find_element(By.XPATH, "(//span[normalize-space()='Save'])[1]").click()

        #   Create Appointment
    driver.find_element(By.XPATH, "//input[@id='appointmentTimeSlot']").click()
    driver.find_element(By.XPATH, "//li[@id='appointmentTimeSlot-option-0']").click()
    driver.find_element(By.XPATH, "//p[normalize-space()='HPV - Cervavac']").click()
    driver.find_element(By.XPATH, "//input[@value='1st']").click()
    driver.find_element(By.XPATH, "//span[normalize-space()='Submit']").click()
        #EOC
    driver.find_element(By.XPATH, "//div[@id='PersonalHistory']//div[@id='panel1bh-header']//span[@class='MuiIconButton-label']//*[name()='svg']").click()
   # driver.find_element(By.XPATH, "//input[@id='select-1']").click()
    driver.find_element(By.XPATH, "//input[@id='select-1']").click()
    action = webdriver.ActionChains(driver)
    action.send_keys(Keys.ARROW_UP).perform()
    action.send_keys(Keys.ENTER).perform()

    #driver.find_element(By.XPATH, "(//input[@id='select-1'])[1]").click()
    driver.implicitly_wait(10)
    driver.find_element(By.XPATH, "(//span[@class='MuiButton-label'][normalize-space()='Submit'])[2]").click()




