import time
from PIL import Image
import pytesseract
import pandas as pd
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys

# Tesseract path
tesseract_path = r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe"

# Configure Selenium WebDriver
driver_path = "E:\\chromedriver_win32\\chromedriver.exe"
service = Service(driver_path)
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=service, options=options)

# Configure OCR
pytesseract.pytesseract.tesseract_cmd = tesseract_path

# Function to capture screenshot of the entire browser page
def capture_full_browser_screenshot(driver, path):
    # Get the full height of the page
    full_height = driver.execute_script("return Math.max( document.body.scrollHeight, document.body.offsetHeight, document.documentElement.clientHeight, document.documentElement.scrollHeight, document.documentElement.offsetHeight );")
    
    # Set the window size to the full height
    driver.set_window_size(driver.get_window_size()["width"], full_height)
    
    # Capture the screenshot
    driver.save_screenshot(path)
    
    # Reset the window size to its original value
    driver.set_window_size(driver.get_window_size()["width"], driver.get_window_size()["height"])
    
    # Open the screenshot image
    image = Image.open(path)
    return image

# Navigate to the webpage
driver.get('https://ntaresults.nic.in/resultservices/JEEMAINauth23s2p1')

# Wait for the captcha image to be present
element_locator = (By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_captchaimg"]')
wait = WebDriverWait(driver, 30)
element = wait.until(EC.presence_of_element_located(element_locator))

# Function to extract captcha text from image
def get_captcha(driver, element, path):
    location = element.location
    size = element.size

    driver.save_screenshot(path)

    left = location['x']
    top = location['y']
    right = location['x'] + size['width']
    bottom = location['y'] + size['height']

    image = Image.open(path)
    image_cropped = image.crop((left, top, right, bottom))
    image_cropped.save(path)

    captcha = pytesseract.image_to_string(image_cropped)
    captcha = captcha.replace(" ", "").strip()

    return captcha

# Capture captcha image and extract text
img_element = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_captchaimg"]')
captcha_path = "captcha.png"
captcha = get_captcha(driver, img_element, captcha_path)

# Load data from Excel file
dataframe = pd.read_excel('JEE.xlsx')

# Dictionary to store the results
results = {}

# Loop through each entry in the dataframe
for i, entry in dataframe.iterrows():
    try:
        # Fill the form fields
        application_number = str(entry['Application Number'])
        day = str(entry['Day'])
        month = str(entry['Month'])
        year = str(entry['Year'])
        
        Application_input = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtRegNo")
        Application_input.clear()
        Application_input.send_keys(application_number)

        day_input = driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$ddlday')
        day_input.send_keys(Keys.CONTROL + "a")
        day_input.send_keys(Keys.DELETE)
        day_input.send_keys(day)

        month_input = driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$ddlmonth')
        month_input.send_keys(Keys.CONTROL + "a")
        month_input.send_keys(Keys.DELETE)
        month_select = Select(month_input)
        month_select.select_by_visible_text(month)

        year_input = driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$ddlyear')
        year_input.send_keys(Keys.CONTROL + "a")
        year_input.send_keys(Keys.DELETE)
        year_input.send_keys(year)

        captcha_input = driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$Secpin')
        captcha_input.clear()
        captcha_input.send_keys(Keys.CONTROL + "a")
        captcha_input.send_keys(Keys.DELETE)
        captcha_input.send_keys(captcha)

        submit_btn = driver.find_element(By.CSS_SELECTOR, '#ctl00_ContentPlaceHolder1_Submit1')
        submit_btn.click()

        # Wait for the result page to load and become visible
        wait.until(EC.visibility_of_element_located((By.ID, 'some_unique_id_of_result_page_element')))

        # Capture screenshot of the entire browser page
        result_screenshot_path = f"result_{i}.png"
        result_screenshot = capture_full_browser_screenshot(driver, result_screenshot_path)

        # Save the result screenshot
        result_screenshot.save(result_screenshot_path)

        # Store the result in the dictionary
        results[i] = {
            'Application Number': application_number,
            'Result Screenshot': result_screenshot_path
        }

        # Reset the captcha input for the next iteration
        captcha_input.clear()

    except Exception as e:
        print(f"An error occurred for entry {i}: {str(e)}")

# Print the results dictionary
for key, value in results.items():
    print(f"Entry {key}: {value}")

# Save the final result screenshot as result.png
final_result_screenshot_path = "result.png"
capture_full_browser_screenshot(driver, final_result_screenshot_path)

# Close the driver after completing the process
driver.quit()
