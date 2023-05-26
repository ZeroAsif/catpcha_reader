from PIL import Image
from selenium import webdriver
from selenium.webdriver.common.by import By
import pytesseract
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
import numpy as np


# Tesseract  path
pytesseract.pytesseract.tesseract_cmd = r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe"


driver_path = "E:\\chromedriver_win32\\chromedriver.exe"


tesseract_path = r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe"

# Chrome driver 
service = Service(driver_path)


options = webdriver.ChromeOptions()



driver = webdriver.Chrome(service=service, options=options)


# NAVIGATING TO THE WEBPAGE
driver.get('https://ntaresults.nic.in/resultservices/JEEMAINauth23s2p1')


# LOCATION TO CAPTCHA
element_locator = (By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_captchaimg"]')
wait = WebDriverWait(driver, 30)

element = wait.until(EC.presence_of_element_located(element_locator))



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

    # OCR
    filename = 'captcha.png'
    img = np.array(Image.open(filename))
    text = pytesseract.image_to_string(img)
    print(text)
    

# Capture  captcha image
img_element = driver.find_element(By.XPATH, '/html/body/form/div[8]/div/div[1]/div[2]/table[2]/tbody/tr[2]/td[2]/img')
get_captcha(driver, img_element, "captcha.png")
