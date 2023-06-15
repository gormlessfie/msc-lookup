from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from openpyxl import Workbook
from datetime import date

def click_booking_number(driver):
    # find & click booking number
    booking_option = driver.find_element(By.XPATH, "//label[@for='bookingradio']")
    booking_option.click()
    
# fill in input box w/ booking
def fill_input(driver, tracker):
    input_box = driver.find_element(By.XPATH, "//input[@id='trackingNumber']")
    input_box.send_keys(tracker)
    
    # press enter to search
    input_box.send_keys(Keys.ENTER)
    
def wait_for_content(driver, element):
    # Wait for the JavaScript to fill in elements
    wait = WebDriverWait(driver, 10)  # Maximum wait time of 10 seconds
    element_locator = (By.XPATH, element)
    wait.until(EC.presence_of_element_located(element_locator))
    
def retrieve_date_info(driver):
    # wait for page to load
    
    try:
        wait_for_content(driver, ".//span[@class='data-value']")
        
        eta_date_element = driver.find_element(By.XPATH, ".//span[@x-text='container.PodEtaDate']")
        return eta_date_element.text
    
    except NoSuchElementException:
        # Handle the case where the search element or any other expected element is not found
        return "Date not found"

    except Exception as e:
        # Handle any other unexpected exception
        print("An unexpected error occurred:", str(e))
        return "Date not found"
    
def clear_input_box(driver):
    input_box = driver.find_element(By.XPATH, "//input[@id='trackingNumber']")
    input_box.clear()


# obtain eta date
# put eta date into list
# create excel worksheet (openpyxl)
# append date into worksheet
 

# Create a new instance of the Firefox driver
driver = webdriver.Firefox()

# Open the website
driver.get('https://www.msc.com/en/track-a-shipment')

list_tracking_numbers = open("list.txt", "r").readlines()

click_booking_number(driver)

for entry in list_tracking_numbers: 
    fill_input(driver, entry)
    print(retrieve_date_info(driver))
    clear_input_box(driver)

driver.close()