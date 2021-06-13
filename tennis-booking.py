import json
import time

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

def readCredentialsJSON(file_name="credentials.json", key="apna-complex"):
    creds = json.load(open(file_name))
    return (creds[key])

def bookTennisSlots(slot_hours, court_num=0):
    apna_complex_creds = readCredentialsJSON()
    options = Options()
    options.binary_location = "C:/Program Files/Google/Chrome/Application/chrome.exe"
    driver = webdriver.Chrome(chrome_options=options, executable_path="chrome-driver/chromedriver.exe", )
    driver.get(apna_complex_creds["url"])
    # Enter email    
    email_box = driver.find_element(by=By.ID, value="email")
    email_box.send_keys(apna_complex_creds["email"])
    # Enter password
    pwd_box = driver.find_element(by=By.ID, value="password")
    pwd_box.send_keys(apna_complex_creds["password"])
    # Submit login form
    pwd_box.submit()
    delay = 5 #seconds
    try:
        facilities_table = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.ID, "facilities")))
        for row in facilities_table.find_elements_by_xpath(".//tbody//tr"):
            all_cells = row.find_elements_by_xpath(".//td")
            facility_name = all_cells[0].text
            is_tennis_court = facility_name.startswith("Tennis Court")
            is_valid_court_num = (court_num == 0) or facility_name.endswith(str(court_num))
            if is_tennis_court and is_valid_court_num:
                booking_links = all_cells[-1]
                all_links = booking_links.find_elements_by_xpath(".//a")
                for booking_link in all_links:
                    image_title = booking_link.find_elements_by_xpath(".//img")[0].get_attribute("title")
                    if image_title == "Make a booking for this facility":
                        booking_url = booking_link.get_attribute("href")
                        print(booking_url)



                        # Booking URL found - go to this url and make booking



                        break

        print("Page is ready!")
    except TimeoutException:
        print("Facilities page did not load correctly.")
    
    time.sleep(delay)
    driver.quit()  


if __name__ == "__main__":
    bookTennisSlots(slot_hours=[6, 7], court_num=1)