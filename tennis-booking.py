import sys
import json
import datetime

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

def readCredentialsJSON(file_name="credentials.json", key="apna-complex"):
    creds = json.load(open(file_name))
    return (creds[key])

def bookTennisSlots(date_delta=1, slot_hour=None, court_num=None):
    apna_complex_creds = readCredentialsJSON()
    driver = get_apnacomplex_driver(creds=apna_complex_creds)
    delay = 5 #seconds
    
    try:
        facilities_table = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.ID, "facilities")))
        for row in facilities_table.find_elements_by_xpath(".//tbody//tr"):
            all_cells = row.find_elements_by_xpath(".//td")
            slot_booked = False
            if is_valid_court(all_cells, court_num):
                booking_links = all_cells[-1]
                all_links = booking_links.find_elements_by_xpath(".//a")
                for booking_link in all_links:
                    image_title = booking_link.find_elements_by_xpath(".//img")[0].get_attribute("title")
                    if image_title == "Make a booking for this facility":
                        booking_url = booking_link.get_attribute("href")
                        slot_booked = make_booking(creds=apna_complex_creds, booking_url=booking_url, \
                            date_delta=date_delta, slot_hour=slot_hour, delay=delay)
                        break
            if slot_booked:
                break
    except TimeoutException:
        print("Facilities page did not load correctly.")
    
    # time.sleep(delay)
    driver.quit()
    return (slot_booked)


def get_apnacomplex_driver(creds, url=None):
    options = Options()
    options.binary_location = "C:/Program Files/Google/Chrome/Application/chrome.exe"
    driver = webdriver.Chrome(options=options, executable_path="chrome-driver/chromedriver.exe")
    if url is None:
        url = creds["url"]
    # Navigate to url
    driver.get(url)
    # Enter email
    email_box = driver.find_element(by=By.ID, value="email")
    email_box.send_keys(creds["email"])
    # Enter password
    pwd_box = driver.find_element(by=By.ID, value="password")
    pwd_box.send_keys(creds["password"])
    # Submit login form
    pwd_box.submit()
    return (driver)


def make_booking(creds, booking_url, date_delta, slot_hour, delay):
    slot_booked = False
    try:
        driver = get_apnacomplex_driver(creds=creds, url=booking_url)
        # Check instructions checkbox
        instructions_checkbox = WebDriverWait(driver, delay).until(EC.element_to_be_clickable((By.ID, "read_instructions")))
        instructions_checkbox.click()
        # Set booking date (today + 1)
        date_selector = driver.find_element(by=By.NAME, value="booking_date")
        booking_date = (datetime.date.today() + datetime.timedelta(days=date_delta)).strftime("%d/%m/%Y")
        date_selector.send_keys(booking_date)
        # Set time slot
        if slot_hour is None:
            slot_hour = datetime.datetime.now().strftime("%H")
        time_slot = str.zfill(str(slot_hour), 2)
        time_slot  = "{}:00 - {}:45".format(time_slot, time_slot)
        slot_selector = driver.find_element_by_xpath("//select[@name='facility_time_slot_id']/option[text()='" + time_slot + "']")
        slot_selector.click()
        # Submit form
        submit_button = driver.find_element(by=By.NAME, value="make_booking")
        submit_button.submit()
        # Confirm submission
        confirm_button = WebDriverWait(driver, delay).until(EC.element_to_be_clickable((By.ID, "confirm")))
        confirm_button.click()
        # Verify confirmation message
        status_message = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.ID, "status_message")))
        print(status_message.text)
        slot_booked = (status_message.text == "Booking completed successfully.")
       
    except TimeoutException:
        print("Booking page did not load correctly.")
        slot_booked = False
    except:
        print("Some error occured during booking.")
        slot_booked = False
    finally:
        driver.quit()
    
    return (slot_booked)

def is_valid_court(all_cells, court_num):
    facility_name = all_cells[0].text
    is_tennis_court = facility_name.startswith("Tennis Court")
    is_valid_court_num = (court_num is None) or facility_name.endswith(str(court_num))
    return (is_tennis_court and is_valid_court_num)

if __name__ == "__main__":
    try:
        date_delta = int(sys.argv[1]) if (len(sys.argv) > 1) else 1
        slot_hour = int(sys.argv[2]) if (len(sys.argv) > 2) else None
        court_num = int(sys.argv[3]) if (len(sys.argv) > 3) else None
    except:
        print("Invalid arguments provided!")

    bookTennisSlots(date_delta=date_delta, slot_hour=slot_hour, court_num=court_num)