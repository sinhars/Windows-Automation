import sys
import json
import datetime
import logging
import yagmail
import pywhatkit

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException


def readCredentialsJSON(key, file_name="credentials.json"):
    creds = json.load(open(file_name))
    return (creds[key])


def bookTennisSlots(date_delta=1, slot_hour=None, court_num=None):
    slot_booked = False
    booking_stage = "Initializing ApnaComplex driver"
    apna_complex_creds = readCredentialsJSON(key="apna-complex")
    driver = get_apnacomplex_driver(creds=apna_complex_creds)
    delay = 60 #seconds
    try:
        booking_stage = "Waiting for facilities list"
        facilities_table = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.ID, "facilities")))
        for row in facilities_table.find_elements_by_xpath(".//tbody//tr"):
            all_cells = row.find_elements_by_xpath(".//td")
            booking_stage = "Searching for valid court"
            if is_valid_court(all_cells, court_num):
                booking_links = all_cells[-1]
                all_links = booking_links.find_elements_by_xpath(".//a")
                booking_stage = "Iterating all facility links"
                for booking_link in all_links:
                    image_title = booking_link.find_elements_by_xpath(".//img")[0].get_attribute("title")
                    booking_stage = "Identifying booking link"
                    if image_title == "Make a booking for this facility":
                        booking_url = booking_link.get_attribute("href")
                        booking_stage = "Initializing a new booking"
                        slot_booked = make_booking(creds=apna_complex_creds, booking_url=booking_url, \
                            date_delta=date_delta, slot_hour=slot_hour, delay=delay)
                        break
            if slot_booked:
                booking_stage = "Booking completed successfully"
                break
    except:
        logging.error("Booking failed at stage: %s" % booking_stage)
    
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


def get_booking_date(date_delta):
    return ((datetime.date.today() + datetime.timedelta(days=date_delta)).strftime("%d/%m/%Y"))


def get_booking_time_slot(slot_hour):
    if slot_hour is None:
            slot_hour = datetime.datetime.now().strftime("%H")
    time_slot = str.zfill(str(slot_hour), 2)
    time_slot  = "{}:00 - {}:45".format(time_slot, time_slot)
    return (time_slot)


def make_booking(creds, booking_url, date_delta, slot_hour, delay):
    slot_booked = False
    try:
        driver = get_apnacomplex_driver(creds=creds, url=booking_url)
        # Check instructions checkbox
        instructions_checkbox = WebDriverWait(driver, delay).until(EC.element_to_be_clickable((By.ID, "read_instructions")))
        instructions_checkbox.click()
        # Set booking date (today + 1)
        date_selector = driver.find_element(by=By.NAME, value="booking_date")
        date_selector.send_keys(get_booking_date(date_delta=date_delta))
        # Set time slot
        slot_selector = driver.find_element_by_xpath("//select[@name='facility_time_slot_id']/option[text()='" + get_booking_time_slot(slot_hour=slot_hour) + "']")
        slot_selector.click()
        # Submit form
        submit_button = driver.find_element(by=By.NAME, value="make_booking")
        submit_button.submit()
        # Confirm submission
        confirm_button = WebDriverWait(driver, delay).until(EC.element_to_be_clickable((By.ID, "confirm")))
        confirm_button.click()
        # Verify confirmation message
        status_message = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.ID, "status_message")))
        slot_booked = (status_message.text == "Booking completed successfully.")
       
    except TimeoutException:
        logging.error("Booking page did not load correctly.")
        slot_booked = False
    except:
        logging.error("Unknown error occured during booking.")
        slot_booked = False
    finally:
        driver.quit()
    return (slot_booked)


def is_valid_court(all_cells, court_num):
    facility_name = all_cells[0].text
    is_tennis_court = facility_name.startswith("Tennis Court")
    is_valid_court_num = (court_num is None) or facility_name.endswith(str(court_num))
    return (is_tennis_court and is_valid_court_num)


def send_status_email(msg_text):
    gmail_creds = readCredentialsJSON(key="gmail")
    yag = yagmail.SMTP(user=gmail_creds["id"], password=gmail_creds["password"])
    yag.send(gmail_creds["id"], msg_text, msg_text)
    

def send_status_whatsapp(msg_text):
    try:
        whatsapp_creds = readCredentialsJSON(key="whatsapp")
        msg_hour = datetime.datetime.now().hour
        msg_min = datetime.datetime.now().minute + 1
        logging.info("Sending confirmation msg to %s" % whatsapp_creds["mobile"])
        pywhatkit.sendwhatmsg(whatsapp_creds["mobile"], msg_text, msg_hour, msg_min, tab_close=True)
    except:
        logging.error("Error sending whatsapp msg to %s" % whatsapp_creds["mobile"])


def main():
    logging.basicConfig(filename="tennis-booking.log", level=logging.INFO)
    try:
        date_delta = int(sys.argv[1]) if (len(sys.argv) > 1) else 1
        slot_hour = int(sys.argv[2]) if (len(sys.argv) > 2) else None
        court_num = int(sys.argv[3]) if (len(sys.argv) > 3) else None
    except:
        logging.error("Invalid arguments provided!")
    
    logging.info("Initiating booking for %s at %s" %(get_booking_date(date_delta=date_delta), get_booking_time_slot(slot_hour=slot_hour)))
    result = bookTennisSlots(date_delta=date_delta, slot_hour=slot_hour, court_num=court_num)
    if result:
        success_msg = "Booking successfully completed for %s at %s" %(get_booking_date(date_delta=date_delta), get_booking_time_slot(slot_hour=slot_hour))
        logging.info(success_msg)
        send_status_email(msg_text=success_msg)
        send_status_whatsapp(msg_text=success_msg)
    else:
        # Send failure email 
        failure_msg = "Booking unsuccessful for %s at %s" %(get_booking_date(date_delta=date_delta), get_booking_time_slot(slot_hour=slot_hour))
        logging.warn(failure_msg)
        send_status_email(msg_text=failure_msg)
        send_status_whatsapp(msg_text=failure_msg)

if __name__ == "__main__":
    main()