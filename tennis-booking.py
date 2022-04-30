import sys
import json
import time
import datetime
import logging
import yagmail
import pywhatkit

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, SessionNotCreatedException


def read_credentials(key, file_name="credentials.json"):
    creds = json.load(open(file_name))
    return creds[key]

def get_apnacomplex_driver(creds, delay):
    options = Options()
    options.binary_location = "C:/Program Files/Google/Chrome/Application/chrome.exe"
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    try:
        driver = webdriver.Chrome(
            options=options, executable_path="chrome-driver/chromedriver.exe"
        )
    except SessionNotCreatedException:
        failure_msg = "Chrome driver is outdated. Please download the latest version from https://chromedriver.chromium.org/downloads"
        logging.error(failure_msg)
        # send_status_whatsapp(msg_text=failure_msg)
        quit()

    # Navigate to facilities page url
    driver.get(creds["url"])
    # Enter email
    email_box = driver.find_element(by=By.ID, value="email")
    email_box.send_keys(creds["email"])
    # Enter password
    pwd_box = driver.find_element(by=By.ID, value="password")
    pwd_box.send_keys(creds["password"])
    # Submit login form
    pwd_box.submit()
    # Wait for page to load
    WebDriverWait(driver, delay).until(
        EC.presence_of_element_located((By.ID, "facilities"))
    )
    return driver


def get_court_links(driver, delay, court_num):
    court_links = dict(
        viewing=dict(Court1=None, Court2=None), booking=dict(Court1=None, Court2=None)
    )
    link_titles = dict(
        viewing="View bookings for this facility",
        booking="Make a booking for this facility",
    )

    facilities_table = driver.find_element_by_id("facilities")
    all_facility_rows = facilities_table.find_elements_by_xpath(".//tbody//tr")
    all_facility_rows.reverse()
    for row in all_facility_rows:
        all_cells = row.find_elements_by_xpath(".//td")
        booking_stage = "Searching for valid court"
        is_valid, current_court_num = is_valid_court(all_cells, court_num)
        if is_valid:
            all_links = all_cells[-1]
            all_links = all_links.find_elements_by_xpath(".//a")
            booking_stage = "Iterating all facility links"
            for current_link in all_links:
                image_title = current_link.find_elements_by_xpath(".//img")[
                    0
                ].get_attribute("title")
                booking_stage = "Identifying links"
                link_url = current_link.get_attribute("href")
                if image_title == link_titles["viewing"]:
                    court_links["viewing"][f"Court{current_court_num}"] = link_url
                elif image_title == link_titles["booking"]:
                    court_links["booking"][f"Court{current_court_num}"] = link_url

    return court_links

def get_existing_bookings(driver, delay, viewing_links, booking_date):
    existing_bookings = dict(Court1=None, Court2=None)
    try:
        for court in viewing_links:
            existing_bookings[court] = active_bookings(
                driver=driver,
                delay=delay,
                viewing_url=viewing_links[court],
                booking_date=booking_date,
            )
    except Exception as ex:
        logging.error(f"Booking failed while checking exising bookings.")
        logging.error(ex)

    return existing_bookings

def active_bookings(driver, delay, viewing_url, booking_date):
    def get_booking_count(booking_calendar, check_expired):
        booking_count = 0
        events_container = booking_calendar.find_elements_by_class_name(
            "fc-event-container"
        )
        all_events = events_container[-1].find_elements_by_class_name("fc-event")
        for booking_event in all_events:
            booking_apartment = booking_event.find_element_by_class_name(
                "fc-event-title"
            ).text
            booking_slot = booking_event.find_element_by_class_name(
                "fc-event-time"
            ).text
            if check_expired and int(booking_slot[0]) < datetime.datetime.now().hour:
                continue
            if "Bougainvillea-E-501" in booking_apartment:
                booking_count += 1
        return booking_count

    booking_count = 0
    try:
        driver.get(viewing_url)
        booking_calendar = WebDriverWait(driver, delay).until(
            EC.element_to_be_clickable((By.ID, "calendar"))
        )
        button_classes = ["fc-button-agendaDay", "fc-button-next"]
        for button_class in button_classes:
            check_expired = button_class == "fc-button-agendaDay"
            day_view_button = booking_calendar.find_element_by_class_name(button_class)
            day_view_button.click()
            time.sleep(1)
            booking_count += get_booking_count(
                booking_calendar=booking_calendar, check_expired=check_expired
            )

    except Exception as ex:
        logging.error("Unknown error occured during active booking checks.")
        logging.error(ex)

    return booking_count

def make_booking(driver, delay, booking_url, booking_date, date_string, slot_string):
    slot_booked = False
    try:
        driver.get(booking_url)
        # Check instructions checkbox
        instructions_checkbox = WebDriverWait(driver, delay).until(
            EC.element_to_be_clickable((By.ID, "read_instructions"))
        )
        instructions_checkbox.click()
        # Set booking date (today + 1)
        date_selector = driver.find_element(by=By.NAME, value="booking_date")
        date_selector.send_keys(date_string)
        # Set time slot
        slot_selector = driver.find_element_by_xpath(
            f"//select[@name='facility_time_slot_id']/option[text()='{slot_string}']"
        )
        slot_selector.click()
        # Wait for correct booking time
        time_to_booking = booking_date - datetime.datetime.now()
        if time_to_booking > datetime.timedelta(hours=25):
            logging.error("Booking time is more than 1 day in the future. ")
            return slot_booked
                
        while time_to_booking >= datetime.timedelta(hours=24):
            time.sleep(5)
            time_to_booking = booking_date - datetime.datetime.now()
        
        # Submit form
        submit_button = driver.find_element(by=By.NAME, value="make_booking")
        submit_button.submit()
        # Confirm submission
        confirm_button = WebDriverWait(driver, delay).until(
            EC.element_to_be_clickable((By.ID, "confirm"))
        )
        confirm_button.click()
        # Verify confirmation message
        status_message = WebDriverWait(driver, delay).until(
            EC.presence_of_element_located((By.ID, "status_message"))
        )
        slot_booked = status_message.text == "Booking completed successfully."

    except TimeoutException:
        logging.error("Booking page did not load correctly.")
        slot_booked = False
    except Exception as ex:
        logging.error("Unknown error occured during booking.")
        logging.error(ex)
        slot_booked = False

    return slot_booked


def get_booking_time_slot(slot_hour):
    if slot_hour is None:
        slot_hour = int(datetime.datetime.now().strftime("%H"))
        current_min = int(datetime.datetime.now().strftime("%M"))
        if current_min > 50:
            slot_hour += 1
    slot_string = str.zfill(str(slot_hour), 2)
    slot_string = "{}:00 - {}:45".format(slot_string, slot_string)
    return slot_hour, slot_string

def get_booking_date(slot_hour):
    booking_date = datetime.datetime.now() + datetime.timedelta(days=1)
    booking_date = booking_date.replace(
        hour=slot_hour, minute=0, second=1, microsecond=0
    )
    date_string = booking_date.strftime("%d/%m/%Y")
    return booking_date, date_string

def get_court_num(existing_bookings):
    if existing_bookings["Court1"] < 2:
        court_num = 1
    elif existing_bookings["Court2"] < 2:
        court_num = 2
    else:
        court_num = None
    return court_num

def is_valid_court(all_cells, court_num):
    facility_name = all_cells[0].text
    is_tennis_court = facility_name.startswith("Tennis Court")
    if not is_tennis_court:
        return False, 0
    current_court_num = int(facility_name.split(" ")[-1])
    is_valid = (court_num is None) or (court_num == current_court_num)
    return is_valid, current_court_num


def send_status_email(msg_text):
    gmail_creds = read_credentials(key="gmail")
    yag = yagmail.SMTP(user=gmail_creds["id"], password=gmail_creds["password"])
    yag.send(gmail_creds["id"], msg_text, msg_text)

def send_status_whatsapp(msg_text):
    try:
        whatsapp_creds = read_credentials(key="whatsapp")
        msg_hour = datetime.datetime.now().hour
        msg_min = datetime.datetime.now().minute + 1
        logging.info("Sending confirmation msg to %s" % whatsapp_creds["mobile"])
        pywhatkit.sendwhatmsg(
            phone_no=whatsapp_creds["mobile"],
            message=msg_text,
            time_hour=msg_hour,
            time_min=msg_min,
            wait_time=30,
            tab_close=True,
        )
    except:
        logging.error("Error sending whatsapp msg to %s" % whatsapp_creds["mobile"])


def main():
    logging.basicConfig(filename="tennis-booking.log", level=logging.INFO)
    
    DEFAULT_DELAY = 30
    NUM_SLOTS = 2
    MAX_RETRIES = 2
    SLOT_HOUR = None
    
    try:
        slot_hour, slot_string = get_booking_time_slot(slot_hour=SLOT_HOUR)
        booking_date, date_string = get_booking_date(slot_hour=slot_hour)
    except:
        logging.error("Invalid arguments provided!")

    # Initialize the webdriver and navigate to facilities page
    apna_complex_creds = read_credentials(key="apna-complex")
    driver = get_apnacomplex_driver(creds=apna_complex_creds, delay=DEFAULT_DELAY)

    # Get the court booking and viewing links from the facilities table
    court_links = get_court_links(driver=driver, delay=DEFAULT_DELAY, court_num=None)

    logging.info(f"Checking existing bookings.")
    
    # Get today's active booking counts for each court
    existing_bookings = get_existing_bookings(
        driver=driver,
        delay=DEFAULT_DELAY,
        viewing_links=court_links["viewing"],
        booking_date=booking_date,
    )
    
    retries = 0
    successfully_booked = 0
    while (successfully_booked < NUM_SLOTS) and (retries < MAX_RETRIES):
        retries += 1
        for i in range(NUM_SLOTS):
            logging.info(f"Booking slot {i + 1} of total {NUM_SLOTS} slots.")
            # Choose the appropriate court for booking
            court_num = get_court_num(existing_bookings=existing_bookings)
            if court_num is None:
                logging.error(
                    "You already have 4 active bookings. Can't book any more."
                )
                break
            # If Court1 failed first time and Court2 is available, try that instead
            if (retries > 1) and (court_num == 1) and existing_bookings["Court2"] < 2:
                court_num = 2
            
            logging_data = f"{date_string} at {slot_string} on Court {court_num}."
            logging.info(f"Initiating booking: {logging_data}")
            
            result = make_booking(
                driver=driver,
                delay=DEFAULT_DELAY,
                booking_url=court_links["booking"][f"Court{court_num}"],
                booking_date=booking_date,
                date_string=date_string,
                slot_string=slot_string,
            )

            if result:
                successfully_booked += 1
                existing_bookings[f"Court{court_num}"] += 1
                success_msg = f"Booking for slot {i + 1} successful: {logging_data}"
                logging.info(success_msg)
                # send_status_whatsapp(msg_text=success_msg)
            else:
                failure_msg = f"Booking failed for slot {i + 1}: {logging_data}"
                logging.error(failure_msg)
                # send_status_whatsapp(msg_text=failure_msg)


if __name__ == "__main__":
    main()
