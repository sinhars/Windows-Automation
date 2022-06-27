import os
import sys
import pathlib
import json
import time
import datetime
import logging
import yagmail
import pywhatkit
import subprocess
from win32gui import (
    FindWindow,
    GetWindowRect,
    SetForegroundWindow,
    IsWindowVisible,
    EnumWindows,
    GetWindowText,
)
import pyautogui

from joblib import Parallel, delayed

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, SessionNotCreatedException


def read_credentials(key, file_name="credentials.json"):
    creds = json.load(open(file_name))
    return creds[key]


def get_apnacomplex_driver(config):
    options = Options()
    options.binary_location = config["chromeBinaryPath"]
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    try:
        driver = webdriver.Chrome(
            options=options, executable_path=config["chromeDriverExe"]
        )
    except SessionNotCreatedException:
        failure_msg = "Chrome driver is outdated. Please download the latest version from https://chromedriver.chromium.org/downloads"
        logging.error(failure_msg)
        # send_status_whatsapp(msg_text=failure_msg)
        quit()

    creds = read_credentials(key="apna-complex")
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
    WebDriverWait(driver, config["webDriverDelay"]).until(
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


def get_existing_bookings(config):
    # Initialize the webdriver and navigate to facilities page
    driver = get_apnacomplex_driver(config=config)
    # Get the court booking and viewing links from the facilities table
    court_links = get_court_links(
        driver=driver, delay=config["webDriverDelay"], court_num=None
    )
    viewing_links = court_links["viewing"]
    existing_bookings = dict(Court1=None, Court2=None)
    try:
        for court in viewing_links:
            existing_bookings[court] = get_active_bookings(
                driver=driver,
                delay=config["webDriverDelay"],
                viewing_url=viewing_links[court],
            )
    except Exception as ex:
        logging.error(f"Booking failed while checking exising bookings.")
        logging.error(ex)

    driver.close()
    driver.quit()
    return existing_bookings, court_links


def get_active_bookings(driver, delay, viewing_url):
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


def get_booking_time_slot(slot_hour):
    if slot_hour is None:
        slot_hour = int(datetime.datetime.now().strftime("%H"))
        current_min = int(datetime.datetime.now().strftime("%M"))
        if current_min > 50:
            slot_hour += 1

    booking_datetime = datetime.datetime.now() + datetime.timedelta(days=1)
    booking_datetime = booking_datetime.replace(
        hour=slot_hour, minute=0, second=0, microsecond=1
    )
    return slot_hour, booking_datetime


def select_court_num(existing_bookings, retries, num_slots):
    max_bookable_slots = 4
    already_used_slots = existing_bookings["Court1"] + existing_bookings["Court2"]
    avaiable_slots = max_bookable_slots - already_used_slots
    if avaiable_slots <= 0:
        return None

    num_slots = min(num_slots, avaiable_slots)
    court_num_list = [None] * num_slots

    new_bookings = existing_bookings.copy()
    courtPreference = ["Court1", "Court2"] if (retries < 2) else ["Court2", "Court1"]
    for i in range(num_slots):
        if new_bookings[courtPreference[0]] < 2:
            court_num_list[i] = int(courtPreference[0][-1])
            new_bookings[courtPreference[0]] += 1
        elif new_bookings[courtPreference[1]] < 2:
            court_num_list[i] = int(courtPreference[1][-1])
            new_bookings[courtPreference[1]] += 1

    return court_num_list


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


def load_apnacomplex_app(config, app_index):
    app_config = config["apnaComplexApps"][app_index]
    subprocess.Popen(app_config["launchPath"])
    time.sleep(config["sleepDuration"]["appLoad"])
    window_handle = FindWindow(None, app_config["windowName"])
    return window_handle


def navigate_to_booking(config, window_handle, booking_args):
    if not IsWindowVisible(window_handle):
        logging.error(f"Window handle {window_handle} not found - unable to navigate")
        return False

    SetForegroundWindow(window_handle)
    window_rectangle = GetWindowRect(window_handle)

    logging.info("Starting navigation to booking page")

    # Open facilities page
    pyautogui.moveTo(
        window_rectangle[0] + config["mousePosition"]["facilitiesButton"]["x"],
        window_rectangle[1] + config["mousePosition"]["facilitiesButton"]["y"],
    )
    pyautogui.click()
    time.sleep(config["sleepDuration"]["pageLoad"])

    # Click on page header before scrolling
    pyautogui.moveTo(
        window_rectangle[0] + config["mousePosition"]["facilitiesHeader"]["x"],
        window_rectangle[1] + config["mousePosition"]["facilitiesHeader"]["y"],
    )
    pyautogui.click()

    # Scroll to the bottom of the facilities page
    counter = 0
    while counter < config["scrollCount"]:
        counter += 1
        pyautogui.scroll(-config["scrollLength"])
        time.sleep(config["sleepDuration"]["smallPause"])

    # Click the tennis court facility icon
    court_num = booking_args["courtNum"]
    pyautogui.moveTo(
        window_rectangle[0]
        + config["mousePosition"][f"tennisCourt{court_num}Button"]["x"],
        window_rectangle[1]
        + config["mousePosition"][f"tennisCourt{court_num}Button"]["y"],
    )
    pyautogui.click()
    time.sleep(config["sleepDuration"]["pageLoad"])

    # Click slot booking button
    pyautogui.moveTo(
        window_rectangle[0] + config["mousePosition"]["slotBookingButton"]["x"],
        window_rectangle[1] + config["mousePosition"]["slotBookingButton"]["y"],
    )
    pyautogui.click()
    time.sleep(config["sleepDuration"]["pageLoad"])

    # Click tomorrow toggle
    pyautogui.moveTo(
        window_rectangle[0] + config["mousePosition"]["tomorrowToggle"]["x"],
        window_rectangle[1] + config["mousePosition"]["tomorrowToggle"]["y"],
    )
    pyautogui.click()
    time.sleep(config["sleepDuration"]["smallPause"])
    time.sleep(config["sleepDuration"]["smallPause"])

    # Drag  slots to bring the correct slot to starting position
    drag_counter = 0
    total_drags = booking_args["slotHour"] - config["initialSlotHour"]
    while drag_counter < total_drags:
        drag_counter += 1
        pyautogui.moveTo(
            window_rectangle[0] + config["mousePosition"]["timeSlotDrag"]["x"],
            window_rectangle[1] + config["mousePosition"]["timeSlotDrag"]["y"],
        )
        pyautogui.click()
        pyautogui.drag(
            -config["slotDragLength"],
            0,
            config["sleepDuration"]["smallPause"],
            button="left",
        )

    # Click the slot at the starting position
    pyautogui.moveTo(
        window_rectangle[0] + config["mousePosition"]["timeSlotButton"]["x"],
        window_rectangle[1] + config["mousePosition"]["timeSlotButton"]["y"],
    )
    pyautogui.click()

    # Click on book now button
    pyautogui.moveTo(
        window_rectangle[0] + config["mousePosition"]["bookNowButton"]["x"],
        window_rectangle[1] + config["mousePosition"]["bookNowButton"]["y"],
    )
    pyautogui.click()
    time.sleep(config["sleepDuration"]["smallPause"])
    return True


def confirm_booking(config, window_handle):
    if not IsWindowVisible(window_handle):
        logging.error(
            f"Window handle {window_handle} not found - unable to confirm booking"
        )
        return False

    SetForegroundWindow(window_handle)
    window_rectangle = GetWindowRect(window_handle)
    pyautogui.moveTo(
        window_rectangle[0] + config["mousePosition"]["confirmButton"]["x"],
        window_rectangle[1] + config["mousePosition"]["confirmButton"]["y"],
    )
    pyautogui.click()
    time.sleep(config["sleepDuration"]["smallPause"])
    return True


def navigate_to_home(config, window_handle):
    if not IsWindowVisible(window_handle):
        logging.error(f"Window handle {window_handle} not found - unable to close")
        return False

    SetForegroundWindow(window_handle)
    window_rectangle = GetWindowRect(window_handle)

    # Navigate back from booking page
    pyautogui.moveTo(
        window_rectangle[0] + config["mousePosition"]["bookingBackButton"]["x"],
        window_rectangle[1] + config["mousePosition"]["bookingBackButton"]["y"],
    )
    pyautogui.click()
    time.sleep(config["sleepDuration"]["pageLoad"])

    # Navigate back from facilities page
    pyautogui.moveTo(
        window_rectangle[0] + config["mousePosition"]["facilitiesBackButton"]["x"],
        window_rectangle[1] + config["mousePosition"]["facilitiesBackButton"]["y"],
    )
    pyautogui.click()
    time.sleep(config["sleepDuration"]["pageLoad"])

    # Click close window
    pyautogui.moveTo(
        window_rectangle[0] + config["mousePosition"]["closeWindowButton"]["x"],
        window_rectangle[1] + config["mousePosition"]["closeWindowButton"]["y"],
    )
    pyautogui.click()
    time.sleep(config["sleepDuration"]["smallPause"])

    # Click confirm close window
    pyautogui.moveTo(
        window_rectangle[0] + config["mousePosition"]["confirmCloseButton"]["x"],
        window_rectangle[1] + config["mousePosition"]["confirmCloseButton"]["y"],
    )
    pyautogui.click()
    time.sleep(config["sleepDuration"]["smallPause"])
    return True


def closeBlueStacksWindow(config):
    window_handle = FindWindow(None, config["blueStacksWindowName"])
    if not IsWindowVisible(window_handle):
        logging.error(f"BlueStacks window {window_handle} not found - unable to close")
        return False

    SetForegroundWindow(window_handle)
    window_rectangle = GetWindowRect(window_handle)
    time.sleep(config["sleepDuration"]["smallPause"])
    pyautogui.moveTo(
        window_rectangle[0] + config["mousePosition"]["blueStacksCloseButton"]["x"],
        window_rectangle[1] + config["mousePosition"]["blueStacksCloseButton"]["y"],
    )
    pyautogui.click()
    time.sleep(config["sleepDuration"]["smallPause"])
    return True


def minimizeAllWindows():
    pyautogui.keyDown("winleft")
    pyautogui.press("d")
    pyautogui.keyUp("winleft")
    screen_size = pyautogui.size()
    pyautogui.moveTo(int(screen_size[0] / 2), int(screen_size[1] / 2))
    pyautogui.click()
    return


def main():
    # Minimize all windows and click on desktop
    minimizeAllWindows()
    
    # Load config
    configFilePath = os.path.join(pathlib.Path(__file__).parent, "config.json")
    with open(configFilePath) as configJson:
        config = json.load(configJson)

    # Set logging config
    logging.basicConfig(filename="tennis-booking.log", level=logging.INFO)

    # Get existing booking counts for each court
    logging.info(f"Checking existing bookings.")
    slot_hour, booking_datetime = get_booking_time_slot(slot_hour=config["slotHour"])
    existing_bookings, court_links = get_existing_bookings(config=config)

    # Get list of available courts
    court_num_list = select_court_num(
        existing_bookings=existing_bookings, retries=0, num_slots=config["numSlots"]
    )
    if court_num_list is None:
        logging.error("Too many active bookings found. Can't book any more.")

    # Create booking arguments for each court booking
    all_booking_args = [
        dict(courtNum=court_num, slotHour=slot_hour) for court_num in court_num_list
    ]

    # Open app windows for each booking and navigate to confirm page
    logging.info("Initializing ApnaComplex app windows for booking.")
    all_window_handles = list()
    for idx, booking_args in enumerate(all_booking_args):
        window_handle = load_apnacomplex_app(config=config, app_index=idx)
        isSuccess = navigate_to_booking(
            config=config, window_handle=window_handle, booking_args=booking_args
        )
        if isSuccess:
            all_window_handles.append(window_handle)

    # Sleep till booking time arrives
    logging.info("Sleeping till booking time arrives.")
    time_to_booking = booking_datetime - datetime.datetime.now()
    sleep_time = 5
    while time_to_booking >= datetime.timedelta(hours=24):
        time.sleep(sleep_time)
        time_to_booking = booking_datetime - datetime.datetime.now()
        if time_to_booking < datetime.timedelta(hours=24, minutes=0, seconds=10):
            sleep_time = 0.001

    # Click confirm once booking time arrives
    logging.info("Confirming bookings.")
    for window_handle in all_window_handles:
        confirm_booking(config=config, window_handle=window_handle)
    time.sleep(config["sleepDuration"]["pageLoad"])

    # Navigate back to the home page and close the windows
    logging.info("Closing booking app windows.")
    for window_handle in all_window_handles:
        navigate_to_home(config=config, window_handle=window_handle)

    closeBlueStacksWindow(config=config)
    return


if __name__ == "__main__":
    main()
