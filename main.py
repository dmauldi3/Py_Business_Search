from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException
from selenium.webdriver.chrome.options import Options
import openpyxl
from openpyxl import load_workbook
from openpyxl import Workbook
import time
from geopy.geocoders import Nominatim
import geopy


def get_lat_long(location):
    try:
        geolocator = Nominatim(user_agent="unique_user_agent_for_your_application")
        location = geolocator.geocode(location)
        if location:
            return location.latitude, location.longitude
        else:
            print("Location not found")
            return None, None
    except geopy.exc.GeocoderServiceError as e:
        print(f"Geocoder service error: {e}")
        return None, None
    except Exception as e:
        print(f"Error during geocoding: {e}")
        return None, None



# Type in City, State or Zip to search area
City = "San Antonio, TX"
latitude, longitude = get_lat_long(City)
print("Searching for stone countertop businesses in:", City)



def get_businesses(location, existing_data):
    latitude, longitude = get_lat_long(location)
    if latitude is None or longitude is None:
        return []


    # Initialize the Chrome driver with headless options
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--window-size=2560,1440")
    driver = webdriver.Chrome(options=chrome_options)

     #driver = webdriver.Chrome()

    new_business_count = 0  # Initialize a counter for new businesses

    # Construct the URL and print it for debugging
    base_url = f'https://www.google.com/maps/search/stone+countertops/@{latitude},{longitude},15z/data=!3m1!4b1!4m2!2m1!6e6'
    print(f"URL: {base_url}")
    driver.get(base_url)

    # Set up an explicit wait to handle dynamic page elements
    wait = WebDriverWait(driver, 10)

    # Initialize variables for storing business data and controlling the scroll loop
    business_list = []
    unique_businesses = set()
    scroll_attempts = 0
    max_scroll_attempts = 50 # Maximum number of scrolled businesses searches. 10 = 100 businesses
    Time_delay = 3  # Time delay for content loading
    scroll_height = 2500  # Scroll height
    last_processed_index = -1  # Last processed business index

    while scroll_attempts < max_scroll_attempts:
        print(f"Scroll attempt: {scroll_attempts}")
        driver.execute_script(f"window.scrollBy(0, {scroll_height});")
        time.sleep(Time_delay)  # Wait for new content to load

        try:
            business_links = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a.hfpxzc")))
            print(f"Number of business links found: {len(business_links)}")

            # Start processing from the last processed index + 1
            for i in range(last_processed_index + 1, len(business_links)):
                try:
                    business = business_links[i]
                    name = business.get_attribute('aria-label').strip()
                    if (name, location) in unique_businesses:
                        continue  # Skip businesses already processed

                    # Scroll to and click on each business to load its details
                    driver.execute_script("arguments[0].scrollIntoView(true);", business)
                    time.sleep(Time_delay)
                    driver.execute_script("arguments[0].click();", business)
                    time.sleep(Time_delay)

                    # Extract the business address
                    address_element = wait.until(
                        EC.visibility_of_element_located((By.CSS_SELECTOR, "div.Io6YTe.fontBodyMedium.kR99db")))
                    address = address_element.text

                    if (name, address) in existing_data or (name, address) in unique_businesses:
                        continue  # Skip businesses already in the Excel file or already processed

                    # Save the business information if it's unique
                    if (name, address) not in unique_businesses:
                        business_list.append({'Name': name, 'Address': address})
                        unique_businesses.add((name, address))
                        new_business_count += 1  # Increment the new business counter
                        print(f"Collected: {name}, {address}")
                        last_processed_index = i  # Update last processed index

                except StaleElementReferenceException:
                    print(f"Stale element encountered for business {name}, retrying...")
                    business_links = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a.hfpxzc")))
                    continue
                except TimeoutException:
                    print(f"Timeout encountered while fetching address for {name}")
                    continue

        except TimeoutException:
            print("Timeout waiting for new elements, attempting to scroll again.")

        new_height = driver.execute_script("return document.body.scrollHeight")
        print(f"New height after scroll: {new_height}")

        scroll_attempts += 1

    driver.quit()
    return business_list, len(business_list), new_business_count

def read_existing_data(file_path):
    try:
        workbook = load_workbook(filename=file_path)
        sheet = workbook.active
        # Directly access the values as they are not Cell objects anymore
        existing_data = {(row[0], row[1]) for row in sheet.iter_rows(min_row=2, values_only=True)}
        return existing_data
    except FileNotFoundError:
        # File doesn't exist, return empty set
        print(f"FileNotFoundError {file_path}")
        return set()

def save_to_excel(business_list, file_path="Businesses.xlsx"):
    # Check if the Excel file already exists
    try:
        wb = load_workbook(file_path)
        ws = wb.active
    except FileNotFoundError:
        # Create a new workbook if the file doesn't exist
        wb = Workbook()
        ws = wb.active
        ws.append(["Business Name", "Address"])

    # Append new business data
    for business in business_list:
        ws.append([business['Name'], business['Address']])

    # Adjust column widths
    ws.column_dimensions['A'].width = 45
    ws.column_dimensions['B'].width = 60
    wb.save(file_path)

if __name__ == "__main__":
    file_path = "Businesses.xlsx"
    existing_data = read_existing_data(file_path)

    businesses, address_count, new_business_count = get_businesses(City, existing_data)
    if businesses:
        print(f"Found {len(businesses)} new businesses.")
        save_to_excel(businesses, file_path)
    else:
        print("No new businesses found.")
    print(f"Number of new addresses added: {new_business_count}")  # Print the count of new addresses added
