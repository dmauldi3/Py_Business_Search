from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException, NoSuchElementException, \
    WebDriverException
from selenium.webdriver.chrome.options import Options
import pandas as pd
from openpyxl import load_workbook
from openpyxl import Workbook
import time
from geopy.geocoders import Nominatim
import geopy
import re


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
City = "Houston, TX"
latitude, longitude = get_lat_long(City)
if latitude is None or longitude is None:
    print("Stopping the execution as Latitude or Longitude couldn't be obtained")
    import sys
    sys.exit(0)

search_keyword = "granite countertops"  # Add this line to set your search keyword
print(f"Searching for {search_keyword} businesses in:", City)
Search_Attempts = 1



from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

# Maximum time to wait (in seconds)
MAX_WAIT = 10


def get_businesses(location, existing_data):
    latitude, longitude = get_lat_long(location)
    if latitude is None or longitude is None:
        return []

    # Initialize the Chrome driver with headless options
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--window-size=2560,1440")
    driver = webdriver.Chrome(options=chrome_options)
    driver.implicitly_wait(MAX_WAIT)

    # driver = webdriver.Chrome()

    new_business_count = 0  # Initialize a counter for new businesses

    # Construct the URL and print it for debugging
    base_url = f'https://www.google.com/maps/search/{search_keyword.replace(" ", "+")}/@{latitude},{longitude},15z/data=!3m1!4b1!4m2!2m1!6e6'
    print(f"URL: {base_url}")
    driver.get(base_url)

    # Initialize variables for storing business data and controlling the scroll loop
    business_list = []
    unique_businesses = set()
    scroll_attempts = 0
    max_scroll_attempts = Search_Attempts  # Maximum number of scrolled businesses searches. 10 = 100 businesses
    Time_delay = 3  # Time delay for content loading
    scroll_height = 2500  # Scroll height
    last_processed_index = -1  # Last processed business index

    while scroll_attempts < max_scroll_attempts:
        print(f"Scroll attempt: {scroll_attempts}")
        driver.execute_script(f"window.scrollBy(0, {scroll_height});")
        wait = WebDriverWait(driver, MAX_WAIT)

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

                        # Only add the business if it's not already known and the address format is valid
                    if (name, address) not in existing_data and (name, address) not in unique_businesses:
                        if is_valid_address_format(address):
                            business_list.append({'Name': name, 'Address': address})
                            unique_businesses.add((name, address))
                            new_business_count += 1
                            print(f"Collected: {name}, {address}")
                        else:
                            print(f"Invalid address format: {address}")

                    last_processed_index = i

                except StaleElementReferenceException:
                    print(f"Stale element encountered for business {name}, retrying...")
                    business_links = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a.hfpxzc")))
                    continue
                except TimeoutException:
                    print(f"Timeout encountered while fetching address for {name}")
                    continue
                except NoSuchElementException:
                    print("Couldn't find the web element.")
                except WebDriverException:
                    print("A WebDriver exception occurred.")

        except TimeoutException:
            print("Timeout waiting for new elements, attempting to scroll again.")

        new_height = driver.execute_script("return document.body.scrollHeight")

        scroll_attempts += 1

    driver.quit()
    return business_list, len(business_list), new_business_count


def is_valid_address_format(address):
    # Updated regex pattern for address validation
    pattern = r'^\d+.*,\s[A-Za-z ]+,\s[A-Za-z]{2}\s\d{5}(-\d{4})?$'
    return bool(re.match(pattern, address))


# Example usage
address = "123 Main Street, Anytown, NY 12345"
if is_valid_address_format(address):
    print(f"'{address}' is in a valid format.")
else:
    print(f"'{address}' is not in a valid format.")


def read_existing_data(file_path):
    try:
        # Read the Excel file into a DataFrame
        df = pd.read_excel(file_path, usecols=['Name', 'Address'])

        # Create a set of concatenated name and address for fast lookup
        existing_data = set(df.apply(lambda row: f"{row['Name']}{row['Address']}", axis=1))
        return existing_data

    except FileNotFoundError:
        print(f"FileNotFoundError {file_path}")
        return set()


def save_to_excel(business_list, file_path="Businesses.xlsx"):
    # Create a DataFrame from the new business list
    new_data_df = pd.DataFrame(business_list)

    try:
        # Read the existing Excel file into a DataFrame
        existing_df = pd.read_excel(file_path)

        # Combine new and existing data
        updated_df = pd.concat([existing_df, new_data_df], ignore_index=True)

        # Drop duplicates based on 'Name' and 'Address' columns
        updated_df.drop_duplicates(subset=['Name', 'Address'], keep='first', inplace=True)

    except FileNotFoundError:
        # If the file doesn't exist, use new data as the DataFrame
        updated_df = new_data_df

    # Save the updated DataFrame to an Excel file
    updated_df.to_excel(file_path, index=False)


if __name__ == "__main__":
    file_path = "Businesses.xlsx"
    existing_data = read_existing_data(file_path)
    initial_unique_count = len(existing_data)

    businesses, _, _ = get_businesses(City, existing_data)
    if businesses:
        save_to_excel(businesses, file_path)

        # read the data again after update
        updated_data = read_existing_data(file_path)
        print(f"Number of new businesses added: {len(updated_data) - initial_unique_count}")
    else:
        print("No new businesses found.")
