from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException
from selenium.webdriver.chrome.options import Options
from openpyxl import Workbook
import time

def get_businesses(zip_code):
    # Initialize the Chrome driver with headless options
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--window-size=2560,1440")
    driver = webdriver.Chrome(options=chrome_options)

    # driver = webdriver.Chrome()


    # Construct the URL for Google Maps search with the provided zip code
    base_url = f'https://www.google.com/maps/search/stone+countertops/@{zip_code},15z/data=!3m1!4b1!4m2!2m1!6e6'
    driver.get(base_url)

    # Set up an explicit wait to handle dynamic page elements
    wait = WebDriverWait(driver, 10)

    # Initialize variables for storing business data and controlling the scroll loop
    business_list = []
    unique_businesses = set()
    scroll_attempts = 0
    max_scroll_attempts = 50
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
                    if (name, zip_code) in unique_businesses:
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

                    # Save the business information if it's unique
                    if (name, address) not in unique_businesses:
                        business_list.append({'Name': name, 'Address': address})
                        unique_businesses.add((name, address))
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
    return business_list, len(business_list)

def save_to_excel(business_list):
    # Create an Excel workbook and add the business data
    wb = Workbook()
    ws = wb.active
    ws.append(["Business Name", "Address"])

    for business in business_list:
        ws.append([business['Name'], business['Address']])

    # Adjust column widths for better readability
    ws.column_dimensions['A'].width = 45
    ws.column_dimensions['B'].width = 60

    wb.save("Businesses.xlsx")

if __name__ == "__main__":
    # Main execution block
    zip_code = '77511'  # Example zip code for Houston
    print("Searching for stone countertop businesses in zip code", zip_code)

    businesses, address_count = get_businesses(zip_code)
    if businesses:
        print("Found", len(businesses), "businesses.")
        print("Addresses collected:", address_count)
        save_to_excel(businesses)
    else:
        print("No businesses found")
