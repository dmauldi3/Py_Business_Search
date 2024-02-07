from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException
from openpyxl import Workbook

def get_businesses(zip_code):
    url = f'https://www.google.com/maps/search/stone+countertops/@{zip_code},15z/data=!3m1!4b1!4m2!2m1!6e6'
    print("Fetching URL:", url)

    driver = webdriver.Chrome()
    driver.get(url)
    wait = WebDriverWait(driver, 10)

    business_list = []
    index = 0


    while True:
        try:
            # Refetch the business links to avoid stale elements
            business_links = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a.hfpxzc")))
            print(f"Number of business links found: {len(business_links)}")
            if index >= len(business_links):
                break  # Exit if no more businesses

            try:
                business = business_links[index]
                name = business.get_attribute('aria-label').strip()
                print(f"Processing business: {name}")
                driver.execute_script("arguments[0].click();", business)

                try:
                    address_element = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "div.Io6YTe.fontBodyMedium.kR99db")))
                    address = address_element.text
                    print(f"Address found: {address}")
                except TimeoutException:
                    address = "Address not found"

                business_list.append({'Name': name, 'Address': address})
                driver.back()
                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "a.hfpxzc")))  # Wait for search results to reload
            except StaleElementReferenceException:
                print("Encountered a stale element, retrying...")
                continue  # If element is stale, retry with the next iteration

            index += 1
        except Exception as e:
            print(f"Error processing business: {e}")
            break

    driver.quit()
    return business_list, len(business_list)



def save_to_excel(business_list):
    wb = Workbook()
    ws = wb.active
    ws.append(["Business Name", "Address"])

    for business in business_list:
        ws.append([business['Name'], business['Address']])

    # Set the width of the columns
    ws.column_dimensions['A'].width = 45
    ws.column_dimensions['B'].width = 60

    wb.save("Businesses.xlsx")

if __name__ == "__main__":
    zip_code = '77043'  # Example zip code for Houston
    print("Searching for stone countertop businesses in zip code", zip_code)


    businesses, address_count = get_businesses(zip_code)
    if businesses:
        print("Found", len(businesses), "businesses.")
        print("Addresses collected:", address_count)
        save_to_excel(businesses)
    else:
        print("No businesses found.")
