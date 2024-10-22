from selenium import webdriver
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException, NoSuchElementException, \
    WebDriverException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as ec
import pandas as pd
import time
from geopy.geocoders import Nominatim
import geopy
import re
import tkinter as tk
from tkinter import filedialog, font
import pickle
import threading
from ttkthemes import ThemedStyle
from tkinter import ttk
from geopy.exc import GeocoderServiceError

# Define global variables for the UI Entry widgets
excelFilePath = ""
entry_location = None
entry_keyword = None
status_label = None
stop_event = threading.Event()


def get_lat_long(location):
    try:
        geolocator = Nominatim(user_agent="unique_user_agent_for_your_application")
        location = geolocator.geocode(location)
        if location:
            return location.latitude, location.longitude
        else:
            print("Location not found")
            set_status_message("Could not find City or Zip")
            return None, None
    except geopy.exc.GeocoderServiceError as geo_service_error:
        print(f"Geocoder service error: {geo_service_error}")
        return None, None
    except Exception as exception_e:
        print(f"Error during geocoding: {exception_e}")
        return None, None


# Number of times you want the program to scroll through pages
Search_Attempts = 50

# Maximum time to wait (in seconds)
MAX_WAIT = 10


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


def get_businesses(location, search_keyword, existing_data):
    global stop_flag
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
    set_status_message("Searching for businesses now")
    driver.get(base_url)

    # Initialize variables for storing business data and controlling the scroll loop
    business_list = []
    unique_businesses = set()
    scroll_attempts = 0
    max_scroll_attempts = Search_Attempts  # Maximum number of scrolled businesses searches. 10 = 100 businesses
    time_delay = 3  # Time delay for content loading
    scroll_height = 2500  # Scroll height
    last_processed_index = -1  # Last processed business index

    while not stop_event.is_set() and scroll_attempts < max_scroll_attempts:
        print(f"Scroll attempt: {scroll_attempts}")
        driver.execute_script(f"window.scrollBy(0, {scroll_height});")
        wait = WebDriverWait(driver, MAX_WAIT)

        try:
            business_links = wait.until(ec.presence_of_all_elements_located((By.CSS_SELECTOR, "a.hfpxzc")))
            print(f"Number of business links found: {len(business_links)}")

            # Start processing from the last processed index + 1
            for i in range(last_processed_index + 1, len(business_links)):
                if stop_event.is_set():
                    set_status_message("Stopping")
                    break
                name = None
                try:
                    business = business_links[i]
                    name = business.get_attribute('aria-label').strip()
                    if (name, location) in unique_businesses:
                        continue  # Skip businesses already processed

                    # Scroll to and click on each business to load its details
                    driver.execute_script("arguments[0].scrollIntoView(true);", business)
                    time.sleep(time_delay)
                    driver.execute_script("arguments[0].click();", business)
                    time.sleep(time_delay)

                    # Extract the business address
                    address_element = wait.until(
                        ec.visibility_of_element_located((By.CSS_SELECTOR, "div.Io6YTe.fontBodyMedium.kR99db")))
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
                            set_status_message(f"Found: {name}")
                        else:
                            print(f"Invalid address format: {address}")

                    last_processed_index = i

                except StaleElementReferenceException:
                    print(f"Stale element encountered for business {name}, retrying...")
                    business_links = wait.until(ec.presence_of_all_elements_located((By.CSS_SELECTOR, "a.hfpxzc")))
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


def read_existing_data(file_path):
    try:
        print(f"Reading Excel data from {file_path}")
        set_status_message("Checking excel data format")

        # Read the Excel file into a DataFrame
        df = pd.read_excel(file_path)

        # Check if the DataFrame is empty
        if df.empty:
            print('The Excel file is blank')
            return set()

        # Rename the relevant columns if they exist
        column_mapping = {}
        if 'Your Current Name Column Header' in df.columns:
            column_mapping['Your Current Name Column Header'] = 'Name'
        if 'Your Current Address Column Header' in df.columns:
            column_mapping['Your Current Address Column Header'] = 'Address'
        df.rename(columns=column_mapping, inplace=True)

        print("Renamed the relevant columns to 'Name' and 'Address'")
        set_status_message("Renamed Excel columns")

        # Constructing the existing_data set
        existing_data = set(df['Name'] + ' ' + df['Address'])

        return existing_data
    except FileNotFoundError:
        print(f"FileNotFoundError: {file_path}")
        return set()
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        return set()


def save_to_excel(business_list, file_path="Businesses.xlsx"):
    try:
        if not file_path.endswith('.xlsx'):
            print("Invalid file path provided.")
            set_status_message("Invalid Excel file path")
            return
        new_data_df = pd.DataFrame(business_list)

        try:
            existing_df = pd.read_excel(file_path)
            updated_df = pd.concat([existing_df, new_data_df], ignore_index=True)
            updated_df.drop_duplicates(subset=['Name', 'Address'], keep='first', inplace=True)

        except FileNotFoundError:
            updated_df = new_data_df

        # print(f"Data to be written to Excel: \n{updated_df}")  # Add this
        updated_df.to_excel(file_path, index=False)
        print("Data has been written to Excel.")  # And this

    except Exception as e:
        print("Error occurred in save_to_excel: ", repr(e))


def set_status_message(message):
    def _set_status_message():
        status_label.config(text=message)

    root.after(100, _set_status_message)


def run_script():
    global excelFilePath, entry_location, entry_keyword, run_button, stop_event
    # If the task is not running, start it
    if run_button["text"] == "Search":
        stop_event.clear()
        run_button.config(text="Stop")
        threading.Thread(target=run_script_thread, args=()).start()
    # If the task is running, stop it
    elif run_button["text"] == "Stop":
        stop_event.set()


def run_script_thread():
    global excelFilePath, entry_location, entry_keyword

    try:
        # Get user input
        search_keyword = entry_keyword.get()
        city = entry_location.get()

        latitude, longitude = get_lat_long(city)
        if latitude is None or longitude is None:
            print("Stopping the execution as Latitude or Longitude couldn't be obtained")
            return

        # Existing data
        existing_data = read_existing_data(excelFilePath)
        # print(f"Existing data: {existing_data}") # added prints for debugging
        initial_unique_count = len(existing_data)

        # Get businesses
        businesses, _, _ = get_businesses(city, search_keyword, existing_data)
        print(f"Businesses: {businesses}")  # added prints for debugging

        # Save, update data and print results
        if businesses:
            # Save locations, keywords, and file paths to a pickle file
            with open('previous_inputs.pkl', 'wb') as f:
                pickle.dump({
                    'location': city,
                    'keyword': search_keyword,
                    'excel_file_path': excelFilePath,
                }, f)
                print(f"Data has been saved to previous_inputs.pkl with the excel file path: {excelFilePath}")

            if excelFilePath != '':
                save_to_excel(businesses, excelFilePath)
                updated_data = read_existing_data(excelFilePath)
                print(f"Number of new businesses added: {len(updated_data) - initial_unique_count}")
                # messagebox.showerror("Error", f"An error occurred: {e}")
                set_status_message(f"Number of new businesses added: {len(updated_data) - initial_unique_count}")

            else:
                print("No new businesses found.")
        else:
            print("No new businesses found.")
    except Exception as e:
        print(f"Exception in run_script_thread: {e}")
    finally:
        # Re-enable the button when the script is done
        def reenable_button():
            run_button.config(state="normal")
        root.after(100, reenable_button)
        def set_button_search():
            run_button.config(text="Search")
            run_button.config(state="normal")  # Enable the button for searching again
        root.after(100, set_button_search)


if __name__ == "__main__":
    # Create the main Tkinter window
    root = tk.Tk()
    root.geometry('350x250')  # Set window size
    root.title("Py Business Search")  # Set the window title

    try:
        root.iconbitmap(r'C:\Users\danie\PycharmProjects\pythonProject\ICON.ico')
    except Exception as e:
        print(f"Error setting icon: {e}")

    style = None

    try:
        root.configure(background='#353535')
        style = ThemedStyle(root)
        style.set_theme("equilux")
        # Change the background of the highlight part and the color of the text
        style.configure('Custom.TEntry', fieldbackground='white', foreground='white')

    except Exception as e:
        print(f"Error setting theme: {e}")

    # Set custom styles
    custom_font = font.nametofont("TkDefaultFont")
    custom_font.configure(size=14)
    style.configure('Custom.TLabel', background='#353535', foreground='white', font=custom_font)
    style.configure('Custom.TButton', font=custom_font)
    style.configure('Custom.TEntry', font=custom_font)

    # Configure grid
    root.columnconfigure(0, weight=1)
    for i in range(6):
        root.rowconfigure(i, weight=1)

    # ~~~~~~~~~~~~Create widgets~~~~~~~~~~~~~~
    # Keyword Header
    ttk.Label(root, text="Keyword:", style='Custom.TLabel').grid(row=0, column=0, sticky='ew')
    # Keyword Textbox
    entry_keyword = ttk.Entry(root, style='Custom.TEntry', width=20)
    entry_keyword.grid(row=1, column=0, padx=10, pady=5, sticky='ew')

    # Location Header
    ttk.Label(root, text="Location:", style='Custom.TLabel').grid(row=2, column=0, sticky='ew')
    # Location Textbox
    entry_location = ttk.Entry(root, style='Custom.TEntry', width=20)
    entry_location.grid(row=3, column=0, padx=10, pady=5, sticky='ew')
    entry_location['state'] = 'normal'  # or 'readonly', 'disabled'
    entry_location.focus_set()

    status_label = ttk.Label(root, text="", style='Custom.TLabel')
    status_label.grid(row=4, column=0, sticky='ew')

    # Load previous inputs if they exist
    previous_inputs = {'location': '', 'keyword': '', 'excel_file_path': ''}
    try:
        with open('previous_inputs.pkl', 'rb') as f:
            previous_inputs = pickle.load(f)
    except (FileNotFoundError, pickle.UnpicklingError):
        pass

    entry_keyword.insert(0, previous_inputs['keyword'])
    entry_location.insert(0, previous_inputs['location'])
    excelFilePath = previous_inputs['excel_file_path']


    def set_excel_file():
        global excelFilePath
        temp_file_path = tk.filedialog.askopenfilename(initialdir="/", title="Select file",
                                                       filetypes=(("Excel files", "*.xlsx"), ("all files", "*.*")))
        if temp_file_path != '':
            excelFilePath = temp_file_path
        print(f"Chosen file path: {excelFilePath}")


    Set_button = ttk.Button(root, text="Set Excel File Path", command=set_excel_file, style='Custom.TButton')
    Set_button.grid(row=5, column=0, padx=10, pady=0, sticky='ew')

    run_button = ttk.Button(root, text="Search", command=run_script, style='Custom.TButton')
    run_button.grid(row=6, column=0, padx=10, pady=10, sticky='ew')

    root.mainloop()
