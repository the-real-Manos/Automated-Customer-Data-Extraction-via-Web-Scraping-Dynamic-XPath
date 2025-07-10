import time
import pandas as pd
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from webdriver_manager.chrome import ChromeDriverManager # Automatically manages chromedriver

# --- Configuration ---
# Credentials (Keep these secure, consider environment variables or a config file)
USERNAME = ""
PASSWORD = ""
LOGIN_URL = ""
PEOPLE_TAB_URL = ""

# Output file name
OUTPUT_EXCEL_FILE = "customer_data_all_pages.xlsx"

# --- Selectors ---
# Login Page Selectors (Verified previously)
LOGIN_USERNAME_SELECTOR = (By.NAME, "txtUser")
LOGIN_PASSWORD_SELECTOR = (By.NAME, "txtPwd")

# People Table Page Selectors (Verified previously)
PEOPLE_TABLE_LINK_SELECTOR = (By.CSS_SELECTOR, "#DataTables_Table_0 a[href^='newuser.php?mode=1&']")
NEXT_BUTTON_SELECTOR = (By.CSS_SELECTOR, "a.paginate_enabled_next")

# --- UPDATED Profile Page Data Selectors (Based on your latest HTML) ---
# Using XPath to find the div following the label containing specific text
PROFILE_NAME_SELECTOR = (By.XPATH, "//label[contains(text(), 'Name:')]/following::div[@class='five columns'][1]")
PROFILE_EMAIL_SELECTOR = (By.XPATH, "//label[contains(text(), 'Email:')]/following::div[@class='five columns'][1]")
PROFILE_ADDRESS_SELECTOR = (By.XPATH, "//label[contains(text(), 'Address:')]/following::div[@class='five columns'][1]")
# Targeting the link's text directly for telephone is cleaner
PROFILE_TELEPHONE_SELECTOR = (By.XPATH, "//label[contains(text(), 'Telephone:')]/following::div[@class='five columns'][1]//a[contains(@href, 'tel:')]")
# --- End Configuration ---


def setup_driver():
    """Sets up the Selenium WebDriver using webdriver-manager."""
    print("Setting up ChromeDriver using webdriver-manager...")
    try:
        service = Service(ChromeDriverManager().install())
        options = Options()
        # options.add_argument("--headless") # Uncomment to run without opening browser window
        # options.add_argument("--disable-gpu")
        # options.add_argument("--window-size=1920,1080")
        driver = webdriver.Chrome(service=service, options=options)
        driver.implicitly_wait(5)
        print("ChromeDriver setup successful.")
        return driver
    except Exception as e:
        print(f"Error setting up ChromeDriver with webdriver-manager: {e}")
        print("Please ensure you have Google Chrome installed and an internet connection.")
        raise

def login(driver, wait, username, password, login_url):
    """Logs into the website."""
    print(f"Navigating to login page: {login_url}")
    driver.get(login_url)
    try:
        print("Entering username...")
        wait.until(EC.visibility_of_element_located(LOGIN_USERNAME_SELECTOR)).send_keys(username)
        print("Entering password...")
        wait.until(EC.visibility_of_element_located(LOGIN_PASSWORD_SELECTOR)).send_keys(password + Keys.RETURN)
        print("Login submitted. Waiting for page load...")
        # IMPORTANT: Replace this sleep with a wait for an element that ONLY appears *after* successful login
        # Example: wait.until(EC.visibility_of_element_located((By.ID, "main-content-area")))
        # Example: wait.until(EC.visibility_of_element_located((By.LINK_TEXT, "Logout")))
        print("Waiting 5 seconds assuming login proceeds...") # Increase if needed, but explicit wait is better
        time.sleep(5)
        # You might need to check if the URL changed or if an error message appeared here
        print("Login successful (assumed based on wait time).")
    except TimeoutException:
        print("Login failed: Timed out waiting for login elements.")
        print(f"Ensure the selectors are correct:\n - Username: {LOGIN_USERNAME_SELECTOR}\n - Password: {LOGIN_PASSWORD_SELECTOR}")
        raise
    except Exception as e:
        print(f"An error occurred during login: {e}")
        raise

def get_all_customer_links(driver, wait, people_tab_url):
    """Navigates through pagination and collects all customer profile links."""
    print(f"Navigating to People tab starting page: {people_tab_url}")
    try:
        driver.get(people_tab_url)
    except Exception as e:
        print(f"Error navigating to People Tab URL ({people_tab_url}): {e}")
        print("Please check the PEOPLE_TAB_URL is correct and accessible after login.")
        return [] # Return empty list if navigation fails

    all_customer_links = []
    page_count = 1

    while True:
        print(f"--- Scraping links from Page {page_count} ---")
        try:
            # Wait for the specific table to be present first
            wait.until(EC.presence_of_element_located((By.ID, "DataTables_Table_0")))
            # Now wait for at least one link matching the *correct* selector within that table
            wait.until(EC.presence_of_element_located(PEOPLE_TABLE_LINK_SELECTOR))
            time.sleep(0.5) # Small delay might help ensure all links are available
            links_on_page = driver.find_elements(*PEOPLE_TABLE_LINK_SELECTOR)

            if not links_on_page:
                print(f" No customer links found on page {page_count} using selector {PEOPLE_TABLE_LINK_SELECTOR}.")
            else:
                page_links = []
                for link_element in links_on_page:
                    href = link_element.get_attribute("href")
                    if href:
                         page_links.append(href)

                print(f" Found {len(page_links)} links on page {page_count}.")
                all_customer_links.extend(page_links)

            # Attempt to find and click the 'Next' button
            try:
                next_button = wait.until(EC.element_to_be_clickable(NEXT_BUTTON_SELECTOR))
                print(" Found 'Next' button. Clicking...")

                if links_on_page:
                    element_to_check_staleness = links_on_page[0]
                else:
                    try: # Fallback: use the table itself for staleness check
                        element_to_check_staleness = driver.find_element(By.ID, "DataTables_Table_0")
                    except NoSuchElementException:
                         print("Could not find a stable element to check for staleness. Proceeding without check.")
                         element_to_check_staleness = None

                driver.execute_script("arguments[0].click();", next_button) # JS click

                if element_to_check_staleness:
                    print(" Waiting for next page to load (checking for staleness)...")
                    try:
                        wait.until(EC.staleness_of(element_to_check_staleness))
                        time.sleep(1.5) # Increased buffer after staleness confirmed
                    except TimeoutException:
                        print(" Warning: Staleness check timed out. Page might not have changed.")
                        print(" Assuming end of pagination due to staleness timeout.")
                        break
                else:
                    print(" Skipping staleness check. Waiting fixed time...")
                    time.sleep(3)

                page_count += 1

            except (NoSuchElementException, TimeoutException):
                print(" 'Next' button not found or not clickable. Assuming end of pagination.")
                break # Exit the loop

        except TimeoutException:
             print(f" Timed out waiting for table or links on page {page_count}.")
             print(f" Table ID selector: (By.ID, 'DataTables_Table_0')")
             print(f" Link selector: {PEOPLE_TABLE_LINK_SELECTOR}")
             break

        except Exception as e:
            print(f" An unexpected error occurred while processing page {page_count}: {e}")
            import traceback
            traceback.print_exc()
            print(" Stopping link collection due to error.")
            break

    print(f"\nFinished collecting links. Total links found: {len(all_customer_links)}.")
    unique_links = list(dict.fromkeys(link for link in all_customer_links if link))
    if len(unique_links) < len(all_customer_links):
        print(f" Removed duplicates, resulting in {len(unique_links)} unique links.")

    if not unique_links:
         print("\nWarning: No customer profile links were collected. Please RE-CHECK configuration and selectors.")

    return unique_links

def extract_customer_data(driver, wait, customer_links):
    """Visits each customer link and extracts data using updated selectors."""
    customer_data_list = []
    total_links = len(customer_links)
    print(f"\n--- Starting data extraction for {total_links} profiles ---")

    for i, link in enumerate(customer_links):
        if not link:
            print(f"Skipping empty link at index {i}.")
            continue

        print(f" Processing profile {i + 1}/{total_links}: {link}")
        try:
            driver.get(link)
            # Add a small wait for page elements to potentially render after load
            time.sleep(0.2)

            # Initialize fields
            name, email, address, telephone = "Not Found", "Not Found", "Not Found", "Not Found"

            # Extract Name
            try:
                # Wait slightly longer for the first element maybe
                name_element = wait.until(EC.visibility_of_element_located(PROFILE_NAME_SELECTOR))
                name = name_element.text.strip()
            except (NoSuchElementException, TimeoutException):
                 print(f"  - Name not found using selector {PROFILE_NAME_SELECTOR}")

            # Extract Email
            try:
                email_element = driver.find_element(*PROFILE_EMAIL_SELECTOR)
                email = email_element.text.strip()
            except NoSuchElementException:
                print(f"  - Email not found using selector {PROFILE_EMAIL_SELECTOR}")

            # Extract Address
            try:
                address_element = driver.find_element(*PROFILE_ADDRESS_SELECTOR)
                # .text should handle the <br> tags correctly, giving newlines
                address = address_element.text.strip()
            except NoSuchElementException:
                print(f"  - Address not found using selector {PROFILE_ADDRESS_SELECTOR}")

            # Extract Telephone
            try:
                telephone_element = driver.find_element(*PROFILE_TELEPHONE_SELECTOR)
                telephone = telephone_element.text.strip() # Get the text from the link
            except NoSuchElementException:
                print(f"  - Telephone not found using selector {PROFILE_TELEPHONE_SELECTOR}")

            customer_data_list.append({
                "Name": name,
                "Email": email,
                "Address": address,
                "Telephone": telephone, # Added telephone
                "Profile URL": link
            })

        except TimeoutException:
            print(f"  - Error: Timed out loading page or waiting for primary element on page: {link}")
            customer_data_list.append({"Name": "Error - Page/Element Timeout", "Email": "", "Address": "", "Telephone": "", "Profile URL": link})
        except Exception as e:
            print(f"  - Error: An unexpected error occurred extracting data for link: {link}. Error: {e}")
            customer_data_list.append({"Name": "Error - Extraction Failed", "Email": "", "Address": "", "Telephone": "", "Profile URL": link})
            # import traceback # Uncomment for debugging specific extraction errors
            # traceback.print_exc()


    print(f"\nFinished extracting data. Processed {len(customer_data_list)} profiles (check logs for 'Not Found' or 'Error' messages).")
    return customer_data_list

def save_to_excel(data_list, filename):
    """Saves the extracted data to an Excel file."""
    if not data_list:
        print("No data was extracted, skipping Excel save.")
        return
    print(f"Saving data to '{filename}'...")
    try:
        df = pd.DataFrame(data_list)
        # Ensure columns are in a consistent order, including Telephone
        columns_order = ["Name", "Email", "Telephone", "Address", "Profile URL"]
        # Add any columns that might be missing in the data and fill with default
        for col in columns_order:
            if col not in df.columns:
                df[col] = "" # Use empty string or "Not Found" as default
        df = df[columns_order] # Reorder/select columns

        df.to_excel(filename, index=False, engine='openpyxl') # Requires openpyxl: pip install openpyxl
        print(f"Data successfully saved to {os.path.abspath(filename)}")
    except ImportError:
         print("\nError: Could not save to Excel. The 'openpyxl' library is required.")
         print("Please install it by running: pip install openpyxl")
    except Exception as e:
        print(f"Error saving data to Excel file '{filename}': {e}")

# --- Main Execution ---
if __name__ == "__main__":
    start_time = time.time()
    driver = None
    print("--- Starting Web Scraping Process ---")
    try:
        driver = setup_driver()
        wait = WebDriverWait(driver, 20) # Reduced wait time slightly, increase if timeouts occur

        # 1. Login
        login(driver, wait, USERNAME, PASSWORD, LOGIN_URL)

        # 2. Get all customer links
        all_links = get_all_customer_links(driver, wait, PEOPLE_TAB_URL)

        # 3. Extract data from each link
        customer_data = []
        if all_links:
            customer_data = extract_customer_data(driver, wait, all_links)
        else:
            print("\nNo customer links were found. Cannot proceed to data extraction.")

        # 4. Save data to Excel
        save_to_excel(customer_data, OUTPUT_EXCEL_FILE)

    except Exception as e:
        print(f"\n--- An Unrecoverable Error Occurred During the Process ---")
        print(f"Error Type: {type(e).__name__}")
        print(f"Error Details: {e}")
        print("\n--- Traceback ---")
        import traceback
        traceback.print_exc()
        print("-----------------")

    finally:
        if driver:
            print("\nClosing the browser...")
            driver.quit()
            print("Browser closed.")

        end_time = time.time()
        print(f"\n--- Scraping Process Finished ---")
        print(f"Total execution time: {end_time - start_time:.2f} seconds")