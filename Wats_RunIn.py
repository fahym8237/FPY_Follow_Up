import subprocess
import tkinter
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException
import re
import time
import json
import os


from config import date_from, date_to

# Constants
USERNAME = ""
PASSWORD = ""
TEST_REPORT_URL = "https://ficosa.wats.com/dist/#/reporting/test-repair/uut-report"
JSON_PATH = "station_report.json"

STATION_NAMES = [
    "61000160_01", "61000160_02",
    "61000161_01", "61000161_02",
    "61000162_01", "61000162_02",
    "61000163_01", "61000163_02",
    "61000164_01", "61000164_02"
]

# Global dictionary to collect results
results_data = {}

def setup_driver():
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    return webdriver.Chrome(options=chrome_options)

def login_to_wats(driver, wait):
    driver.get(TEST_REPORT_URL)
    wait.until(EC.presence_of_element_located((By.ID, "inputUsername")))
    driver.find_element(By.ID, "inputUsername").send_keys(USERNAME)
    driver.find_element(By.ID, "inputPassword").send_keys(PASSWORD)

    remember_me = driver.find_element(By.ID, "RememberMe")
    if not remember_me.is_selected():
        remember_me.click()

    driver.find_element(By.ID, "buttonLogin").click()
    time.sleep(1)
    print("WATS Login ✔")

def apply_custom_filter(driver, wait, label_text):
    max_retries = 3
    retry_count = 0

    while retry_count < max_retries:
        try:
            my_filters = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'My filters')]")))
            my_filters.click()
            time.sleep(3)

            custom_filter = wait.until(EC.element_to_be_clickable(
                (By.XPATH, f"//span[text()='{label_text}']/parent::button")))
            custom_filter.click()
            return True

        except (TimeoutException, NoSuchElementException, ElementClickInterceptedException) as e:
            retry_count += 1
            print(f"Attempt {retry_count} failed: {e} ✘")
            time.sleep(1)

    return False

def set_date_range(wait, from_date, to_date):
    datetime_inputs = wait.until(EC.presence_of_all_elements_located(
        (By.XPATH, "//input[@placeholder='YYYY-MM-DD HH:mm']")))

    datetime_inputs[0].clear()
    datetime_inputs[1].clear()
    time.sleep(1)
    datetime_inputs[0].send_keys(from_date)
    datetime_inputs[1].send_keys(to_date)
    time.sleep(1)

def extract_report_data(driver, wait, station):
    try:
        time.sleep(4)
        display_info = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//div[contains(text(), 'Displaying') and contains(text(), 'reports')]")
        )).text

        values_divs = wait.until(EC.presence_of_all_elements_located(
            (By.XPATH, "//div[contains(@class, 'toolbar-text') and contains(@class, 'mat-tooltip-trigger')]")
        ))
        values = [v.text.strip() for v in values_divs[:3]]
        
        print(f"\n📄 Report Summary for {station}:")
        print(f"🗂 {display_info}")
        print(f"✅ Passed: {values[0]}")
        print(f"⚠️ Failed: {values[1]}")
        print(f"❌ Error: {values[2]}")
        
        raw_numbers = re.findall(r'\d[\d ]*\d|\d', display_info)
        #numbers = [int(num.replace(' ', '')) for num in raw_numbers]
        print(raw_numbers)
        # Store results
        results_data[station] = {
            "Units" :raw_numbers[0],
            "Passed": values[0],
            "Failed": values[1],
            "Error": values[2]
        }

    except Exception as e:
        print(f"❌ Failed to extract values for {station}: {e}")

def Exp_From_Wats(report_url, label_text):
    driver = setup_driver()
    wait = WebDriverWait(driver, 15)

    try:
        login_to_wats(driver, wait)
        driver.get(report_url)
        time.sleep(2)

        if not apply_custom_filter(driver, wait, label_text):
            print("❌ Failed to apply custom filter.")
            return
        
        time.sleep(2)
        from_datetime = f"{date_from} 00:00"
        to_datetime = f"{date_to} 23:00"
        set_date_range(wait, from_datetime, to_datetime)

        for station in STATION_NAMES:
            try:
                label = wait.until(EC.presence_of_element_located(
                    (By.XPATH, "//span[contains(text(), 'Station name')]")
                ))

                input_element = label.find_element(By.XPATH,
                    "./ancestor::div[contains(@class, 'wats-section-element-label')]/following-sibling::div//input")

                input_element.clear()
                time.sleep(2)
                input_element.send_keys(station)

            except Exception as e:
                print(f"⚠️ Station '{station}' input error: {e}")

            apply_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Apply filter')]")))
            apply_button.click()

            # Wait briefly for potential "Generating report" to appear
            time.sleep(2)

            try:
                # Wait up to 10 seconds for the "Generating report" span to appear
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'Generating report')]"))
                )

                # Poll until the span disappears
                while True:
                    try:
                        gen_report_element = driver.find_element(By.XPATH, "//span[contains(text(), 'Generating report')]")
                        if gen_report_element.text.strip() != "Generating report":
                            break
                        print("⏳ Still generating report...")
                        time.sleep(3)
                    except NoSuchElementException:
                        break  # The span disappeared
            except TimeoutException:
                pass  # "Generating report" never showed up

            # Proceed with extraction
            extract_report_data(driver, wait, station)

        # Save to JSON
        try:
            if os.path.exists(JSON_PATH):
                with open(JSON_PATH, "r") as infile:
                    existing_data = json.load(infile)
            else:
                existing_data = {}

            existing_data.update(results_data)

            with open(JSON_PATH, "w") as outfile:
                json.dump(existing_data, outfile, indent=4)

            print(f"\n✅ Report saved to {JSON_PATH}")

        except Exception as e:
            print(f"❌ Failed to save JSON: {e}")

    finally:
        try:
            driver.quit()
        except Exception as e:
            print(f"Error closing driver: {e} ✘")

# Function to run a script
def run_script(script_name):
    try:
        script_path = os.path.join(os.getcwd(), script_name)
        if not os.path.isfile(script_path):
            raise FileNotFoundError(f"{script_name} not found in the current directory.")
        subprocess.run(["python", script_path], check=True)
        #messagebox.showinfo("Success", f"{script_name} executed successfully.")
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Error", f"An error occurred while running {script_name}:\n{e}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred:\n{e}")

    # Initialize tkinter without showing a window
    tkinter.Tk().withdraw()

def Rest_JsonFile():
    
    # Load the JSON file
    with open("station_report.json", "r", encoding="utf-8") as f:
        data = json.load(f)

    # Reset "Passed", "Failed", and "Error" to "0" for each station
    for station, values in data.items():
        values["Units"] = "0"
        values["Passed"] = "0"
        values["Failed"] = "0"
        values["Error"] = "0"

    # Save the updated JSON back to the file
    with open("station_report.json", "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4)


 
# Run
if __name__ == "__main__":
    try:
        Rest_JsonFile()
        Exp_From_Wats(TEST_REPORT_URL, "RUNIN")
        time.sleep(2)
        run_script("RunIn_report.py")
    except Exception as e:
        print(f"An error occurred: {e} ✘")
