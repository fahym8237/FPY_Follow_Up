import os
import glob
import subprocess
import tkinter
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException
import time

from config import date_from,date_to,source_file,source_file2,source_file3

Dynamic_Yield_Level = "https://ficosa.wats.com/dist/#/reporting/yield/dynamic-yield-report?id=Dynamic_Yieldb3a7e688-28b3-43ef-b912-953faf5486a6"
Dynamic_Yield_Retest = "https://ficosa.wats.com/dist/#/reporting/yield/dynamic-yield-report?id=Dynamic_Yield1fc37682-0892-45ae-bf53-8f511e064ef8"
Test_Reports = "https://ficosa.wats.com/dist/#/reporting/test-repair/uut-report"




def Exp_From_Wats(Dynamic_Yield_Link,label_text):
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=chrome_options)
    driver.get(Dynamic_Yield_Link)
    
    username = "abdelfattah.fahym@ficosa.com"
    password = "7140c9Wv$0O5rdR"
    

    from_date = f"{date_from} 00"
    to_date = f"{date_to} 23"
    wait = WebDriverWait(driver, 15)
    wait.until(EC.presence_of_element_located((By.ID, "inputUsername")))
    driver.find_element(By.ID, "inputUsername").send_keys(username)
    driver.find_element(By.ID, "inputPassword").send_keys(password)
    remember_me = driver.find_element(By.ID, "RememberMe")
    if not remember_me.is_selected():
        remember_me.click()
    driver.find_element(By.ID, "buttonLogin").click()
    time.sleep(1)
    print("WATS Login \u2714")

    driver.get(Dynamic_Yield_Link)
    #-----------
    max_retries = 3
    retry_count = 0
    success = False

    while retry_count < max_retries and not success:
        try:
            my_filters_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'My filters')]")))
            my_filters_button.click()
            time.sleep(3)  # Consider using a wait condition instead if possible
            
            fpy_button = wait.until(EC.element_to_be_clickable((By.XPATH, f"//span[text()='{label_text}']/parent::button")))
            fpy_button.click()

            success = True  # Mark as successful to exit loop
        except (TimeoutException, NoSuchElementException, ElementClickInterceptedException) as e:
            retry_count += 1
            print(f"Attempt {retry_count} failed with error: {e} {'\u2718'}")
            time.sleep(1)  # Wait before retrying

    
    #-----------
    time.sleep(1)  
    datetime_inputs = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//input[@placeholder='YYYY-MM-DD HH']")))
    datetime_inputs[0].clear()
    datetime_inputs[1].clear()
    time.sleep(1) 
    datetime_inputs[0].send_keys(from_date)
    datetime_inputs[1].send_keys(to_date)
    time.sleep(1)
    apply_filter_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Apply filter')]")))
    apply_filter_button.click()

    time.sleep(3) 
    # ✅ Get Results Number Text
    while True:
        try:
            result_div = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "report-count-info")))
            results_text = result_div.text.strip()
            results_number = int(results_text.split(":")[1].strip())
            print(results_number)
        except Exception as e:
            print(f"Failed to get results count: {e} ✘")
            results_number = 0  # Ensure the loop continues if an error occurs

        if results_number > 0:
            break

        time.sleep(3)

    
    export_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[.//img[contains(@src, 'export.svg')]]")))
    export_button.click()
    excel_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(., 'Excel') and .//img[contains(@src, 'excel.svg')]]")))
    excel_button.click()
    time.sleep(2)
    
    try:
        driver.quit()
    except Exception as e:
        print(f"Error closing driver: {e} {'\u2718'}")



def Remove_file(file_path):
    if os.path.exists(file_path):
        os.remove(file_path)
        print(f"Deleted: {file_path} {'\u2714'}")
    else:
        print("fpy.xlsx not found. {'\u2718'}")

def Rename_file(custom_name):
    download_dir = r"C:\Users\F6CAF02\Downloads"
    # --- Wait a few seconds to ensure download is complete ---
    time.sleep(5)
    # --- Find all Excel files ---
    excel_files = glob.glob(os.path.join(download_dir, "*.xlsx"))
    if not excel_files:
        print("No Excel files found. {'\u2718'}")
    else:
        # --- Get the most recently modified Excel file ---
        latest_file = max(excel_files, key=os.path.getmtime)
        # --- Define new name ---
        target_path = os.path.join(download_dir, custom_name)
        # --- Avoid overwrite ---
        if os.path.exists(target_path):
            os.remove(target_path)

        os.rename(latest_file, target_path)
        print(f"Renamed:\n{latest_file} -> {target_path}  {'\u2714'}")



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

    
if __name__ == "__main__":
    try:
        
        #--- FPY BY PROJECT ----
        print(f"--- FPY BY PROJECT ----")
        Remove_file(source_file2)
        Exp_From_Wats(Dynamic_Yield_Retest,"FPY")
        Rename_file("fbp.xlsx")
        run_script("fpy_by_project.py")
        #--- TIME SLEEP ---
        time.sleep(2)
        #--- FPY BY STATION ----
        print(f"--- FPY BY STATION ----")
        Remove_file(source_file)
        Exp_From_Wats(Dynamic_Yield_Level,"FPY")
        Rename_file("fbs.xlsx")
        run_script("fpy_by_station.py")
        #--- TIME SLEEP ---
        time.sleep(2)
        run_script("copy_data.py")
        #--- TIME SLEEP ---

        
        time.sleep(2)
        #--- RUNIN TEST REPORT ----
        #print(f"--- RUNIN TEST REPORT ----")
        #run_script("Wats_RunIn.py")

    except Exception as e:
        print(f"An error occurred: {e} {'\u2718'}")


