import os
import glob
import subprocess
import tkinter
from tkinter import messagebox
import datetime
import re

CONFIG_FILE = 'config.py'

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

def update_config_dates(start_date, end_date,nb_week):
    with open(CONFIG_FILE, 'r') as file:
        lines = file.readlines()

    # Update the date_from and date_to lines using regex
    new_lines = []
    for line in lines:
        if re.match(r'^date_from\s*=', line):
            new_lines.append(f'date_from = "{start_date}"\n')
        elif re.match(r'^date_to\s*=', line):
            new_lines.append(f'date_to = "{end_date}"\n')
        elif re.match(r'^nb_week\s*=', line):
            new_lines.append(f'nb_week = "{"W"}{nb_week}"\n')
        else:
            new_lines.append(line)

    with open(CONFIG_FILE, 'w') as file:
        file.writelines(new_lines)

# Get today's date
today = datetime.date.today()
previous_week_date = today - datetime.timedelta(days=7)

# Get ISO calendar values
iso_year, iso_week, iso_weekday = previous_week_date.isocalendar()

# Calculate start and end of previous week
start_of_prev_week = previous_week_date - datetime.timedelta(days=iso_weekday - 1)
end_of_prev_week = start_of_prev_week + datetime.timedelta(days=6)

# Format dates
start_str = start_of_prev_week.strftime("%Y-%b-%d")
end_str = end_of_prev_week.strftime("%Y-%b-%d")

# Output for logging/debug
print(f"Previous Week Number Selected: {iso_week}")
print(f"This Week Start From : {start_str} To {end_str}")

# Update the config.py file
update_config_dates(start_str, end_str,iso_week)
run_script("export_from_wats.py")
