import time
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import sys

# Function to browse and select the Excel file
def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    file_entry.delete(0, tk.END)  # Clear any previous file path
    file_entry.insert(0, file_path)  # Insert selected file path

# Function to run the automation task
def run_script():
    excel_file = file_entry.get()
    username = username_entry.get()
    password = password_entry.get()

    if not excel_file or not username or not password:
        messagebox.showerror("Error", "Please provide all required fields (Excel file, username, and password).")
        return

    try:
        # Set up ChromeDriver (Ensure 'chromedriver.exe' is in the same directory or specify the full path)
        if getattr(sys, 'frozen', False):
            # Running as a bundled executable
            chromedriver_path = os.path.join(sys._MEIPASS, 'chromedriver.exe')
        else:
            # Running as a normal Python script
            chromedriver_path = os.path.join(os.getcwd(), 'chromedriver.exe')

        # Initialize the WebDriver
        service = Service(executable_path=chromedriver_path)
        driver = webdriver.Chrome(service=service)

        # Load Excel file and get task IDs and approval statuses using pandas
        df = pd.read_excel(excel_file)

        # Create a dictionary to map Task IDs to approval statuses
        tasks = {str(row['ID Task']): row['Metfone Approve / Not approve'] for index, row in df.iterrows()}

        # Open the login page
        driver.get("http://imt.metfone.com.kh/tmskhmer/account/?target=login")

        # Wait for the login page to load
        wait = WebDriverWait(driver, 15)

        # Login
        username_field = wait.until(EC.presence_of_element_located((By.NAME, "username")))
        password_field = wait.until(EC.presence_of_element_located((By.NAME, "pwd")))

        username_field.send_keys(username)
        password_field.send_keys(password)
        password_field.send_keys(Keys.RETURN)

        # Wait for the login process to complete
        wait.until(EC.presence_of_element_located((By.ID, "ext-gen15")))  # Ensure dashboard loads

        # Navigate through the menu
        internet_service = wait.until(EC.element_to_be_clickable((By.ID, "ext-gen15")))
        internet_service.click()

        fbb = wait.until(EC.element_to_be_clickable((By.ID, "ext-gen70")))
        fbb.click()

        daily = wait.until(EC.element_to_be_clickable((By.ID, "ext-comp-1226")))
        daily.click()

        # Loop through the tasks and check their status from the Excel data
        for task_id, approval_status in tasks.items():
            # Search for a task by ID
            id_task = wait.until(EC.presence_of_element_located((By.ID, "ms_task_id")))
            id_task.clear()  # Clear the input field before entering new task ID
            id_task.send_keys(task_id)
            btn_search = wait.until(EC.element_to_be_clickable((By.ID, "btn_search")))
            btn_search.click()

            # Wait for the search results to load
            time.sleep(1)

            # Locate all rows in the table
            rows = driver.find_elements(By.CSS_SELECTOR, "tr.light_row.w3-border.w3-hoverable.w3-hover-color-primary")

            # Loop through the rows to find the one that contains the specific value (task_id)
            for row in rows:
                # Get all columns in the row
                columns = row.find_elements(By.TAG_NAME, "td")

                # Check if the row contains the specific task ID
                if columns[2].text == task_id:  # Task ID is in the second column (index 1)
                    # Locate the select dropdown for Metfone status (Assumed to be in the 16th column)
                    metfone_dropdown = columns[16].find_element(By.CLASS_NAME, "my-select-box")

                    # Click on the dropdown to open it
                    metfone_dropdown.click()

                    # Select the appropriate approval option based on the Excel file
                    if approval_status == 'approve':
                        approve_option = columns[16].find_element(By.XPATH, ".//option[@value='approve']")
                        approve_option.click()
                    else:
                        not_approve_option = columns[16].find_element(By.XPATH, ".//option[@value='not_approve']")
                        not_approve_option.click()

                    # Click the confirm button to submit
                    btn_ok = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "swal-button--confirm")))
                    btn_ok.click()
                    break  # Exit the loop once the row is found and action is completed

        # Show success message
        messagebox.showinfo("Success", "Automation completed successfully!")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

    finally:
        # Ensure the browser closes properly
        time.sleep(10)
        driver.quit()

# Set up the main application window
root = tk.Tk()
root.title("Task Automation")
root.geometry("500x300")  # Set window size

# File selection label and entry field
file_label = tk.Label(root, text="Select Excel File:")
file_label.pack(padx=10, pady=5)

file_entry = tk.Entry(root, width=40)
file_entry.pack(padx=10, pady=5)

browse_button = tk.Button(root, text="Browse", command=browse_file)
browse_button.pack(padx=10, pady=5)

# Username label and entry field
username_label = tk.Label(root, text="Username:")
username_label.pack(padx=10, pady=5)

username_entry = tk.Entry(root, width=40)
username_entry.pack(padx=10, pady=5)

# Password label and entry field
password_label = tk.Label(root, text="Password:")
password_label.pack(padx=10, pady=5)

password_entry = tk.Entry(root, show="*", width=40)
password_entry.pack(padx=10, pady=5)

# Run button
run_button = tk.Button(root, text="Run Automation", command=run_script)
run_button.pack(padx=10, pady=20)

# Start the Tkinter event loop
root.mainloop()
