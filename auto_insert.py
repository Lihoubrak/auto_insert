import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd

# Set up the Service object for ChromeDriver
service = Service(executable_path="chromedriver.exe")

# Initialize the WebDriver
driver = webdriver.Chrome(service=service)

# Open the login page (replace with actual login URL)
driver.get("https://www.example.com/login")

# Wait for the login form to load
wait = WebDriverWait(driver, 10)

# Login (replace with actual login form fields and credentials)
username = wait.until(EC.presence_of_element_located((By.NAME, "username")))  # Replace with actual field name
password = wait.until(EC.presence_of_element_located((By.NAME, "password")))  # Replace with actual field name
username.send_keys("your_username")
password.send_keys("your_password")
password.send_keys(Keys.RETURN)

# Wait for the page after login to load
wait.until(EC.presence_of_element_located((By.ID, "dashboard")))  # Replace with an element ID that confirms successful login

# Read tasks from the Excel file using pandas
df = pd.read_excel("tasks.xlsx")  # Adjust path and sheet name if necessary

# Iterate over each task in the DataFrame
for index, row in df.iterrows():
    task_id = row['ID Task']  # Assuming 'ID Task' is the column with the task ID
    complaint = str(row['Complaint']).lower()  # Normalize for consistent comparison

    # Locate the table row for the task ID
    task_element = wait.until(
        EC.presence_of_element_located((By.XPATH, f"//tr[td[text()='{task_id}']]"))  # Adjust XPath as needed
    )

    # Find the dropdown element within the task row
    dropdown_element = task_element.find_element(By.ID, "approval-dropdown-id")  # Replace with the actual dropdown ID

    # Use Select to interact with the dropdown
    dropdown = Select(dropdown_element)

    # Select the appropriate option based on the condition
    if "approve" in complaint:
        dropdown.select_by_visible_text("Approve")  # Replace with the exact text of the "Approve" option
    else:
        dropdown.select_by_visible_text("Not Approve")  # Replace with the exact text of the "Not Approve" option

    # Wait for confirmation of the selection
    wait.until(EC.presence_of_element_located((By.ID, "success-message-id")))  # Replace with the actual confirmation ID

# Close the browser
driver.quit()
