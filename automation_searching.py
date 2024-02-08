from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd

# Read passport numbers from Excel file
passport_df = pd.read_excel("passport_data.xlsx")
passport_list = passport_df["Passport"].tolist()

# Initialize the Selenium WebDriver (Make sure to download the appropriate driver for your browser)
driver = webdriver.Chrome()

# Navigate to the website
driver.get("https://cekdptonline.kpu.go.id")

# Initialize an empty list to store the results
result_list = []

# Function to perform the search and handle the conditions
def perform_search(passport):
    search_bar = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "passport_number")))
    search_bar.clear()
    search_bar.send_keys(passport)
    search_bar.send_keys(Keys.RETURN)

    try:
        # Check if the result contains DPT or DPTK
        result = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, "//div[@class='mb-2']"))
        )
        if "DPT" in result.text:
            result_list.append({"Passport": passport, "Status": "DPT"})
        elif "DPTK" in result.text:
            result_list.append({"Passport": passport, "Status": "DPTK"})
        else:
            result_list.append({"Passport": passport, "Status": "Your data is Not Registered!"})
    except:
        result_list.append({"Passport": passport, "Status": "Your data is Not Registered!"})
    
    # Simulate pressing the back button to go back to the search page
    driver.execute_script("window.history.go(-1)")

# Loop through the passport list and perform searches
for passport in passport_list:
    perform_search(passport)

# Close the browser
driver.quit()

# Create a DataFrame from the result list
df = pd.DataFrame(result_list)

# Save the DataFrame to an Excel file
df.to_excel("passport_results.xlsx", index=False)

# Print the results
print(df)
