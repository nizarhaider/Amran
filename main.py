from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from openpyxl import load_workbook
from selenium.webdriver.common.by import By
import time

# Replace with your login portal URL
url = "https://www.choiceadvantage.com/choicehotels/sign_in.jsp"

# Replace with your login credentials
username = "azubai.ky322"
password = "iPhone8s"

# Load Excel file
workbook = load_workbook(filename="sheet.xlsx")
worksheet = workbook.active

# Get the range of cells to read from
column_i = worksheet['I']
start_row = 9  # Row index starts at 1, so start at 9 instead of 8
end_row = len(list(filter(lambda x: x.value is not None, column_i)))

# Open browser and navigate to login portal
driver = webdriver.Chrome()
driver.get(url)

# Enter login credentials and submit
username_field = driver.find_element(By.NAME, "j_username")
password_field = driver.find_element(By.NAME, "j_password")
submit_button = driver.find_element(By.CLASS_NAME, "CHI_Button")
username_field.send_keys(username)
password_field.send_keys(password)
submit_button.click()

# Add time delay
time.sleep(10)



# Loop through each cell in the column and enter its value in the input field
for row_index in range(start_row, end_row+1):
    driver.get("https://www.choiceadvantage.com/choicehotels/Welcome.init")

    findReservation = driver.find_element(By.ID, "bannerFavButton_3")
    findReservation.click()

    # Get the cell value and remove commas

    cell_value = str(column_i[row_index-1].value).replace(",", "")
    time.sleep(5)
    confirmNumber = driver.find_element(By.NAME, "searchIdentifierNumber")
    confirmNumber.clear()
    confirmNumber.send_keys(cell_value)
    confirmNumber.send_keys(Keys.RETURN)
    time.sleep(5)

    guestFolio = driver.find_element(By.ID, "guestFolioEnabled")
    guestFolio.click()
    time.sleep(5)
    printReciept = driver.find_element(By.ID, "button_2")
    printReciept.click()
    time.sleep(5)





# Close the browser
driver.quit()
