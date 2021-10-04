from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time

import pandas as pd

# Excel & Chrome WebDriver (Requires local paths input)
# Load the excel document with contact columns: 'First', 'Email', 'Company'
excel_path = "../excel_sheets/py_excel_sample.xlsx"
df = pd.read_excel(excel_path)  # Put the excel data into a dataframe

# Load the downloaded chrome driver by following the local path it's located in
driver = webdriver.Chrome(executable_path="../drivers/chromedriver")

# Gmail Login information (Requires Username & Password)
url = 'https://accounts.google.com/signin/v2/identifier?passive=1209600&' \
      'continue=https%3A%2F%2Faccounts.google.com%2Fb%2F1%2FAddMailService&' \
      'followup=https%3A%2F%2Faccounts.google.com%2Fb%2F1%2FAddMailService&flow' \
      'Name=GlifWebSignIn&flowEntry=ServiceLogin'
driver.get(url)

# input email (above)
email = 'fake.coding.email@gmail.com'
email_input = driver.find_element_by_xpath('//*[@id="identifierId"]')
email_input.send_keys(email)

# Next button
button_next_1 = driver.find_element_by_xpath('//*[@id="identifierNext"]/div/button')
button_next_1.click()

# Wait (for browser to load)
driver.implicitly_wait(2)

# input password
password = '1fake.email1'
password_input = driver.find_element_by_xpath('//*[@id="password"]/div[1]/div/div[1]/input')
password_input.send_keys(password)

# Next button
button_next_2 = driver.find_element_by_xpath('//*[@id="passwordNext"]/div/button')
button_next_2.click()

# Wait (for browser to load)
driver.implicitly_wait(4)
time.sleep(3)

# Sets up the driver to use keyboard commands
action = webdriver.ActionChains(driver)

# Send emails
count = 0  # Variable that counts how many emails have been sent

for ind in df.index:  # Loops through the Excel data frame
     driver.implicitly_wait(2)
     time.sleep(1)
     # Counter pauses code for 5 seconds every 10 emails sent (to avoid bot detection)
     count += 1
     if count % 10 ==0:
         time.sleep(5)
         print("{} emails have been sent, will rest for 5 seconds.".format(count))

     # Compose message (using G-MAIL keyboard shortcut)
     action.send_keys('c').perform()
     time.sleep(2)

     # Email inputs
     recipient = df['Email'][ind]
     first_name = df['First Name'][ind]
     BCC = 'fake.coding.email@gmail.com'  # Optional, can input your own email to see if the emails were sent
     company = df['Company'][ind]

     # Input BCC email information
     action.send_keys(recipient).perform()  # Send recipient input
     driver.implicitly_wait(2)
     action.key_down(Keys.COMMAND).key_down(Keys.SHIFT).send_keys('B').perform()  # Keyboard shortcut to BCC
     driver.implicitly_wait(2)
     time.sleep(1)
     action.key_up(Keys.COMMAND).key_up(Keys.SHIFT).perform()  # Deactivate keyboard shortcut
     action.send_keys(BCC).perform()  # Send BCC input

     # This presses the TAB button twice (to get to SUBJECT line)
     action.key_down(Keys.TAB).key_up(Keys.TAB)
     action.key_down(Keys.TAB).key_up(Keys.TAB)

     # Subject line
     action.send_keys('Test Automation Email').perform()

     # This presses the TAB button once (to get the BODY MESSAGE)
     action.key_down(Keys.TAB).key_up(Keys.TAB)

     # BODY Message
     action.send_keys("Hi {}, \n".format(first_name),
                 "\n"
                 "I saw that you work at {} and I wanted to reach out to ".format(company),
                 "hear more about your experiences. Let me know if you want to chat, and I hope to hear back soon."
                 "\n \n"
                 "Regards,\n"
                 "Fake Name").perform()

     # Send the email
     action.key_down(Keys.COMMAND).key_down(Keys.ENTER).perform()  # Presses COMMAND and ENTER
     time.sleep(2)
     action.key_up(Keys.COMMAND).key_up(Keys.ENTER).perform()  # Releases COMMAND and ENTER

print("\n A total of {} emails were sent.".format(count))

# Open the last message on your inbox (which can be the last email you
# sent, if you included your own email in the BCC)
action.send_keys('o').perform()