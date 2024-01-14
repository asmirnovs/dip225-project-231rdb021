import imaplib
import email
import re, time
from datetime import datetime

gmail = imaplib.IMAP4_SSL("imap.gmail.com")
gmail.login(login, password)
gmail.select('INBOX')


# since = input("Enter date (dd-Mon-YYYY): ")
since = "06-Sep-2023"

value = 'info@citybee.lv'
print('Searching mail...')
_, data = gmail.search(None, 'FROM', value, f'(SINCE {str(since)})')
data = data[0].split()
print('Search done')

res = []

print('Fetching email...', end="")
for mail in data:
    _, msg = gmail.fetch(mail, '(RFC822)')
    msg = email.message_from_bytes(msg[0][1])

    if '=?UTF-8?B?csSTxLdpbnM=?=' in msg['subject']: # if "rēķins" in subject
        for part in msg.walk():
            if part.get_content_type() == 'text/plain':
                body = part.get_payload(None, True).decode()
                url = re.search("Pārskatīt rēķinu (.+)", body).group().lstrip("Pārskatīt rēķinu (").replace(" )", "").strip().strip(" ")
                parole = re.search("Rēķinu var apskatīt tikai pēc paroles ievadīšanas.+\n", body).group().lstrip("Rēķinu var apskatīt tikai pēc paroles ievadīšanas")[2:].strip()
                date = msg['date'][:-12] # get mail date
                date = datetime.strftime(datetime.strptime(date, "%a, %d %b %Y %H:%M:%S"), "%d.%m.%Y") # Mon, 21 Aug 2023 04:16:20 +0000 (UTC) -> 21.08.2023
                res.append([url, parole, date]) # add info to list
    print(".", end="", flush=True)
print('\nFetch done')
gmail.close()



from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, NoSuchElementException, NoSuchWindowException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

driver = webdriver.Chrome(service=Service(), options=webdriver.ChromeOptions())

cars = []
total = 0.00

for doc in res:
    driver.get(str(doc[0]))
    for _ in range(3): # try to get info 3 times before giving up
        try:
            driver.refresh()
            WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.NAME, "invoicePassword")))

            # enter password and open doc
            driver.find_element(By.NAME, "invoicePassword").send_keys(doc[1].strip())
            driver.find_element(By.CLASS_NAME, "mdc-button__label").click()
            WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".report-box.version2.mt-32.ng-star-inserted")))

            # get total cost
            subtotal = driver.find_elements(By.CSS_SELECTOR, '.text-right.emphasis.bolder')
            i = None
            for i in subtotal:
                if i.get_attribute("innerHTML") == "Kopā / Total incl. VAT, EUR":
                    break
            subtotal = subtotal[subtotal.index(i)+1].get_attribute("innerHTML") # TOTAL COST
            subtotal = float(subtotal)

            for i in driver.find_elements(By.CSS_SELECTOR, '.invoice-line.ng-star-inserted'):
                row = i.find_elements(By.TAG_NAME, "td")
                cars.append(f'{row[2].get_attribute("innerHTML")[:11].strip()}: {row[1].get_attribute("innerHTML")} -> {row[3].get_attribute("innerHTML")}€')
            cars.append(f"---- INVOICE SUBTOTAL: {round(subtotal, 2)}€ ----")
            total += round(subtotal, 2)
            break
        except (TimeoutException, NoSuchElementException):
            continue
        except (KeyboardInterrupt, NoSuchWindowException):
            print(*cars, sep="\n")
            print('Process stopped')
            quit()

driver.close()

print(*cars, sep="\n")
print(f"\n TOTAL COST: {round(total, 2)}€")