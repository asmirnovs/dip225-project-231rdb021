import imaplib
import email
import re, time
from datetime import datetime
import tabulate


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

print('Fetching email...')
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
print('Fetching done')
gmail.close()

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By

driver = webdriver.Chrome(service=Service(), options=webdriver.ChromeOptions())

cars = []
total = 0.00

for doc in res:
    driver.get(str(doc[0]))
    time.sleep(2)
    driver.find_element(By.NAME, "invoicePassword").send_keys(doc[1].strip())
    driver.find_element(By.CLASS_NAME, "mdc-button__label").click()
    time.sleep(2)

    subtotal = driver.find_elements(By.CSS_SELECTOR, '.text-right.emphasis.bolder')
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

driver.close()

print(*cars, sep="\n")
print(f"\n TOTAL COST: {round(total, 2)}€")