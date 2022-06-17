"""
Tidy EVERYTHING up

While loop that runs while input != "Run please"
Stores the urls in an array

While loop that loops through the array and gets data on all of the matches

Add comments

Fix if it is SUS

Add in the 15+ column

"""

import requests
import openpyxl as xl
from bs4 import BeautifulSoup
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from datetime import datetime, timedelta, date
import shutil
import os

filename ="C:\\Users\\jhansen3\\OneDrive - KPMG\\Documents\\Python\\Gambling\\Disposals Tracking.xlsx"
#filename = "C:\\Users\\jayha\\Documents\\Gambling\\Automated\\Disposals Tracking.xlsx"

chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument("--disable-extensions")
chrome_options.add_argument("--incognito")
os.chmod('C:\\Users\\jhansen3\\OneDrive - KPMG\\Documents\\Python\\Gambling\\chromedriver_win32\\chromedriver.exe', 0o755)

driver = webdriver.Chrome(executable_path=r"C:\Users\jhansen3\OneDrive - KPMG\Documents\Python\Gambling\chromedriver_win32\chromedriver.exe")
#driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

driver.get('https://www.sportsbet.com.au/betting/australian-rules/afl/port-adelaide-v-sydney-6597348')

element = driver.find_elements(By.XPATH, '//*[@data-automation-id="market-group-accordion-header-title"]')

driver.implicitly_wait(10)

for i in range(0, len(element)):
    if element[i].text == "Top Markets":
        element[i].click()

driver.implicitly_wait(10)

elements = driver.find_elements(By.XPATH, '//*[@data-automation-id="market-group-accordion-header"]')

driver.implicitly_wait(10)

for i in range(0, len(elements)):
    if elements[i].text == "Disposal Markets":
        elements[i].click()

driver.implicitly_wait(10)




disposals = driver.find_elements(By.XPATH, '//*[@data-automation-id="accordion-header"]')

driver.implicitly_wait(10)

for i in range(0, len(disposals)):
    if disposals[i].text == "To Get 15 or More Disposals":
        disposals[i].click()

driver.implicitly_wait(10)

messy_players_15 = driver.find_elements(By.XPATH, '//*[contains(@data-automation-id, "-column-grid-outcome-name")]')
messy_odds_15 = driver.find_elements(By.XPATH, '//*[@data-automation-id="price-text"]')

players_15 = []
odds_15 = []
players_20 = []
odds_20 = []
players_25 = []
odds_25 = []
players_30 = []
odds_30 = []

for i in range(0, len(messy_players_15)):
    players_15.append(messy_players_15[i].text)
    odds_15.append(messy_odds_15[i].text)

driver.implicitly_wait(10)

for i in range(0, len(disposals)):
    if disposals[i].text == "To Get 15 or More Disposals":
        disposals[i].click()

driver.implicitly_wait(10)

for i in range(0, len(disposals)):
    if disposals[i].text == "To Get 20 or More Disposals":
        disposals[i].click()

driver.implicitly_wait(10)


messy_players_20 = driver.find_elements(By.XPATH, '//*[contains(@data-automation-id, "-column-grid-outcome-name")]')
messy_odds_20 = driver.find_elements(By.XPATH, '//*[@data-automation-id="price-text"]')

for i in range(0, len(messy_players_20)):
    players_20.append(messy_players_20[i].text)
    odds_20.append(messy_odds_20[i].text)

driver.implicitly_wait(10)

for i in range(0, len(disposals)):
    if disposals[i].text == "To Get 20 or More Disposals":
        disposals[i].click()

for i in range(0, len(disposals)):
    if disposals[i].text == "To Get 25 or More Disposals":
        disposals[i].click()

messy_players_25 = driver.find_elements(By.XPATH, '//*[contains(@data-automation-id, "-column-grid-outcome-name")]')
messy_odds_25 = driver.find_elements(By.XPATH, '//*[@data-automation-id="price-text"]')

for i in range(0, len(messy_players_25)):
    players_25.append(messy_players_25[i].text)
    odds_25.append(messy_odds_25[i].text)

for i in range(0, len(disposals)):
    if disposals[i].text == "To Get 25 or More Disposals":
        disposals[i].click()


for i in range(0, len(disposals)):
    if disposals[i].text == "To Get 30 or More Disposals":
        disposals[i].click()

messy_players_30 = driver.find_elements(By.XPATH, '//*[contains(@data-automation-id, "-column-grid-outcome-name")]')
messy_odds_30 = driver.find_elements(By.XPATH, '//*[@data-automation-id="price-text"]')

for i in range(0, len(messy_players_30)):
    players_30.append(messy_players_30[i].text)
    odds_30.append(messy_odds_30[i].text)

driver.quit()

games = 13

workbook = xl.load_workbook(filename)
sheet = workbook.worksheets[0]

for m in range(1, 1048576):
    if sheet.cell(m, 1).value is None:
        break
m=m-1

for i in range(0, len(players_20)):
    for a in range(1, m):
        if players_20[i] == sheet.cell(a, 1).value:
            sheet.cell(a, games+6).value = float(odds_20[i])


for i in range(0, len(players_25)):
    for a in range(1, m):
        if players_25[i] == sheet.cell(a, 1).value:
            sheet.cell(a, games+10).value = float(odds_25[i])


for i in range(0, len(players_30)):
    for a in range(1, m):
        if players_30[i] == sheet.cell(a, 1).value:
            sheet.cell(a, games+14).value = float(odds_30[i])
            

workbook.save("Disposals Tracking.xlsx")


