"""
NEXT STEPS:
Pull out the games being played that day and filter to show just them
    What ones were entered when user was prompted

Add a "Cleanse" option to remove all odds and fill again

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

urls = []
url = ""


while url != "Run Please":
    url = input("Enter the URL of the game: ")
    if url != "Run Please":
        urls.append(url)

j=0

for j in range(0, len(urls)):
    #Initiates all of the arrays
    players_15 = []
    odds_15 = []
    players_20 = []
    odds_20 = []
    players_25 = []
    odds_25 = []
    players_30 = []
    odds_30 = []
    players_35 = []
    odds_35 = []

    #Chomedriver options for some reason
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_argument("--incognito")
    os.chmod('C:\\Users\\jhansen3\\OneDrive - KPMG\\Documents\\Python\\Gambling\\chromedriver_win32\\chromedriver.exe', 0o755)


    #Switch the below comments when swapping laptop
    filename ="C:\\Users\\jhansen3\\OneDrive - KPMG\\Documents\\Python\\Gambling\\Disposals Tracking.xlsx"
    #filename = "C:\\Users\\jayha\\Documents\\Gambling\\Automated\\Disposals Tracking.xlsx"

    driver = webdriver.Chrome(executable_path=r"C:\Users\jhansen3\OneDrive - KPMG\Documents\Python\Gambling\chromedriver_win32\chromedriver.exe")
    #driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))


    #Driver to get the info from that specific URL
    driver.get(urls[j])


    #Finds and clicks the Top Markets then the Disposal Markets
    element = driver.find_elements(By.XPATH, '//*[@data-automation-id="market-group-accordion-header-title"]')

    for i in range(0, len(element)):
        if element[i].text == "Top Markets":
            element[i].click()

    elements = driver.find_elements(By.XPATH, '//*[@data-automation-id="market-group-accordion-header"]')

    for i in range(0, len(elements)):
        if elements[i].text == "Disposal Markets":
            elements[i].click()


    #Finds and clicks each of the disposal markets
    disposals = driver.find_elements(By.XPATH, '//*[@data-automation-id="accordion-header"]')

    for i in range(0, len(disposals)):
        if disposals[i].text == "To Get 15 or More Disposals":
            disposals[i].click()

    #Pulls the data into an array
    messy_players_15 = driver.find_elements(By.XPATH, '//*[contains(@data-automation-id, "-column-grid-outcome-name")]')
    messy_odds_15 = driver.find_elements(By.XPATH, '//*[@data-automation-id="price-text"]')


    #Appends the text to a new array
    for i in range(0, len(messy_players_15)):
        players_15.append(messy_players_15[i].text)
        if messy_odds_15[i].text != "SUS":
            odds_15.append(messy_odds_15[i].text)
        else:
            odds_15.append(0)

    #Closes the accordion heading
    for i in range(0, len(disposals)):
        if disposals[i].text == "To Get 15 or More Disposals":
            disposals[i].click()


    #As above
    for i in range(0, len(disposals)):
        if disposals[i].text == "To Get 20 or More Disposals":
            disposals[i].click()


    messy_players_20 = driver.find_elements(By.XPATH, '//*[contains(@data-automation-id, "-column-grid-outcome-name")]')
    messy_odds_20 = driver.find_elements(By.XPATH, '//*[@data-automation-id="price-text"]')

    for i in range(0, len(messy_players_20)):
        players_20.append(messy_players_20[i].text)
        if messy_odds_20[i].text != "SUS":
            odds_20.append(messy_odds_20[i].text)
        else:
            odds_20.append(0)

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
        if messy_odds_25[i].text != "SUS":
            odds_25.append(messy_odds_25[i].text)
        else:
            odds_25.append(0)

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
        if messy_odds_30[i].text != "SUS":
            odds_30.append(messy_odds_30[i].text)
        else:
            odds_30.append(0)

    for i in range(0, len(disposals)):
        if disposals[i].text == "To Get 30 or More Disposals":
            disposals[i].click()

    for i in range(0, len(disposals)):
        if disposals[i].text == "To Get 35 or More Disposals":
            disposals[i].click()

    messy_players_35 = driver.find_elements(By.XPATH, '//*[contains(@data-automation-id, "-column-grid-outcome-name")]')
    messy_odds_35 = driver.find_elements(By.XPATH, '//*[@data-automation-id="price-text"]')

    for i in range(0, len(messy_players_35)):
        players_35.append(messy_players_35[i].text)
        if messy_odds_35[i].text != "SUS":
            odds_35.append(messy_odds_35[i].text)
        else:
            odds_35.append(0)

    
    driver.quit()

    #Stores the data in the Disposals Tracking spreadsheet
    workbook = xl.load_workbook(filename)
    sheet = workbook.worksheets[0]

    i = 1

    games = 13
    
    #Finds the bottom cell
    for m in range(1, 1048576):
        if sheet.cell(m, 1).value is None:
            break
    m=m-1

    #Loops through and stores the value in the correct cell
    for i in range(0, len(players_15)):
        for a in range(1, m):
            if players_15[i] == sheet.cell(a, 2).value:
                sheet.cell(a, games+7).value = float(odds_15[i])

    for i in range(0, len(players_20)):
        for a in range(1, m):
            if players_20[i] == sheet.cell(a, 2).value:
                sheet.cell(a, games+11).value = float(odds_20[i])


    for i in range(0, len(players_25)):
        for a in range(1, m):
            if players_25[i] == sheet.cell(a, 2).value:
                sheet.cell(a, games+15).value = float(odds_25[i])


    for i in range(0, len(players_30)):
        for a in range(1, m):
            if players_30[i] == sheet.cell(a, 2).value:
                sheet.cell(a, games+19).value = float(odds_30[i])

    for i in range(0, len(players_35)):
        for a in range(1, m):
            if players_35[i] == sheet.cell(a, 2).value:
                sheet.cell(a, games+23).value = float(odds_35[i])
    j = j+1

workbook.save("Disposals Tracking.xlsx")


