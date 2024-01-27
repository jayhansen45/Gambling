"""

"""


import requests
import openpyxl as xl
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from datetime import datetime, timedelta, date
import time
import shutil
import os
from selenium.webdriver.common.action_chains import ActionChains

#Bunch of options and shit for the webdriver
chrome_options = webdriver.ChromeOptions()
#chrome_options.binary_location = "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe"
chrome_options.binary_location = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
chrome_options.add_argument('--no-sandbox')
#chrome_options.headless = True
chrome_options.add_argument("--disable-extensions")
chrome_options.add_argument("--incognito")
driver = webdriver.Chrome(executable_path=r"C:\\Users\\jhansen3\\OneDrive - KPMG\\Documents\\Python\\Gambling\\Other\\Chrome Driver\\chromedriver.exe", options = chrome_options)
#driver = webdriver.Chrome(ChromeDriverManager(version='118.0.5993.70').install(), options = chrome_options)



#Work Laptop
filename = "C:\\Users\\jhansen3\\OneDrive - KPMG\\Documents\\Python\\Gambling\\FFlyer\\Data.xlsx"

workbook = xl.load_workbook(filename)
sheet = workbook['Data']


start = 290


for q in range(20, 200):
    #date
    filedate=date.today()+timedelta(days=q+3)
    day = (filedate.strftime('%m-%d-%Y'))
    
    #site = "https://book.virginaustralia.com/dx/VADX/#/flight-selection?journeyType=one-way&activeMonth="+day+"&locale=en-GB&awardBooking=false&class=First&ADT=1&CHD=0&INF=0&origin=MEL&destination=ATH&date="+day+"&promoCode=&execution=undefined"
    site = "https://book.virginaustralia.com/dx/VADX/#/date-selection?journeyType=one-way&activeMonth="+day+"&locale=en-GB&awardBooking=false&searchType=BRANDED&class=First&ADT=1&CHD=0&INF=0&origin=MEL&destination=ABE&promoCode=&direction=0&execution=undefined"
    driver.get(site)
    time.sleep(3)
    #For loop goes here
    for c in range(1+start, 633):
        #No flights catch
        print("Next Country")
        check = 0
        try:
            time.sleep(3)
            print("Are there flights?")
            flights = driver.find_element(By.XPATH, '//*[@data-translation="flightNotFound.title"]')
        except:
            print("Yes")
            #Stores country
            departing = driver.find_element(By.XPATH, '//*[@data-translation="summaryBar.vaDescription.fullOneWay"]')

            airport = (departing.text).split("Melbourne to ")
            final_destination = airport[1].split(", departing")
            
            print(final_destination[0])
            time.sleep(3)

            try:
                status = driver.find_element(By.XPATH, '//*[@data-translation="app.currency.FFCURRENCY.symbol"]')
                print("Switching to currency")
                clickable = driver.find_element(By.XPATH, '//*[@id="flight-selection-points-currency-toggle-0"]')
                clickable.click()
                time.sleep(3)

                try:
                    print("Are there flights?")
                    flights = driver.find_element(By.XPATH, '//*[@data-translation="flightNotFound.title"]')
                    check = 1
                except:
                    print("Yes")
                print("Is there more than one?")
                try:
                    print("Yes")
                    clickable = driver.find_element(By.XPATH, '//*[@id="dxp-sort-toggle-button"]')
                    clickable.click()
                    time.sleep(3)
                    print("Yes")
                    print("Sort by price")
                    clickable = driver.find_element(By.XPATH, '//*[@id="radio-flight-sort-0-price-asc"]')
                    clickable.click()
                    time.sleep(3)
                    clickable = driver.find_element(By.XPATH, '//*[@class="dxp-button va-modal-action-button spark-btn update action-primary secondary medium"]')
                    clickable.click()
                    time.sleep(3)

                    print("Get Dollars")
                    dollars_array = driver.find_elements(By.XPATH, '//*[@class="number"]')
                    for d in dollars_array:
                        if d.text != "" and d.text != "0" and d.text != " ":
                            if int(float(d.text.replace(",", "")))>0:
                                dollars = int(float(d.text.replace(",", "")))
                                break
                    time.sleep(3)


                except:
                    print("No")
                    print("Get Dollars")
                    dollars_array = driver.find_elements(By.XPATH, '//*[@class="number"]')
                    for d in dollars_array:
                        if d.text != "" and d.text != "0" and d.text != " ":
                            if int(float(d.text.replace(",", "")))>0:
                                dollars = int(float(d.text.replace(",", "")))
                                break
                    time.sleep(3)

                print("Switching to points")
                clickable = driver.find_element(By.XPATH, '//*[@id="flight-selection-points-currency-toggle-1"]')
                clickable.click()
                time.sleep(3)            

                print("In Points")
                print("Sort options")

                print("Is there more than one?")
                try:
                    print("Yes")
                    clickable = driver.find_element(By.XPATH, '//*[@id="dxp-sort-toggle-button"]')
                    clickable.click()
                    time.sleep(3)
                    print("Yes")
                    print("Sort by price")
                    clickable = driver.find_element(By.XPATH, '//*[@id="radio-flight-sort-0-price-asc"]')
                    clickable.click()
                    time.sleep(3)
                    clickable = driver.find_element(By.XPATH, '//*[@class="dxp-button va-modal-action-button spark-btn update action-primary secondary medium"]')
                    clickable.click()
                    time.sleep(3)
                    
                    print("Get points")
                    points = driver.find_elements(By.XPATH, '//*[@class="number"]')
                except:
                    print("No")
                    print("Get points")
                    points = driver.find_elements(By.XPATH, '//*[@class="number"]')
                    
                
            except:
                try:
                    print("Are there flights?")
                    flights = driver.find_element(By.XPATH, '//*[@data-translation="flightNotFound.title"]')
                except:
                    print("Is there more than one?")
                    try:
                        print("Yes")
                        clickable = driver.find_element(By.XPATH, '//*[@id="dxp-sort-toggle-button"]')
                        clickable.click()
                        time.sleep(3)
                        print("Yes")
                        print("Sort by price")
                        clickable = driver.find_element(By.XPATH, '//*[@id="radio-flight-sort-0-price-asc"]')
                        clickable.click()
                        time.sleep(3)
                        clickable = driver.find_element(By.XPATH, '//*[@class="dxp-button va-modal-action-button spark-btn update action-primary secondary medium"]')
                        clickable.click()
                        time.sleep(3)

                        print("Get Dollars")
                        dollars_array = driver.find_elements(By.XPATH, '//*[@class="number"]')
                        for d in dollars_array:
                            if d.text != "" and d.text != "0" and d.text != " ":
                                if int(float(d.text.replace(",", "")))>0:
                                    dollars = int(float(d.text.replace(",", "")))
                                    break
                        time.sleep(3)
                    except:
                        print("No")
                        print("Get Dollars")
                        dollars_array = driver.find_elements(By.XPATH, '//*[@class="number"]')
                        for d in dollars_array:
                            if d.text != "" and d.text != "0" and d.text != " ":
                                if int(float(d.text.replace(",", "")))>0:
                                    dollars = int(float(d.text.replace(",", "")))
                                    break
                        time.sleep(3)
                    
                    print("Switching to points")
                    clickable = driver.find_element(By.XPATH, '//*[@id="flight-selection-points-currency-toggle-1"]')
                    clickable.click()
                    time.sleep(3)            

                    print("In Points")
                    print("Sort options")

                    #MIGHT NEED A "ARE THERE FLIGHTS" HERE TBD

                    print("Is there more than one?")
                    try:
                        print("Yes")
                        clickable = driver.find_element(By.XPATH, '//*[@id="dxp-sort-toggle-button"]')
                        clickable.click()
                        time.sleep(3)
                        print("Yes")
                        print("Sort by price")
                        clickable = driver.find_element(By.XPATH, '//*[@id="radio-flight-sort-0-price-asc"]')
                        clickable.click()
                        time.sleep(3)
                        clickable = driver.find_element(By.XPATH, '//*[@class="dxp-button va-modal-action-button spark-btn update action-primary secondary medium"]')
                        clickable.click()
                        time.sleep(3)
                        
                        print("Get points")
                        points = driver.find_elements(By.XPATH, '//*[@class="number"]')
                    except:
                        print("No")
                        print("Get points")
                        points = driver.find_elements(By.XPATH, '//*[@class="number"]')



            if check == 0:
                #Copies values into excel
                i = 0

                #Finds first row that hasn't been used
                m=3

                for m in range(3, 1048576):
                    if sheet.cell(m, 2).value is None:
                        break

                print(final_destination[0])

                
                for a in points:
                    if a.text!="" and a.text!="0" and a.text!=" ":
                        sheet.cell(m+i, 2).value = final_destination[0]
                        sheet.cell(m+i, 3).value = day
                        sheet.cell(m+i, 6).value = dollars
                        if a.text.find('.') == -1:
                            sheet.cell(m+i, 4).value = a.text
                        else:
                            i = i-1
                            sheet.cell(m+i, 5).value = a.text
                        i = i+1

        else:
            print("No")

        print(" ")
        #Clicks edit search button
        clickable = driver.find_element(By.XPATH, '//*[@class="dxp-button va-edit-search-button is-unstyled"]')
        clickable.click()
        time.sleep(3)

        #Opens drop down box
        clickable = driver.find_element(By.XPATH, '//*[@id="arriving-airport0"]')
        clickable.click()
        time.sleep(3)

        #Goes to the next airport
        #clickable = driver.find_element(By.XPATH, '//*[@id="react-autowhatever-arriving-airport0-auto-suggest--item-2"]')
        search = '//*[@id="react-autowhatever-arriving-airport0-auto-suggest--item-'+ str(c) + '"]'
        clickable = driver.find_element(By.XPATH, search)
        clickable.click()
        time.sleep(3)
        clickable = driver.find_elements(By.XPATH, '//*[@id="dxp-page-navigation-continue-button"]')
        if len(clickable) == 1:
            clickable[0].click()
        else:
            clickable[1].click()
        time.sleep(3)

        start = 0

        clickable = driver.find_element(By.XPATH, '//*[@class="dxp-button va-modal-action-button action-primary secondary medium"]')
        clickable.click()
        time.sleep(3)
        workbook.save("Data.xlsx")


driver.quit()
workbook.save("Data.xlsx")


