import openpyxl as xl
import requests
import statistics
import numpy as np
from scipy.stats import norm
from bs4 import BeautifulSoup

webpage_response = requests.get('https://afltables.com/afl/stats/teams/adelaide/2022_gbg.html')
webpage = webpage_response.content
soup = BeautifulSoup(webpage, "html.parser")

table = soup.find_all(attrs={'class':'sortable'})

messy_disposals = table[0].find_all('td')
messy_tog = table[21].find_all('td')

disposals = []
tog = []

games = 12

for a in messy_disposals:
    if (a.text == '\xa0' or a.text == '-'):
        a = 0
        disposals.append(a)
    else:
        disposals.append(a.text)

for a in messy_tog:
    if (a.text == '\xa0' or a.text == '-'):
        a = 0
        tog.append(a)
    else:
        tog.append(a.text)

players = []

i=0

while (i < len(disposals)):
    players.append(disposals.pop(i))
    tog.pop(i)
    i=i+13

i=12

while (i < len(disposals)):
    disposals.pop(i)
    tog.pop(i)
    i=i+12

b=0
c=0



disposals_numpy = np.array(disposals)

disp = disposals_numpy.reshape(len(players), games)

tog_numpy = np.array(tog)

time = tog_numpy.reshape(len(players), games)


"""
for i in range(0, games):
    for a in range(0, len(players)):
        if int(time[i][a]) < 60:
"""       

sum = 0

workbook = xl.Workbook()
worksheet = workbook.worksheets[0]

worksheet.cell(1, 1).value = "Players"
worksheet.column_dimensions['A'].width = 15
for i in range(2, games+2):
    worksheet.cell(1, i).value = "Round " + str(i-1)
worksheet.cell(1, i+1).value = "Average"
worksheet.cell(1, i+2).value = "StDev"
worksheet.cell(1, i+3).value = "20+ Percentage"
worksheet.column_dimensions['P'].width = 15
worksheet.cell(1, i+4).value = "20+ Req Odds"
worksheet.column_dimensions['Q'].width = 15
worksheet.cell(1, i+5).value = "20+ Odds"
worksheet.cell(1, i+6).value = "20+ Difference"
worksheet.column_dimensions['S'].width = 15

worksheet.cell(1, i+7).value = "30+ Percentage"
worksheet.column_dimensions['T'].width = 15
worksheet.cell(1, i+8).value = "30+ Req Odds"
worksheet.column_dimensions['U'].width = 15
worksheet.cell(1, i+9).value = "30+ Odds"
worksheet.cell(1, i+10).value = "30+ Difference"
worksheet.column_dimensions['W'].width = 15

for a in range(0, len(players)):
    sum = 0
    count = 0
    values = []
    worksheet.cell(a+2, 1).value = players[a]
    for t in range(0, games):
        worksheet.cell(a+2, t+2).value = int(disp[a][t])
        if int(disp[a][t]) != 0:
            sum = sum + int(disp[a][t])
            count = count + 1
            values.append(int(disp[a][t]))
    avg = sum/count
    worksheet.cell(a+2, t+3).value = avg
    values_numpy = np.array(values)
    worksheet.cell(a+2, t+4).value = np.std(values_numpy)
    worksheet.cell(a+2, t+5).value = (1-norm.cdf(19, avg, np.std(values_numpy)))*100
    worksheet.cell(a+2, t+6).value = 1/(worksheet.cell(a+2, t+5).value/100)

    worksheet.cell(a+2, t+9).value = np.std(values_numpy)
    worksheet.cell(a+2, t+10).value = (1-norm.cdf(29, avg, np.std(values_numpy)))*100
    worksheet.cell(a+2, t+11).value = 1/(worksheet.cell(a+2, t+9).value/100)
workbook.save("Disposals Tracking.xlsx")
