import openpyxl as xl
import requests
import numpy as np
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

#for a in range(0, len(disposals), 14):
 #   players.append(disposals.pop(a))

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



for i in range(0, games):
    for a in range(0, len(players)):
        if int(time[i][a]) < 60:
            
