import openpyxl as xl
import requests
from bs4 import BeautifulSoup

webpage_response = requests.get('https://afltables.com/afl/stats/teams/adelaide/2022_gbg.html')
webpage = webpage_response.content
soup = BeautifulSoup(webpage, "html.parser")

table = soup.find_all(attrs={'class':'sortable'})

messy_disposals = table[0].find_all('td')
messy_tog = table[21].find_all('td')

disposals = []
tog = []

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

for a in range(0, len(disposals), 14):
    players.append(disposals[a])
    
print(players)
