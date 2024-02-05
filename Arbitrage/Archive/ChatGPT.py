from selenium import webdriver
from selenium.webdriver.chrome.service import Service

# Start a new Chrome browser
service = Service('path/to/chromedriver') # Download chromedriver.exe and provide its path
driver = webdriver.Chrome(service=service)
driver.maximize_window()

# Navigate to the IPL cricket page
driver.get('https://www.sportsbet.com.au/betting/cricket/indian-premier-league')

# Wait for the page to load and for the game odds to become visible
driver.implicitly_wait(10)

# Extract the game odds for each game listed on the page
game_odds = []
games = driver.find_elements_by_css_selector('.sbr-EventMarketRow')
for game in games:
    team_names = game.find_elements_by_css_selector('.sbr-ParticipantStackedInfo_Name')
    odds = game.find_elements_by_css_selector('.sbr-OddsDecimal')
    game_odds.append({
        'team1': team_names[0].text,
        'team2': team_names[1].text,
        'team1_odds': float(odds[0].text),
        'team2_odds': float(odds[1].text),
    })

# Print the extracted game odds
for game in game_odds:
    print(game['team1'], 'vs', game['team2'])
    print('   ', game['team1_odds'], '-', game['team2_odds'])

# Close the browser
driver.quit()
