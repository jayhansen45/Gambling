import itertools

total_bet = 200
pointsbet = 750
spend = 700


brisbane112 = 6
geelong112 = 7
geelong1324 = 6
geelong2536 = 6.5
geelong37 = 7


brisbane112_bet = []
geelong112_bet = []
geelong1324_bet = []
geelong2536_bet = []
geelong37_bet = []
combos = []

for j in range(26, 501, 25):
    for i in range(j-26, j):
        brisbane112_bet.append(i)
        geelong112_bet.append(i)
        geelong1324_bet.append(i)
        geelong2536_bet.append(i)
        geelong37_bet.append(i)

    bets = [brisbane112_bet, geelong112_bet, geelong1324_bet, geelong2536_bet, geelong37_bet]
    temp = list(itertools.product(*bets))

    print(temp)

    brisbane112_bet = []
    geelong112_bet = []
    geelong1324_bet = []
    geelong2536_bet = []
    geelong37_bet = []   


    
    wins = 0

    for m in range(0, len(temp)):
        for h in range(0, 5):
            total_bet = total_bet + temp[m][h]
        """if (temp[m][0]*brisbane112>total_bet) and (temp[m][1]*geelong112>total_bet) and (temp[m][2]*geelong1324>total_bet) and (temp[m][3]*geelong2536>total_bet) and (temp[m][4]*geelong37>total_bet):
            print(temp[m])
            print(total_bet)
            print()"""
        total_bet = 200
