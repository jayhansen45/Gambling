"""
Work out how to rotate the board
    Find the absolute of the difference and then find the difference
Add barriers
    Use coordinates to do it?
    Add these coordinates to the game board


"""

import string

player_1_row = 17
player_1_col = 9
player_2_row = 1
player_2_col = 9

w, h = 19, 19
board = [[0 for x in range(w)] for y in range(h)] 

for i in range(0, 19):
    if i%2 == 0:
        for j in range(0, 19):
            if j%2 == 0:
                board[i][j] = "."
            else:
                board[i][j] = " "
    else:
        for k in range(0, 19):
            board[i][k] = " "



def print_board():
    board[player_1_row][player_1_col] = "X"
    board[player_2_row][player_2_col] = "0"

    for i in range(0, 19):
        for j in range(0, 19):
            print(board[i][j], end = " ")
        print("\n")

print_board()

move = ""

while move != "Winning":
    move = input("Player 1 Move: ")
    if move == "up":
        board[player_1_row][player_1_col] = " "
        player_1_row = player_1_row  - 2
    elif move == "down":
        board[player_1_row][player_1_col] = " "
        player_1_row = player_1_row + 2
    elif move == "left":
        board[player_1_row][player_1_col] = " "
        player_1_col = player_1_col - 2
    elif move == "right":
        board[player_1_row][player_1_col] = " "
        player_1_col = player_1_col + 2
    elif move.split()[0] == "h":
        coords = move.split()
        row = coords[1].split(",")[0]
        col = coords[1].split(",")[1]
        board[2*int(row)-2][2*int(col)-1] = "_"
    elif move.split()[0] == "v":
        coords = move.split()
        row = coords[1].split(",")[0]
        col = coords[1].split(",")[1]
        board[2*int(row)-1][2*int(col)-2] = "|"
    else:
        print("Try Again")
        
    print_board()
    
    move = input("Player 2 Move: ")
    if move == "up":
        board[player_2_row][player_2_col] = " "
        player_2_row = player_2_row  - 2
    elif move == "down":
        board[player_2_row][player_2_col] = " "
        player_2_row = player_2_row + 2
    elif move == "left":
        board[player_2_row][player_2_col] = " "
        player_2_col = player_2_col - 2
    elif move == "right":
        board[player_2_row][player_2_col] = " "
        player_2_col = player_2_col + 2
    elif move.split()[0] == "h":
        coords = move.split()
        row = coords[1].split(",")[0]
        col = coords[1].split(",")[1]
        board[2*int(row)-2][2*int(col)-1] = "_"
    elif move.split()[0] == "v":
        coords = move.split()
        row = coords[1].split(",")[0]
        col = coords[1].split(",")[1]
        board[2*int(row)-1][2*int(col)-2] = "|"
    else:
        print("Try Again")
    
    print_board()
