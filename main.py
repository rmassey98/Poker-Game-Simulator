import poker

""" 

Please enter number of players, number of games to play and where to save the excel file
Note that an excel file is limited to 1,048,576 rows
Each games will populate (number of players + 1) rows
Therefore the maximum games for each number of players is as follows

Two players - 349,525
Three players - 262,144
Four players - 209,715
Five players - 174,762
Six players - 149,796
Seven players - 131,072
Eight Players - 116,508

"""

number_of_players = 6
number_of_games = 20
file_path = r"C:\Dev\Python\Poker\Poker Simulations\PokerGames.xlsx"

if __name__ == "__main__":
    Poker = poker.PlayPoker(number_of_players, number_of_games, file_path)
    Poker.play()

