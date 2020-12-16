from random import randint 
 
t =["Rock", "Paper", "Scissors"] 
 
auto = t[randint(0, 2)] 
 
player = False 
 
while player == False: 
    player = input("type q for quite \nRock, Paper, Scissors? => ") 
    if player == auto: 
        print("Its Tie") 
    elif player == "Rock": 
        if auto == "Paper": 
            print("You loose ", auto," covers ",player) 
        else: 
            print("You Win! ",player, " smashes ",auto) 
    elif player == "Paper": 
        if auto == "Scissors": 
            print("You loose! ",auto ," cut ",player) 
        else: 
            print("You Win! ",player," cover ", auto) 
    elif player == "Scissors": 
        if auto == "Rock": 
            print("You loose! ",auto," smashes ",player) 
        else: 
            print("You Win! ",player," cut ",auto) 
    elif player == "q": 
        break 
    else: 
        print("Thats not a valide entry. Check the spelling ") 
    player = False 
    auto = t[randint(0, 2)] 
