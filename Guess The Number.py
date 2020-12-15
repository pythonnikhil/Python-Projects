import random 
 
wanna_play = "yes" 
 
while wanna_play == "yes" or wanna_play == "y": 
    print("====Game of Guess the number===") 
    n = int(input("Guess the number between 1 to 5 and enter ")) 
    r = random.randint(1, 5) 
    print ("Random number: ",r) 
    print ("Your Guessed = ",n) 
    if n==r: 
        print("Hurray, You won. Here is your Champagne !!!") 
    else: 
        print("Bad luck!!! Better luck next time") 
 
    wanna_play = input("Wanna Play? ")
