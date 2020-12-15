from tkinter import * 

import random 

 

expression ="" 

expression1 = "" 

 

 

def press(num): 

     

       

    global expression 

    expression = num 

    equation1.set(expression) 

     

    global expression1 

    expression1 = random.randint(1, 5) 

    equation2.set(expression1) 

     

    if expression == expression1: 

        msg1 = Message(gui, text="Congratulation. You won the game!!! Keep playing!!!", fg="blue").place(x=130, y=300) 

             

    else: 

        msg2 = Message(gui, text="Sorry you lost. Better luck next time!!Keep playing", fg="red").place(x=130, y=300) 

         

     

     

 

 

if __name__ == "__main__": 

    gui = Tk() 

    gui.configure(background = "black") 

    gui.title("Guess The number") 

    gui.geometry("360x400") 

 

    button1 = Button(gui, text='1', fg = "white", bg='gray', command=lambda: press(1) , height = 2, width = 9) 

    button1.grid(row = 7, column = 2) 

    button2 = Button(gui, text='2', fg = "white", bg='gray', command=lambda: press(2) , height = 2, width = 9) 

    button2.grid(row = 5, column = 3) 

    button3 = Button(gui, text='3', fg = "white", bg='gray', command=lambda: press(3) , height = 2, width = 9) 

    button3.grid(row = 4, column = 4) 

    button4 = Button(gui, text='4', fg = "white", bg='gray', command=lambda: press(4) , height = 2, width = 9) 

    button4.grid(row = 5, column = 5) 

    button5 = Button(gui, text='5', fg = "white", bg='gray', command=lambda: press(5) , height = 2, width = 9) 

    button5.grid(row = 7, column = 6) 

 

    label1 = Label(gui, text="Your number", fg ="yellow", bg="red").place(x=60, y=190) 

    equation1 = StringVar() 

    equation1_field = Entry(gui, textvariable=equation1) 

    equation1_field.place(x=200, y=190) 

 

    lable2 = Label(gui, text="Random number", fg="yellow", bg="red").place(x=60, y=230) 

    equation2 = StringVar() 

    equation2_field = Entry(gui, textvariable=equation2) 

    equation2_field.place(x=200, y=230) 

     

         

         

    gui.mainloop()
