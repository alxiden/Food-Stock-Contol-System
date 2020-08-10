import Setup as s
import Stock as t
from tkinter import *
import os

a = os.listdir()
        
root = Tk()

root.title("SWM Launcher")

welcome = Label(root, text = "Welcome to the Stock and Wastage Manager Launcher")
welcome.grid(row = 0, column =1, columnspan =3)

l1 = Label (root, text = "Please choose from the options below")
l1.grid(row = 1, column =1, columnspan =3)
    
submit = Button(root, text = "Exsisting Account", height = "1", width = "15", command = lambda: t.startup())
submit.grid(row = 3, column =1)

submit = Button(root, text = "New Account", height = "1", width = "15", command = lambda: s.setep())
submit.grid(row = 3, column =2)

submit = Button(root, text = "Master Account", height = "1", width = "15", command = lambda: t.startup())
submit.grid(row = 5, column =1)

submit = Button(root, text = "Something probably", height = "1", width = "15", command = lambda: t.startup())
submit.grid(row = 5, column =2)
