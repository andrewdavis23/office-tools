# Generate combos from two lists with duplicates removed. ver.3/28/2022

# from tkinter import *
from tkinter import Tk
from tkinter import Label
from tkinter import Text
from tkinter import Button
from tkinter import Frame
from tkinter import END
from tkinter import INSERT

def compile():
    # output
    combos = ''

    # vertical text box, split on line break, create list, remove duplicates
    l1 = list(set(str(tb_1.get('1.0', END)).split('\n')))
    l2 = list(set(str(tb_2.get('1.0', END)).split('\n')))

    l1.remove('')
    l2.remove('')

    l1.sort(key=str)
    l2.sort(key=str)

    # compile
    for i in l1:
        for j in l2:
            combos += i + '\t' + j + '\n'

    # insert into third text boxs
    tb_3.insert(INSERT, str(combos))

def clear_1():
    tb_1.delete('1.0', END)

def clear_2():
    tb_2.delete('1.0', END)

def clear_3():
    tb_3.delete('1.0', END)

def clear_all():
    clear_1()
    clear_2()
    clear_3()

root = Tk()
root.title("Combinations")
root.geometry('365x477')
root.resizable(False, False)

fbg = '#fffafa'
tbg = '#fafafa'
bbg = '#fafaff'

frame1 = Frame(root, bd=5, bg=fbg)
frame1.pack()

version = Label(frame1, text='Generate combos from two lists with duplicates removed. ver.3/28/2022', font=('Arial',8), bg=fbg)
version.grid(row=0, column=0, columnspan=3)

tb_1 = Text(frame1, width=10, bd=4, bg=tbg)
tb_1.grid(row=1, column=0)

tb_2 = Text(frame1, width=10, bd=4, bg=tbg)
tb_2.grid(row=1, column=1)

tb_3 = Text(frame1, width=20, bd=4, bg=tbg)
tb_3.grid(row=1, column=2)

btn_clr1 = Button(frame1, text="Clear 1", command=clear_1, bg=bbg)
btn_clr2 = Button(frame1, text="Clear 2", command=clear_2, bg=bbg)
btn_clr3 = Button(frame1, text="Clear Combos", command=clear_3, bg=bbg)
btn_clrA = Button(frame1, text="Clear All", command=clear_all, bg=bbg)
btn_compile = Button(frame1, text="Compile", command=compile, bg=bbg)

btn_clr1.grid(row=2, column=0, sticky='nsew', padx=1)
btn_clr2.grid(row=2, column=1, sticky='nsew', padx=1)
btn_clr3.grid(row=2, column=2, sticky='nsew', padx=1)
btn_clrA.grid(row=3, column=0, columnspan=2, sticky='nsew', padx=1)
btn_compile.grid(row=3, column=2, columnspan=2, sticky='nsew', padx=1)

root.mainloop()