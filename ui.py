# -*- coding:utf-8 -*-
from tkinter import *
import tkinter.messagebox
from shoes import startConvert


root = Tk()

root.title('Excel格式转换工具')

root.geometry('700x200')
root.resizable(width=True, height=True)

frame1 = Frame(root)

path = Label(frame1, text="请输入路径:", font=("宋体", 12), width=20, height=2)
path.pack(side = LEFT)

text = Entry(frame1, show=None, width=60)

text.pack(side = LEFT)

frame1.pack()

frame2 = Frame(root)


regular = Label(frame2, text="请输入规格（10/20）:", font=("宋体", 12), width=20, height=2)
regular.pack(side = LEFT)
regular_content = Entry(frame2, show=None, width=10)

regular_content.pack(side = LEFT)
frame2.pack()

def show():
    startConvert(text.get(), regular_content.get())
    tkinter.messagebox.showinfo('温馨提示', '转换成功')

button =  Button(root, text="开始转换", command = show, width = 8, height = 1)
button.pack()

root.mainloop()


