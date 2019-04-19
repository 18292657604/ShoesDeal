# -*- coding:utf-8 -*-
from tkinter import *



root = Tk()

root.title('Excel格式转换工具')

text = Text(root, width=30, height=5)
text.pack()

def show():
    print('成功了')

# b1 = Button(text, '开始转换', command=show)
#
# text.window_create(INSERT, window=b1)


root.mainloop()


