# -*- coding:utf-8 -*-
from tkinter import *
import tkinter.messagebox
import tkinter.ttk
from shoes import startConvert


root = Tk()
root.title('Excel格式转换工具')

root.geometry('700x200')
root.resizable(width=True, height=True)


# 0行0列
regular = Label(root, text="请选择规格:", font=("宋体", 10), width=20, height=2)
regular.grid(row=0, column=0, pady=6)

# 0行1列
# 创建下拉菜单
sel = tkinter.ttk.Combobox(root)
# 设置下拉菜单中的值
sel['value'] = ('请选择', '5', '6', '10', '20')
# 设置默认值，即默认下拉框中的内容
sel.current(0)
sel.grid(row=0, column=1, sticky=W, pady=8)


path = Label(text="请输入路径:", font=("宋体", 10), width=20, height=2)
path.grid(row=1, column=0)
text = Entry(show=None, width=60)
text.grid(row=1, column=1)

'''
def scan():
    pass

scanBtn =  Button(root, text="浏览", command = scan, width = 8, height = 1, padx=10)
scanBtn.grid(row=1, column=2)
'''

def show():
    startConvert(text.get(), sel.get())
    tkinter.messagebox.showinfo('温馨提示', '转换成功')

def clear():
    sel.current(0)
    text.delete(0, END)

button =  Button(root, text="开始转换", command=show, width = 8, height = 1)
button.grid(row=2, column=0, rowspan=2, columnspan=2, pady=8)

clearBtn = Button(root, text="清除", command=clear, width = 8, height = 1)
clearBtn.grid(row=2, column=1, rowspan=2, columnspan=2, pady=8)
root.mainloop()


