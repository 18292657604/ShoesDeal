# -*- coding:utf-8 -*-
from tkinter import *
import tkinter.messagebox
import tkinter.ttk
from tkinter import filedialog
from shoes import startConvert
from post_office import post


root = Tk()
root.title('Excel格式转换工具')

root.geometry('700x200')
root.resizable(width=True, height=True)


# 0行0列
regular = Label(root, text="请选择规格:", font=("宋体", 10), width=20, height=2)
regular.grid(row=0, column=0, sticky=W, pady=6)

# 0行1列
# 创建下拉菜单
sel = tkinter.ttk.Combobox(root)
# 设置下拉菜单中的值
sel['value'] = ('请选择', '5', '6', '10', '20')
# 设置默认值，即默认下拉框中的内容
sel.current(0)
sel.grid(row=0, column=1, sticky=W, pady=8)

# 文件夹路径变量
mkdir_path = StringVar()
path = Label(text="请输入文件夹路径:", font=("宋体", 10), width=20, height=2)
path.grid(row=1, column=0, sticky=W)
text = Entry(show=None, width=60, textvariable=mkdir_path)
text.grid(row=1, column=1, pady=8)

# 浏览
def mkdir_scan():
    dire_path = filedialog.askdirectory()
    mkdir_path.set(dire_path)

scan_mkdir_btn =  Button(root, text="浏览", command = mkdir_scan, width = 8, height = 1, padx=10)
scan_mkdir_btn.grid(row=1, column=2)

# 类型
var = StringVar()
type = [('订货单', 0), ('订货地址', 1)]
var.set(0)

# 是否隐藏导出路径
def type_cmd():
    # 订货单
    if int(var.get())==0:
        to_path.grid_forget()
        to_text.grid_forget()
        scan_file_btn.grid_forget()
    else:
        to_path.grid(row=3, column=0, sticky=W)
        to_text.grid(row=3, column=1)
        scan_file_btn.grid(row=3, column=2)

for lan,num in type:
    Radiobutton(root, text=lan, value=num, variable=var, command=type_cmd).grid(row=2,column=num, columnspan=2)

# 浏览
file_path = StringVar()
to_path = Label(text="请输入导出文件路径:", font=("宋体", 10), width=20, height=2)
to_path.grid_forget()
# to_path.grid(row=3, column=0, sticky=W)
to_text = Entry(show=None, width=60, textvariable=file_path)
to_text.grid_forget()
# to_text.grid(row=3, column=1)

def file_scan():
    file_name = filedialog.askopenfilename()
    file_path.set(file_name)


scan_file_btn =  Button(root, text="浏览", command = file_scan, width = 8, height = 1, padx=10)
scan_file_btn.grid_forget()


def show():
    # 判断是否为空
    if sel.get()=='请选择':
        tkinter.messagebox.showinfo('温馨提示', '请选择规格！')
        return
    elif text.get()=='':
        tkinter.messagebox.showinfo('温馨提示', '请选择文件夹路径！')
        return

    if int(var.get())==0:
        startConvert(text.get(), int(sel.get()))
    elif int(var.get())==1:
        if to_text.get() == '':
            tkinter.messagebox.showinfo('温馨提示', '请选择导出文件路径！')
            return
        post(text.get(), to_text.get(), int(sel.get()))
    tkinter.messagebox.showinfo('温馨提示', '转换成功')

def clear():
    sel.current(0)
    text.delete(0, END)
    to_text.delete(0, END)

button =  Button(root, text="开始转换", command=show, width = 8, height = 1)
button.grid(row=4, column=0, rowspan=2, columnspan=2, pady=8)

clearBtn = Button(root, text="清除", command=clear, width = 8, height = 1)
clearBtn.grid(row=4, column=1, rowspan=2, columnspan=2, pady=8)
root.mainloop()


