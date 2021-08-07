# -*- coding=utf-8 -*-
# -----------------------------
# @Time      : 2020/12/8 15:31
# @Author    : wangyd5
# @File      : demo.py
# @Project   : 定制
# @Function  ： #-----------------------------
import tkinter.filedialog
from tkinter import *
import tkinter as tk
from PIL import Image,ImageTk
from tkinter import ttk, messagebox, filedialog
import os
import pandas as pd
import win32com
from win32com.client import Dispatch
import sys
import os.path

class Root(tk.Tk):
    def __init__(self):
        # 设置参数
        self.path = os.path.abspath(".")  # 能够在下面self.path 引用到
        print(self.path)
        super(Root, self).__init__()
        self.title('MJ')
        self.attributes("-alpha", 0.97)
        self.geometry("660x570") # 宽*高
        # self.minsize(600, 500)
        # self.maxsize(800, 1000)



        # menu
        menu = Menu(self)
        self.config(menu=menu)
        self.selct_path = Menu(menu, tearoff=0)
        self.selct_path.add_command(
            label="打开", accelerator="Ctrl + O", command=self.open_dir)
        menu.add_cascade(label="文件", menu=self.selct_path)

        about = Menu(menu, tearoff=0)
        about.add_command(label="版本", accelerator="v1.0.0")
        about.add_command(label="作者", accelerator="boy friend")
        menu.add_cascade(label="关于", menu=about)

        # 显示选中路径
        # 顶部frame
        self.top_var = StringVar()
        self.top_frame = Frame(self, bg="#fff")
        self.top_frame.pack(side=TOP, fill=X)
        self.label = Label(self.top_frame, textvariable=self.top_var, bg="#fff")
        self.top_var.set('当前选中路径：%s' % self.path)
        self.label.pack(side=LEFT)
        #
        self.frame = ttk.Frame(self,height=500, width=200, relief=RIDGE, borderwidth=2)
        # self.frame.pack(side=TOP,fill=Y,expand=True)
        self.canvas = Canvas(self.frame,background="#D2D2D2")

        self.scrollbar = ttk.Scrollbar(self.frame,orient='vertical',command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas, width=400, height=700)
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        # 显示文件列表
        self.frame1 = ttk.Frame(self.scrollable_frame,height=100,width=300, relief=RIDGE, borderwidth=2)
        # self.canvas.create_window(0,0, window=self.frame1, anchor='nw')
        self.frame1.pack(side=TOP,fill=X) #fill=Y, ipady=2,
        self.btn1 = Button(self.frame1, text="显示文件列表", width=10, command=self.load_file)
        # self.btn1.bind("<MouseWheel>", self._on_mousewheel)
        # 将组件添加到窗口中显示，并指定好位置
        # side上下左右，anchor东南西北，fill填充到...
        # btn1.pack(side=tkinter.LEFT, anchor=tkinter.N, fill=tkinter.Y)
        # btn1.pack(side=tkinter.RIGHT)
        self.btn1.pack(padx=30, pady=80, side=LEFT, anchor=N)
        # 文件下拉列表
        self.lb = Listbox(self.frame1, width=20, height=10)
        # self.lb.bind("<MouseWheel>", self._on_mousewheel)
        self.lb.pack( padx=10, pady=10,fill=X, side=TOP)


        # # 输入文件关键字
        self.frame2 = tkinter.Frame(self.scrollable_frame, height=10, width=100, relief=RIDGE, bg='#fff', bd=2, borderwidth=2)
        # self.canvas.create_window(10,60, window=self.frame2, anchor='nw')
        self.frame2.pack(side=TOP,expand=True, fill=X,ipady=2)
        # 显示过滤文件
        self.btn2 = Button(self.frame2, text='过滤文件', width=10, command=self.filter_file)
        # self.btn2.bind("<MouseWheel>", self._on_mousewheel)
        self.btn2.pack(fill=X, padx=110, pady=4, side=LEFT)

        self.keword = StringVar()
        self.en_word = Entry(self.frame2, textvariable=self.keword, width=36, bg='#D3D3D3')
        self.keword.set('请输入文件关键字,默认为  xlsx ')
        # self.en_word.bind("<MouseWheel>", self._on_mousewheel)
        self.en_word.pack(padx=16, pady=20, side=LEFT)

        # excel
        self.frame_excel = tkinter.Frame(self.scrollable_frame, height=40, width=60, relief=RIDGE, bg='#fff', bd=2, borderwidth=2)
        self.frame_excel.pack(fill=X, ipady=5)


        # excel 图片

        if hasattr(sys, "_MEIPASS"):
            excel_dir = os.path.join(sys._MEIPASS, 'excel.gif')
            ppt_dir = os.path.join(sys._MEIPASS, 'pptx.gif')
        else:
            excel_dir = 'excel.gif'
            ppt_dir = 'pptx.gif'
        self.frame_image_excel = tkinter.Frame(self.frame_excel, height=30, width=60, relief="flat", bg='#fff', bd=2, borderwidth=2)
        photo = Image.open(excel_dir)
        photo = photo.resize((30, 30))
        photo = ImageTk.PhotoImage(photo)
        excel_label = Label(self.frame_image_excel, image=photo)
        excel_label.image = photo
        excel_label.pack()
        self.frame_image_excel.grid(column=0, rowspan=2, columnspan=2, sticky=W + E + N + S, padx=5, pady=5)

        # 输入行数
        self.frame3 = tkinter.Frame(self.frame_excel, height=30, width=60, relief='groove', bg='#fff', bd=2, borderwidth=0.6)
        # self.canvas.create_window(10,60,window=self.frame3, anchor='nw')
        # self.frame3.pack(fill=X, ipady=10)
        self.frame3.grid(row=1,column=2,padx=10,pady=15)
        self.envar = StringVar()
        self.en_row = Entry(self.frame3, textvariable=self.envar, width=23, bg='#D3D3D3')
        self.envar.set('请输入表头行数，默认为 1')
        # self.en_row.bind("<MouseWheel>", self._on_mousewheel)
        self.en_row.pack(padx=10, pady=10, side=TOP)

        # 显示过滤文件
        self.btn3 = Button(self.frame3, text='合并excel文件', width=13, command=self.merge_excel_file)
        # self.btn3.bind("<MouseWheel>", self._on_mousewheel)
        self.btn3.pack(fill=X, padx=10, pady=10, side=TOP)


        # 保存文件
        self.frame4 = tkinter.Frame(self.frame_excel, height=10, width=60, relief='groove', bg='#fff', bd=2, borderwidth=0.6)
        # self.canvas.create_window(10,60, window=self.frame4, anchor='nw')
        # self.frame4.pack(side=TOP,fill=X,expand=True, ipady=10)
        self.frame4.grid(row=1,column=3,padx=10,pady=15)
        self.varname = StringVar()
        self.en_name = Entry(self.frame4, textvariable=self.varname, width=40, bg='#D3D3D3')
        self.varname.set('请输入合并后的文件名，默认为 merge.xlsx ')
        # self.en_name.bind("<MouseWheel>", self._on_mousewheel)
        self.en_name.pack(padx=10, pady=10, side=TOP)


        self.btn4 = Button(self.frame4, text='保存excel文件', width=40, command=self.save_excel_file)
        # self.btn4.bind("<MouseWheel>", self._on_mousewheel)
        self.btn4.pack(fill=X, padx=30, pady=10, side=TOP)



        # ppt
        # ppt 图片
        self.frame_ppt = tkinter.Frame(self.scrollable_frame, height=40, width=600, relief=RIDGE, bg='#fff', bd=2, borderwidth=2)
        self.frame_ppt.pack(fill=X)

        self.frame_image_ppt = tkinter.Frame(self.frame_ppt, height=60, width=60, relief="flat", bg='#fff', bd=2, borderwidth=2)
        photo = Image.open(ppt_dir)
        photo = photo.resize((30, 30))
        photo = ImageTk.PhotoImage(photo)
        ppt_label = Label(self.frame_image_ppt, image=photo)
        ppt_label.image = photo
        ppt_label.pack()

        self.frame_image_ppt.grid(column=0, rowspan=2, sticky=W + E + N + S, padx=5, pady=5)
        # ppt
        self.frame5 = tkinter.Frame(self.frame_ppt, height=60, width=500, relief='groove', bg='#fff', bd=2, borderwidth=0.6)
        # self.canvas.create_window( 10,60,window=self.frame5, anchor='nw')
        # self.frame5.pack(side=TOP,fill=X,expand=True, ipady=10)
        self.frame5.grid(row=1,column=2,pady=10,padx=40,sticky=W+E+S+W)

        self.pptname = StringVar()
        self.en_ppt_name = Entry(self.frame5, textvariable=self.pptname, width=40, bg='#D3D3D3')
        self.pptname.set('请输入合并后的文件名，默认为 merge.pptx ')
        # self.en_ppt_name.bind("<MouseWheel>", self._on_mousewheel)
        # self.en_ppt_name.pack( side=TOP)
        self.en_ppt_name.grid(row=0,pady=10,column=1,columnspan=2)
        # 显示过滤文件
        self.btn5 = Button(self.frame5, text='合并ppt文件', width=20, command=self.merge_ppt_file)
        # self.btn5.bind("<MouseWheel>", self._on_mousewheel)
        # self.btn5.pack(fill=X, padx=30, pady=10, side=LEFT)
        self.btn5.grid(row=1,column=1,padx=50)

        self.btn6 = Button(self.frame5, text='保存ppt文件', width=20, command=self.save_ppt_file)
        # self.btn6.bind("<MouseWheel>", self._on_mousewheel)
        # self.btn6.pack(fill=X, padx=1, pady=10, side=RIGHT)
        self.btn6.grid(row=1,column=2,padx=30,pady=10)
        #
        #
        #
        self.frame.pack(side=TOP,fill=BOTH,expand=True)
        self.canvas.pack(side="left", fill="both", expand=True)
        # self.scrollable_frame.pack()
        self.scrollbar.pack(side="right", fill="y")

        # self.update()
        # self.canvas.config(scrollregion=self.canvas.bbox("all"))
        # self.frame1.bind("<MouseWheel>", self._on_mousewheel)
        # self.frame2.bind("<MouseWheel>", self._on_mousewheel)
        # self.frame3.bind("<MouseWheel>", self._on_mousewheel)
        # self.frame4.bind("<MouseWheel>", self._on_mousewheel)
        # self.frame5.bind("<MouseWheel>", self._on_mousewheel)
        # self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)



    def _on_mousewheel(self,event):
        self.canvas.yview_scroll(-1*int(event.delta/120), "units")

    def save_excel_file(self):
        ''' 保存文件'''
        varname = self.varname.get() if '默认' not in self.varname.get() else 'merge.xlsx'
        if not os.path.exists(self.path.replace('\\','/')+'/' + 'output'):
            os.mkdir(self.path.replace('\\','/')+'/' + 'output')
        path = self.path.replace('\\','/')+'/'+ 'output/' + varname
        print(path)

        if (varname[-4:] =='xlsx') | (varname[-3:]=='xls'):
            if not os.path.exists(path):
                self.data.to_excel(path,index=False)
                messagebox.showinfo("提示", "保存成功")
            else:
                os.remove(path)
                self.data.to_excel(path, index=False)
                messagebox.showinfo("提示", "保存成功,文件被覆盖")
        else:
            messagebox.showwarning("提示", "不是excel 文件，请重新输入文件名")

    def save_ppt_file(self):
        ''' 保存ppt文件'''
        pptname = self.pptname.get() if '默认' not in self.pptname.get() else 'mj_merge.pptx'
        if not os.path.exists(self.path.replace('\\','/')+'/' + 'output'):
            os.mkdir(self.path.replace('\\','/')+'/' + 'output')
        path = self.path.replace('\\','/')+'/'+ 'output/' + pptname
        # print(path)

        if (pptname[-4:]=='pptx') | (pptname[-3:]=='ppt'):
            if not os.path.exists(path):
                self.new_ppt.SaveAs(path )
                self.new_ppt.Close()
                self.ppt.Quit()
                messagebox.showinfo("提示", "ppt保存成功")
            else:
                os.remove(path)
                self.new_ppt.SaveAs(path + 'mj_ppt_merge.pptx')
                self.new_ppt.Close()
                self.ppt.Quit()
                messagebox.showinfo("提示", "保存成功,文件被覆盖")
        else:
            messagebox.showwarning("提示", "不是pptx 文件，请重新输入文件名")

    def merge_excel_file(self):
        self.data = pd.DataFrame([])
        skip_row = int(self.envar.get()) if '默认' not in self.envar.get() else 1
        path = self.path.replace('\\','/')+'/'
        self.de_file_list.sort()
        result = []
        for i, file_path in enumerate(self.de_file_list):
            print(file_path)

            abspath = path + file_path
            if i == 0:
                # 将所有数据都转化为字符串
                excel = pd.ExcelFile(abspath)
                for sheet in excel.sheet_names:
                    columns = excel.parse(sheet).columns
                    converters = {column: str for column in columns}
                    tmp_df = excel.parse(sheet, converters=converters,header=None)
                    if skip_row == 1:
                        first_part = tmp_df.iloc[:skip_row].copy()
                        # first_part = pd.DataFrame(first_part)
                        print(first_part)

                    else:
                        first_part = tmp_df.iloc[:skip_row].copy()

                    after_part = tmp_df.iloc[skip_row:].copy()
                    result.append(first_part)
                    result.append(after_part)
            else:
                excel = pd.ExcelFile(abspath)
                for sheet in excel.sheet_names:
                    columns = excel.parse(sheet).columns
                    converters = {column: str for column in columns}
                    tmp_df = excel.parse(sheet, converters=converters, header=None,skiprows=skip_row)
                    result.append(tmp_df)
        self.data = pd.concat(result)
        self.data.reset_index(drop=True,inplace=True)
        col_len = len(self.data.columns)
        self.data.columns = list(self.data.iloc[0]) + ['0']* (col_len-len(self.data.iloc[0]))
        self.data.drop(0,inplace=True)

        # print(self.data)
        messagebox.showinfo("提示", "合并成功")

    def merge_ppt_file(self):
        """合并ppt 文件"""
        self.ppt = win32com.client.Dispatch('PowerPoint.Application')
        self.ppt.Visible = 1
        self.ppt.DisplayAlerts = 0  # 不显示，不警告
        self.new_ppt = self.ppt.Presentations.Add()
        direct_path = self.path.replace('\\', '/') + '/'
        for f in self.de_file_list:
            path = direct_path + f
            pptSel = self.ppt.Presentations.Open(path)
            win32com.client.gencache.EnsureDispatch('PowerPoint.Application')
            slide_count = pptSel.Slides.Count
            pptSel.Close()
            self.ppt_num = self.new_ppt.Slides.InsertFromFile(path, self.new_ppt.Slides.Count, 1, slide_count)

        messagebox.showinfo("提示", "ppt合并成功")

    def open_dir(self):
        ''' 设置默认搜索路径 '''
        self.path = filedialog.askdirectory(title=u"设置目录", initialdir=self.path)
        print("设置路径：" + self.path)
        self.top_var.set('当前选中路径：%s'% self.path)

    def load_file(self):
        self.lb.delete(0,END)
        print('load file',self.path)
        self.file_list = os.listdir(self.path)
        for file_name in self.file_list:
            self.lb.insert(END,file_name)

    def filter_file(self):
        self.lb.delete(0,END)
        tmp = self.keword.get()
        if '默认' in tmp:
            tmp ='xlsx'

        self.de_file_list =[x for x in self.file_list if tmp in x ]
        self.de_file_list = [x for x in self.de_file_list if '~$' not in x]
        for file_name in self.de_file_list:
            self.lb.insert(END,file_name)


if __name__ == '__main__':

    root = Root()
    root.resizable(width=True, height=True)
    root.mainloop()

