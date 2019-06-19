# -*- coding: utf-8 -*-
import os
from tkinter import *
from tkinter import filedialog, messagebox, ttk


# from matplotlib import rcParams
# rcParams['font.sans-serif'] = ['SimHei']

class MyGUI:
    def __init__(self):
        self.cur_pos = None
        self.need_rename_list = []
        self.init_window = Tk()

        self.pos_label = Label(self.init_window, text="当前位置", font="bold, 8")
        self.cur_pos_label = Label(self.init_window, font="bold, 8")

        self.open_files_button = Button(self.init_window, text="选择需要命名的文件")
        self.open_xls_button = Button(self.init_window, text="选择文件名存放文档")
        self.confirm_button = Button(self.init_window, text="确认更改")

        self.set_init_window()
        self.init_window.mainloop()

    def set_init_window(self):
        self.init_window.title("一键批量修改文件名")
        self.init_window.geometry("500x265+100+100")

        self.open_files_button.place(relx=0.1, rely=0.1, relwidth=0.3, relheight=0.4)
        self.open_xls_button.place(relx=0.6, rely=0.1, relwidth=0.3, relheight=0.4)
        self.confirm_button.place(relx=0.1, rely=0.6, relwidth=0.3, relheight=0.4)

    def open_files(self):
        file_names = filedialog.askopenfilenames(title="请选中需要更改文件名的文件",
                                                 filetypes=[("All Files", '*')])
        self.need_rename_list = file_names
        print(file_names)



MyGUI()
