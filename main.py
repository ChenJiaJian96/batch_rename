# -*- coding: utf-8 -*-
import os
from tkinter import *
from tkinter import filedialog, messagebox, ttk


class MyGUI:
    def __init__(self):
        self.cur_pos = None
        self.need_rename_list = []
        self.init_window = Tk()

        self.pos_label = Label(self.init_window, text="当前位置")
        self.cur_pos_label = Label(self.init_window, text="2132", justify=LEFT)
        # 表格
        columns = ("1", "2", "3")
        self.tree_view = ttk.Treeview(self.init_window, show="headings", columns=columns)  # 表格

        self.open_files_button = Button(self.init_window, text="选择文件", command=open_files)
        self.open_xls_button = Button(self.init_window, text="选择模板")
        self.confirm_button = Button(self.init_window, text="确认更改")

        self.set_init_window()
        self.init_window.mainloop()

    def set_init_window(self):
        self.init_window.title("一键批量修改文件名")
        self.init_window.geometry("530x265+100+100")

        self.pos_label.place(relx=0.05, rely=0.05, relheight=0.1)
        self.cur_pos_label.place(relx=0.2, rely=0.05, relheight=0.1)

        self.tree_view.place(relx=0.05, rely=0.2, relwidth=0.7, relheight=0.7)
        self.tree_view.column("1", width=150, anchor='center')  # 表示列,不显示
        self.tree_view.column("2", width=150, anchor='center')
        self.tree_view.column("3", width=50, anchor='center')

        self.tree_view.heading("1", text="原文件名")  # 显示表头
        self.tree_view.heading("2", text="新文件名")
        self.tree_view.heading("3", text="类型")

        self.open_files_button.place(relx=0.8, rely=0.2, relwidth=0.15, relheight=0.15)
        self.open_xls_button.place(relx=0.8, rely=0.45, relwidth=0.15, relheight=0.15)
        self.confirm_button.place(relx=0.8, rely=0.7, relwidth=0.15, relheight=0.15)

    def open_files(self):
        file_names = filedialog.askopenfilenames(title="请选中需要更改文件名的文件",
                                                 filetypes=[("All Files", '*')])
        self.need_rename_list = file_names
        print(file_names)


MyGUI()
