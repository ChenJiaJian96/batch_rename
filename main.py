# -*- coding: utf-8 -*-
import os
from tkinter import *
from tkinter import filedialog, messagebox, ttk


class MyGUI:
    def __init__(self):
        self.cur_pos = None
        self.table_name_list = []   # 表格中现有的文件路径列表
        self.table_name_list0 = []  # 表格中现有的旧文件名列表
        self.table_name_list1 = []  # 表格中现有的新文件名列表
        self.table_ext_list = []  # 表格中现有的拓展名列表
        self.rename_path = ""  # 新文件名文档的路径
        self.init_window = Tk()

        self.pos_label = Label(self.init_window, text="当前位置")
        self.cur_pos_label = Label(self.init_window, justify=LEFT)
        # 表格
        columns = ("1", "2", "3")
        self.tree_view = ttk.Treeview(self.init_window, show="headings", columns=columns)  # 表格

        self.open_files_button = Button(self.init_window, text="选择文件", command=self.open_files)
        self.open_xls_button = Button(self.init_window, text="选择模板", command=self.open_template)
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
        file_paths = filedialog.askopenfilenames(title="请选中需要更改文件名的文件",
                                                 filetypes=[("All Files", '*')])
        if len(file_paths) > 0:
            self.set_file_location(file_paths[0])
            for path in file_paths:
                if not self.is_file_added(path):
                    self.add_file_to_table(path)
                    print("选中的文件" + path)

    def open_template(self):
        file_paths = filedialog.askopenfilename(title="请选中新文件名的模板文件",
                                                filetypes=[("表格文件", '*.xls; *.xlsx; *.et')])
        self.rename_path = file_paths
        print(file_paths)

    def set_file_location(self, file_path):
        (file_path, temp_filename) = os.path.split(file_path)
        self.cur_pos_label.configure(text=file_path)

    def is_file_added(self, file_paths):
        for path in file_paths:
            (file_path, file_name) = os.path.split(path)
            if file_name in self.table_name_list:
                return True
            else:
                self.table_name_list.append(file_name)
                return False

    def add_file_to_table(self, path):
        (path, name) = os.path.split(path)
        (name, ext) = os.path.splitext(name)

        self.table_name_list0.append(name)
        self.table_ext_list.append(ext)
        self.tree_view.insert('', len(name) - 1, values=(
            self.table_name_list0[len(self.table_name_list0) - 1], '',
            self.table_ext_list[len(self.table_ext_list) - 1]))
        self.tree_view.update()


MyGUI()
