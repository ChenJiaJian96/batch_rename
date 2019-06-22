# -*- coding: utf-8 -*-
import os
from tkinter import *
from tkinter import filedialog, messagebox, ttk

from xlrd import open_workbook, XLRDError


# 温馨提示
# 1.请勿在文件名中添加'.'


class MyGUI:
    def __init__(self):
        self.cur_pos = None
        self.cur_path = ""  # 当前路径
        self.table_name_list = []  # 表格中现有的文件路径列表（含后缀名）
        # display list
        self.table_name_list0 = []  # 表格中现有的旧文件名列表
        self.table_name_list1 = []  # 表格中现有的新文件名列表
        self.table_ext_list = []  # 表格中现有的拓展名列表
        self.name_reflect_dict = dict()  # 新旧文件名映射（从模板中获取）
        self.template_path = ""  # 新文件名文档的路径
        self.template_data = None  # 新文件名文档数据
        self.init_window = Tk()

        self.pos_label = Label(self.init_window, text="当前位置")
        self.edit_tips_label = Label(self.init_window, text="TIPS:双击可修改新文件名")
        self.cur_pos_label = Label(self.init_window, justify=LEFT)
        # 表格
        columns = ("1", "2", "3")
        self.tree_view = ttk.Treeview(self.init_window, show="headings", columns=columns)  # 表格
        self.vbar = ttk.Scrollbar(self.init_window, orient=VERTICAL, command=self.tree_view.yview)
        # 定义树形结构与滚动条
        self.tree_view.configure(yscrollcommand=self.vbar.set)

        self.open_files_button = Button(self.init_window, text="选择文件", command=self.open_files)
        self.open_xls_button = Button(self.init_window, text="选择模板", command=self.open_template)
        self.confirm_button = Button(self.init_window, text="确认更改", command=self.check_new_name)

        self.set_init_window()
        self.init_window.mainloop()

    def set_init_window(self):
        self.init_window.title("一键批量修改文件名")
        self.init_window.geometry("530x275+100+100")

        self.pos_label.place(relx=0.05, rely=0.05, relheight=0.1)
        self.cur_pos_label.place(relx=0.2, rely=0.05, relheight=0.1)

        self.tree_view.place(relx=0.05, rely=0.2, relwidth=0.7, relheight=0.6)
        self.tree_view.bind('<Double-1>', self.set_cell_value)  # 双击左键进入编辑
        self.vbar.place(relx=0.75, rely=0.2, relheight=0.6)
        self.tree_view.column("1", width=150, anchor='center')  # 表示列,不显示
        self.tree_view.column("2", width=150, anchor='center')
        self.tree_view.column("3", width=50, anchor='center')

        self.tree_view.heading("1", text="原文件名")  # 显示表头
        self.tree_view.heading("2", text="新文件名")
        self.tree_view.heading("3", text="类型")

        self.open_files_button.place(relx=0.8, rely=0.2, relwidth=0.15, relheight=0.18)
        self.open_xls_button.place(relx=0.8, rely=0.49, relwidth=0.15, relheight=0.18)
        self.confirm_button.place(relx=0.8, rely=0.77, relwidth=0.15, relheight=0.18)

        self.edit_tips_label.place(relx=0.05, rely=0.85, relwidth=0.3)

        self.set_button_state(0)

    def open_files(self):
        file_paths = filedialog.askopenfilenames(title="请选中需要更改文件名的文件",
                                                 filetypes=[("All Files", '*')])
        if len(file_paths) > 0:
            if not self.is_same_location(file_paths[0]):
                messagebox.showwarning("文件路径冲突", "文件路径发生冲突，单次修改请在同一路径下操作")
                self.clear_tree()

            self.set_file_location(file_paths[0])
            for path in file_paths:
                if not self.is_file_added(path):
                    self.add_file_to_table(path)
                    print("选中的文件" + path)

    def open_template(self):
        template_path = filedialog.askopenfilename(title="请选中新文件名的模板文件",
                                                   filetypes=[("表格文件", '*.xls; *.xlsx; *.et')])
        print(template_path)
        try:
            temp_data = open_workbook(template_path)
        except FileNotFoundError:
            pass
        except XLRDError:
            messagebox.showwarning("请打开正确格式的模板文件")
        else:
            self.template_path = template_path
            self.template_data = ExcelMaster(temp_data)
            if self.check_file_integrity():
                self.init_name_reflect_dict()

    def check_file_integrity(self):
        flag = 0
        if self.template_data.col_index('旧文件名') == -1:
            flag = 1
        if self.template_data.col_index('新文件名') == -1:
            flag = 1
        if flag:
            messagebox.showwarning("模板文件格式有误", "模板文件格式有误，请检查后重新导入")
            self.set_button_state(0)
            return False
        else:
            return True

    def set_file_location(self, file_path):
        (file_path, temp_filename) = os.path.split(file_path)
        self.cur_path = file_path + '/'
        self.cur_pos_label.configure(text=self.cur_path)

    def is_same_location(self, file_path):
        (file_path, temp_filename) = os.path.split(file_path)
        if self.cur_path == file_path or self.cur_path == "":
            return True
        else:
            return False

    def is_file_added(self, file_path):
        (file_path, file_name) = os.path.split(file_path)
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

    # 初始化名字映射dict
    def init_name_reflect_dict(self):
        temp_dict = self.template_data.get_name_dict()
        if temp_dict is None:
            messagebox.showwarning("模板文件有误", "模板文件为空或格式损坏，请检查文件内容")
        else:
            self.name_reflect_dict = temp_dict
            print(self.name_reflect_dict)
            self.update_table_by_dict()

    # 使用文件名映射更新表格（全局）
    def update_table_by_dict(self):
        for i in range(len(self.table_name_list0)):
            if self.name_reflect_dict.keys().__contains__(self.table_name_list0[i]):
                temp_name = self.name_reflect_dict[self.table_name_list0[i]]
                if not is_name_legal(temp_name):
                    messagebox.showwarning("新文件名格式有误", "新文件名有误：" + temp_name + "\n合法文件名不应存在\\ / \" * : ? < > | 等字符")
                    self.table_name_list1.clear()
                    return
                else:
                    self.table_name_list1.append(temp_name)
            else:
                self.table_name_list1.append("")
        # 更新新文件名列表后插入表格
        row_str = self.tree_view.get_children()
        for cur_row in row_str:
            rn = int(str(cur_row).replace('I', ''))
            print(cur_row)
            self.tree_view.set(cur_row, column='#2', value=self.table_name_list1[rn - 1])
        self.set_button_state(1)

    # 使按钮失效，无法使用
    def set_button_state(self, i):
        if i == 0:
            self.confirm_button.config(state=DISABLED)
        elif i == 1:
            self.confirm_button.config(state=ACTIVE)

    def clear_tree(self):
        x = self.tree_view.children
        for item in x:
            self.tree_view.delete(item)
        self.table_name_list.clear()
        self.table_name_list0.clear()
        self.table_name_list1.clear()
        self.table_ext_list.clear()

    def set_cell_value(self, event):  # 双击进入编辑状态
        for item in self.tree_view.selection():
            print("set_cell_value" + item)
            item_text = self.tree_view.item(item, "values")
            print(item_text[0:3])  # 输出所选行的值

        column = self.tree_view.identify_column(event.x)  # 列
        row = self.tree_view.identify_row(event.y)
        cn = int(str(column).replace('#', ''))
        rn = int(str(row).replace('I', ''))
        print(rn)
        if cn == 2:
            entry_edit = Text(self.init_window, width=20, height=1)
            entry_edit.place(relx=0.35, rely=0.85, relwidth=0.3, relheight=0.08)

            def save_edit():
                new_name = entry_edit.get(1.0, "end")
                print("new name:" + new_name)
                if not is_name_legal(new_name):
                    messagebox.showwarning("新文件名格式有误", "新文件名有误：" + new_name + "\n合法文件名不应存在\\ / \" * : ? < > | 等字符")
                    entry_edit.delete(0.0, "end")
                else:
                    try:
                        self.tree_view.set(item, column=column, value=new_name)
                        self.table_name_list1[rn - 1] = new_name  # 刷新新文件名的列表
                    except NameError:
                        messagebox.showwarning("文件未添加", "请先添加文件后修改")
                    entry_edit.destroy()
                    ok_btn.destroy()

            ok_btn = ttk.Button(self.init_window, text='OK', command=save_edit)
            ok_btn.place(relx=0.65, rely=0.85, relwidth=0.1, relheight=0.08)

    def check_new_name(self):
        print("def check_new_name")
        for new_name in self.table_name_list1:
            if len(new_name) == 0:
                a = messagebox.askokcancel('空文件名提示', '个别文件无新文件名，是否继续更名？')
                if a:
                    self.confirm_rename()
                else:
                    return
        self.confirm_rename()

    def confirm_rename(self):
        print("def confirm_rename")
        print(self.cur_path)
        os.chdir(self.cur_path)
        for i in range(len(self.table_name_list0)):
            if not len(self.table_name_list1[i]) == 0:
                name = self.table_name_list[i]
                (name, ext) = os.path.splitext(name)
                old_name = self.table_name_list0[i] + ext
                new_name = self.table_name_list1[i] + ext
                print("old_name" + old_name)
                print("new_name" + new_name)
                try:
                    os.rename(old_name, new_name)
                except FileNotFoundError:
                    messagebox.showwarning("无法找到文件", "系统找不到指定的文件" + old_name)
                    return

        messagebox.showinfo("文件重命名成功", "所有文件名修改成功")
        # todo: 清空表格
        # todo: 删除添加的文件


# 数据类
class ExcelMaster:
    def __init__(self, data):
        self.data = data  # 源文件
        self.table = None  # 保存当前正在处理的表格
        # 初始化表格
        self.set_table(0)
        # 获取表格总行数
        self.nrow = self.table.nrows

    # index:第index个sheet,入参需要检查
    def set_table(self, index=0):
        if self.data is None:
            return "文件为空，无法打开工作表！"
        else:
            self.table = self.data.sheet_by_index(index)

    # 返回新旧文件名的映射
    def get_name_dict(self):
        i = self.col_index('旧文件名')
        name0_list = self.table.col_values(i, start_rowx=1, end_rowx=None)
        j = self.col_index('新文件名')
        name1_list = self.table.col_values(j, start_rowx=1, end_rowx=None)
        if len(name0_list) != len(name1_list) or len(name0_list) == 0:
            return None
        else:
            type0_list = self.table.col_types(i, start_rowx=1, end_rowx=None)
            type1_list = self.table.col_types(j, start_rowx=1, end_rowx=None)
            # 净化识别为数字的数据
            for i in range(len(type0_list)):
                if type0_list[i] == 2:
                    name0_list[i] = str(name0_list[i]).split(".")[0]
            for i in range(len(type1_list)):
                if type1_list[i] == 2:
                    name1_list[i] = str(name1_list[i]).split(".")[0]
            name_dict = dict()
            for i in range(len(name0_list)):
                name_dict[str(name0_list[i])] = str(name1_list[i])
            return name_dict

    # 返回列名返回列索引
    def col_index(self, col_name):
        first_col_list = self.table.row_values(0)  # 第一行元素生成列表
        try:
            i = first_col_list.index(col_name)
        except ValueError:
            return -1
        else:
            return i


def is_name_legal(text):
    if len(text) == 0:
        return True
    # 正则表达式判断特殊字符
    if not re.search(u'[\\\\/:*?\"<>|]', text):
        return True
    else:
        return False


MyGUI()
