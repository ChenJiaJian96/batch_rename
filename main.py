# -*- coding: utf-8 -*-
import os
from tkinter import *
from tkinter import filedialog, messagebox, ttk, scrolledtext
from time import localtime, time
from xlrd import open_workbook, XLRDError

ico_path = ".\CSPGCL.ico"


# 温馨提示
# 1.请勿在文件名中添加'.'

# 打包exe文件
# pyinstaller -F -w main.py

class MyGUI:
    def __init__(self):
        self.cur_pos = None
        self.cur_path = ""  # 当前路径
        self.table_name_list = []  # 表格中现有的文件路径列表（含后缀名）
        # display list
        self.table_name_list0 = []  # 表格中现有的旧文件名列表
        self.table_name_list1 = []  # 表格中现有的新文件名列表
        self.table_ext_list = []  # 表格中现有的拓展名列表
        self.disable_list = []  # 记录被删除的数据（适应表格删除行后索引不会更新的特点）
        self.name_reflect_dict = dict()  # 新旧文件名映射（从模板中获取）
        self.template_path = ""  # 新文件名文档的路径
        self.template_data = None  # 新文件名文档数据
        self.init_window = Tk()

        self.pos_label = Label(self.init_window, text="当前位置")
        self.edit_tips_label = Label(self.init_window, text="TIPS:双击原文件名删除该列\n 双击新文件名修改")
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
        self.clear_button = Button(self.init_window, text="清空表格", command=self.reset_tree)

        self.question_label = Label(self.init_window, text=" ? ", font="bold, 8")
        self.exclamation_label = Label(self.init_window, text=" ! ", font="bold, 8")

        self.set_init_window()
        self.init_window.mainloop()

    def reset_tree(self):
        for item in self.tree_view.get_children():
            self.tree_view.delete(item)
            rn = int(str(item).replace('I', ''))
            self.disable_list.append(rn - 1)

    def set_init_window(self):
        self.init_window.title("一键批量修改文件名")
        self.init_window.geometry("560x300+100+100")
        self.init_window.iconbitmap(ico_path)

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

        self.open_files_button.place(relx=0.8, rely=0.2, relwidth=0.15, relheight=0.15)
        self.open_xls_button.place(relx=0.8, rely=0.4, relwidth=0.15, relheight=0.15)
        self.confirm_button.place(relx=0.8, rely=0.6, relwidth=0.15, relheight=0.15)
        self.clear_button.place(relx=0.8, rely=0.8, relwidth=0.15, relheight=0.15)
        self.edit_tips_label.place(relx=0.05, rely=0.83, relwidth=0.3)

        # 生成右侧提示按钮
        self.question_label.bind("<Button-1>", self.show_instruction)
        self.question_label.place(relx=0.96, rely=0.8, relwidth=0.03, relheight=0.08)
        self.exclamation_label.bind("<Button-1>", func=self.show_software_detail)
        self.exclamation_label.place(relx=0.96, rely=0.9, relwidth=0.03, relheight=0.08)

    # 显示软件详情
    @staticmethod
    def show_software_detail(event):
        messagebox.showinfo("关于", "ISBN:\n著作权人:\n出版单位:")

    # 显示操作说明
    def show_instruction(self, event):
        instruction_dialog = InstructionDialog()
        self.init_window.wait_window(instruction_dialog.rootWindow)

    @staticmethod
    def get_greetings(hour):
        if 6 <= hour <= 11:
            return "早上好"
        elif 11 <= hour <= 13:
            return "中午好"
        elif 13 <= hour <= 18:
            return "下午好"
        else:
            return "晚上好"

    def open_files(self):
        file_paths = filedialog.askopenfilenames(title="请选中需要更改文件名的文件",
                                                 filetypes=[("All Files", '*')])
        if len(file_paths) > 0:
            if not self.is_same_location(file_paths[0]):
                messagebox.showwarning("文件路径冲突", "文件路径发生冲突，单次修改请在同一路径下操作")
                return

            self.set_file_location(file_paths[0])
            for path in file_paths:
                if not self.is_file_added(path):
                    self.add_file_to_table(path)

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
            return False
        else:
            return True

    def set_file_location(self, file_path):
        (file_path, temp_filename) = os.path.split(file_path)
        self.cur_path = file_path + '/'
        self.cur_pos_label.configure(text=self.cur_path)

    def is_same_location(self, file_path):
        (file_path, temp_filename) = os.path.split(file_path)
        if self.cur_path == file_path + '/' or self.cur_path == "":
            return True
        else:
            return False

    def is_file_added(self, file_path):
        (file_path, file_name) = os.path.split(file_path)
        temp_name_list = []
        for i in range(len(self.table_name_list)):
            if i not in self.disable_list:
                temp_name_list.append(self.table_name_list[i])
        if file_name in temp_name_list:
            print("have added")
            return True
        else:
            print("haven't added")
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
            # 首先检查新文件名的合法性
            for new_name in temp_dict.values():
                if not is_name_legal(new_name):
                    messagebox.showwarning("新文件名格式有误", "文档中含非法文件名：" + new_name + "\n合法文件名不应存在\\ / \" * : ? < > | 等字符")
                    return
            self.name_reflect_dict = temp_dict
            self.update_table_by_dict()

    # 使用文件名映射更新表格（全局）
    def update_table_by_dict(self):
        for i in range(len(self.table_name_list0)):
            if self.name_reflect_dict.keys().__contains__(self.table_name_list0[i]):
                if i not in self.disable_list:
                    temp_name = self.name_reflect_dict[self.table_name_list0[i]]
                    self.table_name_list1.append(temp_name)
            else:
                self.table_name_list1.append("")
        # 更新新文件名列表后插入表格
        print(self.table_name_list1)
        row_str = self.tree_view.get_children()
        for cur_row in row_str:
            rn = int(str(cur_row).replace('I', ''))
            print(cur_row)
            self.tree_view.set(cur_row, column='#2', value=self.table_name_list1[rn - 1])
        self.tree_view.update()

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
        if cn == 1:
            def del_tree_column():
                self.tree_view.delete(item)
                self.disable_list.append(rn - 1)

            del_tree_column()
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
        if len(self.table_name_list0) - len(self.disable_list) == 0:
            print("Nothing needed to be changed.")
            return
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
            if not len(self.table_name_list1[i]) == 0 and i not in self.disable_list:
                name = self.table_name_list[i]
                (name, ext) = os.path.splitext(name)
                old_name = self.table_name_list0[i] + ext
                new_name = self.table_name_list1[i] + ext
                print("old_name: " + old_name)
                print("new_name: " + new_name)
                try:
                    os.rename(old_name, new_name)
                except FileNotFoundError:
                    messagebox.showwarning("无法找到文件", "系统找不到指定的文件" + old_name)
                    return

        messagebox.showinfo("文件重命名成功", "所有文件名修改成功")


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


# 显示使用流程弹窗
class InstructionDialog:
    def __init__(self):
        self.rootWindow = Toplevel()
        self.rootWindow.title('使用流程和常见问题')
        self.rootWindow.geometry("500x400+250+250")
        self.rootWindow.iconbitmap(ico_path)

        self.guide_button = Button(self.rootWindow, text="使用流程", command=lambda: self.update_text(1))
        self.quest_button = Button(self.rootWindow, text="常见问题", command=lambda: self.update_text(2))
        self.wel_text = MyGUI.get_greetings(localtime(time()).tm_hour) + ",欢迎查阅使用流程及常见问题\n\n请点击上面按钮进行查询↑↑↑"
        self.guide_text = "使用说明\n\n" \
                          "一、使用流程\n" \
                          "选中需要重命名的文件->打开模板文件->查看重命名表格->确定重命名\n\n" \
                          "二、选中需要重命名的文件\n" \
                          "2.1--本系统可以直接重命名各种类型的文件。请在打开文件窗中选中需要重命名的文件，" \
                          "未选中文件则无法使用接下来的功能；\n" \
                          "2.2--请保证需要重命名的文件处于同一文件夹，且文件名不含【.】\n" \
                          "（由于模板文件由Excel文件储存，小数点容易导致数据错误）\n\n" \
                          "三、打开模板文件\n" \
                          "3.1--模板文件由表格文件组成，如.xlsx/.xls/.et等文件。请在打开文件窗中选中对应的模板文件，" \
                          "若文件过大需要耐心等待一段时间，模板文件由两列构成，分别为【旧文件名】及【新文件名】；\n" \
                          "3.2--文件完整性检查，为了使本系统的功能均能正常使用，打开文件后会进行完整性检查，主要检查文件中是否" \
                          "存在以下列，包括‘旧文件名’、‘新文件名’等，若数据不完整，将无法进行进一步操作；\n" \
                          "3.3--文件名合法性检查，按照Windows标准合法文件名中不能包含【\\/:*?\"<>|】，若包含非法字符则直接导入失败\n\n" \
                          "四、查看重命名表格\n" \
                          "4.1--选择需要重命名的文件及打开模板文件后，新旧文件名会以倒序方式排列在主界面表格中，旧文件会对应重命名" \
                          "为新文件名；\n" \
                          "4.2--删除表格列，用户双击表格中对应旧文件名可以删除表格列，取消对应文件的重命名请求；\n" \
                          "4.3--修改新文件名，用户可以双击表格中对应新文件名进行自定义修改，双击后请在弹出文本框中编辑，并点击确认。" \
                          "若输入文件名非法，则无法修改。\n\n"\
                          "五、确认重命名\n" \
                          "用户在确认表格信息无误后可直接点击【确认更改】按钮执行操作。\n\n" \
                          "六、主界面快捷按钮说明\n" \
                          "6.1--【？】按钮：显示使用流程及常见问题弹窗；\n" \
                          "6.2--【！】按钮：显示软件版权信息。"
        self.quest_text = "常见问题说明\n\n" \
                          "Q1.为什么有些按钮无法点击？\n" \
                          "A1.在所需步骤未执行时软件会禁用部分按钮，请按照正确使用流程操作。\n\n" \
                          "Q2.请打开正确格式的文件\n" \
                          "A2.该软件仅能打开.xlsx/.xls/.et等表格文件，请检查打开文件格式。\n\n" \
                          "Q3.文件路径冲突\n" \
                          "A3.软件默认每次只能修改同一路径（文件夹）下的文件。\n\n" \
                          "Q4.文件名格式有误\n" \
                          "A4.请检查新文件名中是否包含非法字符。\n\n" \
                          "Q5.无法找到文件\n" \
                          "A5.请检查主界面显示的当前路径下是否包含对应文件。\n\n" \
                          "其他问题请重启软件，无法解决请联系开发者。"
        self.content_text = scrolledtext.ScrolledText(self.rootWindow, wrap=WORD)
        self.box_scrollbar_y = Scrollbar(self.rootWindow)

        self.guide_button.place(relx=0.27, rely=0.03, relwidth=0.2, relheight=0.1)
        self.quest_button.place(relx=0.53, rely=0.03, relwidth=0.2, relheight=0.1)
        self.content_text.place(relx=0.02, rely=0.16, relwidth=0.96, relheight=0.81)
        self.update_text(0)

    def update_text(self, update_type):
        self.content_text.delete(1.0, END)
        if update_type == 1:
            self.content_text.insert(INSERT, self.guide_text)
        elif update_type == 2:
            self.content_text.insert(INSERT, self.quest_text)
        else:
            self.content_text.insert(INSERT, self.wel_text)


MyGUI()
