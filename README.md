# 文件名批量修改系统--batch_rename
开发时间：2019/06/19 -- 2019/08/28

## 使用说明

一、使用流程
选择需要重命名的文件->导入模板文件->查看主界面窗口->确认更改

二、选择需要重命名的文件
2.1--本系统可以直接重命名各种类型的文件。点击“选择文件”添加文件，在打开文件窗中选择添加需要重命名的文件；
2.2--请保证需要重命名的文件处于同一文件夹，且文件名不含小数点'.'。

三、导入模板文件
3.1--完成文件选择后系统会自动在文件所处的文件夹中创建默认模板文件[template.xls]，并将文件名导入模板文件；
3.2--模板文件为.xlsx/.xls/.et等表格文件。请先打开默认模板文件进行查看与更改，模板文件由两列构成，分别为【旧文件名】及【新文件名】，若文件过大需要耐心等待一段时间；
3.3--点击主界面“打开模板”按钮导入模板；
3.4--为保证本系统的功能正常使用，请首先确保‘旧文件名’列、‘新文件名’列等数据完整，其次确保文件名中不包含【\\/:*?\"<>|】。

四、查看主界面窗口
4.1--完成以上步骤后，新旧文件名会以倒序方式排列在主界面表格中并依次对应；
4.2--删除误选文件，用户双击主界面表格中的“旧文件名”可以删除该行，取消对应文件的重命名；
4.3--修改新文件名，用户可以双击主界面表格中的“新文件名”进行自定义修改，双击后请在弹出文本框中编辑，并点击确认。若输入文件名非法，则修改无法完成。

五、确认更改
用户在确认主界面表格信息无误后可直接点击【确认更改】按钮执行操作。

六、主界面快捷按钮说明
6.1--【？】按钮：显示使用流程及常见问题弹窗；
6.2--【！】按钮：显示软件版权信息。


## 常见问题说明

Q1.请关闭文件夹下的模板文件后再重新导入
A1.由于该系统会自动导入用户选中的旧文件名，请在选择需求重命名的文件时，确保模板文件处于关闭状态，否则会出现读写冲突。

Q2.请打开正确格式的模板文件
A2.该软件仅能打开.xlsx/.xls/.et等表格文件，请检查打开文件格式。

Q3.文件路径冲突
A3.软件默认每次只能修改同一路径下的文件，请确保需重命名文件均在同一文件夹内。

Q4.文件名格式有误
A4.请检查涉及文件名中是否包含非法字符，如【\\/:*?\"<>|】。

Q5.无法找到文件
A5.请检查主界面显示的当前路径下是否包含对应文件。

其他问题请重启软件，无法解决请及时联系开发者550728110@qq.com。

