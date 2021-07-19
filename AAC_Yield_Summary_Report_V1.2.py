from tkinter import *
from tkinter import filedialog
import tkinter.messagebox
import os
from datetime import datetime
import txt_report_1V1
from txt_print import txt_print


def run_template():
    default_dir = r"文件路径"
    global filepath_template
    filepath_template = filedialog.askopenfilename(title=u'选择模板', initialdir=(os.path.expanduser(default_dir)))
    text_template.delete(0, END)
    text_template.insert(END, filepath_template)


def run_yield_file():
    default_dir = r"文件路径"
    global filepath_yield_file
    filepath_yield_file = filedialog.askdirectory(title=u'选择TXT路径', initialdir=(os.path.expanduser(default_dir)))
    v.set(filepath_yield_file)


def run():
    report_name = txt_report_1V1.txt_report(filepath_template, filepath_yield_file, text_info)

    print_info = f'Finished -> All files !'
    txt_print(text_info, 'tag2', print_info, 50, 'Green', 'Times', 10)

    print_info = f'Finished -> New Report on {report_name} !'
    txt_print(text_info, 'tag2', print_info, 90, 'Green', 'Times', 10)


# 主窗口
root = Tk()
root.title('AAC_Yield_Summary_Report_V1.0')
root.iconbitmap(r'D:\Python\ico\q7.ico')
root.resizable(0, 0)
root.geometry('700x500')

y_start = 0.01
y_interval = 0.07

# 模板路径
text_template = Entry(root, font=('Helvetica', 10))
text_template.place(relx=0.01, rely=y_start, relwidth=0.86, relheight=0.06)
btn_template = Button(root, text='Template', command=run_template)
btn_template.place(relx=0.88, rely=y_start, relwidth=0.11, relheight=0.06)

# Yield_File 路径
v = StringVar()
text_yield_file = Label(root, justify="left", font=('Helvetica', 10), relief=GROOVE, textvariable=v)
text_yield_file.place(relx=0.01, rely=y_start + y_interval, relwidth=0.86, relheight=0.06)
btn_yield_file = Button(root, text='TXT Path', command=run_yield_file)
btn_yield_file.place(relx=0.88, rely=y_start + y_interval, relwidth=0.11, relheight=0.06)

# 信息输出框 & 执行按钮
text_info = Text(root, font=('Times', 10))
text_info.place(relx=0.01, rely=y_start + 2 * y_interval, relwidth=0.86, relheight=0.84)

# 信息输出框的滚动条
scroll = Scrollbar()
# 放到窗口的右侧, 填充Y竖直方向
scroll.place(relx=0.8492, rely=y_start + 2 * y_interval + 0.001, relwidth=0.02, relheight=0.84 - 0.002)
scroll.config(command=text_info.yview)
text_info.config(yscrollcommand=scroll.set)
btn_run = Button(root, text='START', command=run)
btn_run.place(relx=0.88, rely=y_start + 2 * y_interval, relwidth=0.11, relheight=0.06)

dt = datetime.now()
now_date = dt.strftime('%Y%m%d')
overdue_date = '20230704'

if now_date > overdue_date:
    tkinter.messagebox.askokcancel(title='Error', message='License Expired !')
    exit()

root.mainloop()  # 进入消息循环
