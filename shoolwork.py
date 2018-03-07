import os
import xlrd
import tkinter as tk
from tkinter import *
from tkinter import messagebox
import tkinter.filedialog
import traceback

def change_names(work_path,excel_path):
    try:
        data = xlrd.open_workbook(excel_path)
        table = data.sheet_by_name('reportofstudent')
        #生成学生学号与身份证号对应字典
        first_row = table.row_values(0)
        c_student_number = first_row.index('学号')
        c_id_number = first_row.index('身份证号')
        list_student_number = table.col_values(c_student_number)
        list_id_number = table.col_values(c_id_number)
        dict_sn_idn = {}
        for c in range(1,len(list_id_number)-1):
            dict_sn_idn[list_student_number[c]] = list_id_number[c]
    except :
        print("发生excel读取错误！")
        messagebox.showerror("错误","发生excel读取错误！错误内容为%s。" % (traceback.format_exc()))
    #修改所有图片的名称
    path = work_path
    list_error = []
    for f in os.listdir(path):
        if os.path.isfile(os.path.join(path,f))==True:
            if f.find('.jpg') > 0:
                if dict_sn_idn.get(f[:f.find('.jpg')],False):
                    try:
                        os.rename(os.path.join(path,f),os.path.join(path,dict_sn_idn[f[:f.find('.jpg')]]+".jpg"))
                    except:
                        list_error.append(f)
                else:
                    list_error.append(f)
    print('共有'+str(len(list_error))+'个学号未能找到对应身份证号。\n')
    messagebox.showinfo("成功","改名已完成，有%s个学号未能找到对应身份证号或身份证信息有误。"%(str(len(list_error))))

def on_click_btn_excel_path():
    filename = tkinter.filedialog.askopenfilename()
    if filename != '':
        text_excel_path.delete('1.0',END)
        text_excel_path.insert(INSERT,filename)

def on_click_btn_main_path():
    directoryname = tkinter.filedialog.askdirectory()
    if directoryname != '':
        text_main_path.delete('1.0',END)
        text_main_path.insert(INSERT,directoryname)

def on_click_btn_change():
    excel_path = text_excel_path.get('1.0',END)
    excel_path = excel_path[:len(excel_path)-1]
    main_path = text_main_path.get('1.0',END)
    main_path = main_path[:len(main_path)-1]
    if os.path.isfile(excel_path) and os.path.isdir(main_path):
        if excel_path.find(".xls") > 0 or excel_path.find(".XLS") > 0:
            print(1)
            change_names(main_path,excel_path)
    else:
        messagebox.showerror("错误","路径或文件错误！")

window = tk.Tk()
window.title('修改学号为身份证')
window.geometry('500x150')

lable_excel_path = tk.Label(window,text="excel路径：")
lable_excel_path.place(x=10,y=20)
text_excel_path = tk.Text(window,width=50,height=1)
text_excel_path.place(x=80,y=25)
btn_excel_path = tk.Button(window,text="选择",height=1,command=on_click_btn_excel_path)
btn_excel_path.place(x=450,y=18)

lable_main_path = tk.Label(window,text="图片路径：")
lable_main_path.place(x=10,y=60)
text_main_path = tk.Text(window,width=50,height=1)
text_main_path.place(x=80,y=65)
btn_main_path = tk.Button(window,text="选择",height=1,command=on_click_btn_main_path)
btn_main_path.place(x=450,y=58)

btn_change = tk.Button(window,text='开始',height=1,command=on_click_btn_change)
btn_change.place(x=250,y=100)

window.mainloop()