import pandas as pd
import os
import time
from tkinter import *
import tkinter.filedialog

# 格式化成2016-03-20 11:45:39形式
time = time.strftime("%y%m%d-%H%M%S", time.localtime())

def run(filedir):
    older_path, file_name = os.path.split(filedir)


    data = pd.read_excel(filedir)  # 打开原始工作表
    rows = data.shape[0]  # 获取行数 shape[0]获取行数

    txtName = older_path + "\\phonenumbers.txt"
    f = open(txtName, mode='w')
    addres = str(data.iloc[0, 1])
    for i in range(1, rows):

        name = addres + '-' + str(data.iloc[i, 0])
        phonenum = str(data.iloc[i, 1])
        if len(phonenum) == 11:
            new_context = "BEGIN:VCARD\n" \
                          + "VERSION:3.0\n" \
                          + "N;CHARSET=gb2312:" + str(name) + '\n' \
                          + "FN;CHARSET=gb2312:" + str(name) + "\n" \
                          + "TEL;TYPE=CELL:" + str(phonenum) + '\n' \
                          + "END:VCARD\n"
            f.write(new_context)
    f.close()
    os.rename(older_path + '\\phonenumbers.txt', older_path + '\\' + time + '.vcf')


def xz():
    filename=tkinter.filedialog.askopenfilename()
    if filename[-4:] == '.xls' or filename[-5:] == '.xlsx':

        lb.config(text='通讯录文件已保存到原目录')
        run(filename)
    else:
        lb.config(text='您没有选择任何文件')

root = Tk()
root.title('py生成通讯录')
root.geometry('240x240') # 这里的乘号不是 * ，而是小写英文字母 x
lb = Label(root,text='')
lb.pack()
btn=Button(root,text='选择文件',command=xz)
btn.pack()
root.mainloop()

