import tkinter as tk
from tkinter import ttk

from setting import *


class BaseView:
    # 创建分离控制器
    def create_control(self):
        pass


# 各个页面
class IndexView:
    def init_ui(self, top):
        # 创建分离控制器
        tabControl = ttk.Notebook(top)
        tabControl.pack(expand=5, fill="both")
        # 创建视图零
        tab0 = ttk.Frame(tabControl)
        tabControl.add(tab0, text='首页')
        canvas = tk.Canvas(tab0, height=585, width=580)
        canvas.pack()
        canvas.create_rectangle(0, 0, 580, 585, fill="white")
        canvas.create_text(290, 80, text="OZNER", fill='skyblue', font=('Comic Sans MS', 25))
        canvas.create_text(290, 130, text="浩泽CRM工具", fill='skyblue', font=('Arial', 18))
        # 导入图片
        img = tk.PhotoImage(file="image/backpic1.png")
        canvas.create_image(0, 165, anchor=tk.NW, image=img)
        # canvas.create_rectangle(0,290,202,294,fill='white',outline='Khaki')
        # canvas.create_rectangle(198,294,202,580,fill='white',outline='Khaki')


class HZCancelView():
    def init_ui(self, top):
        tab1 = ttk.Frame(tabControl)
        tabControl.add(tab1, text='浩泽撤单')
        la_hz = tk.Label(tab1, text="浩  泽", fg='blue', font=("Symbol", "15"))
        la_hz.pack()
        tk.Label(tab1, text='----------------------------------------------------------------------', font=('', 10)).pack()
        la1_hz = tk.Label(tab1, text="注:请输入XX000000-0000格式的单号", fg='DeepPink', font=("Comic Sans MS", "8"))
        la1_hz.pack()
        tk.Label(tab1, text='-----------开户退换单据撤单请注意同时作废浩优以及WMS单据--------------', fg='DeepPink', font=("Comic Sans MS", "8")).pack()
        t_hz = tk.Text(tab1, height=20, width=80)
        t_hz.pack()
        showbill_hz = tk.Text(tab1, height=15, width=80)
        showbill_hz.pack()
        # Text颜色实现
        '''
        #第一个参数为自定义标签的名字
        #第二个参数为设置的起始位置，第三个参数为结束位置
        #第四个参数为另一个位置
        showbill.tag_add('tag1','1.0','end')
        #用tag_config函数来设置标签的属性
        showbill.tag_config('tag1',background='LightCyan',foreground='red')
        '''
        b1 = tk.Button(tab1, text='浩泽撤单', activebackground='blue', activeforeground='Black', bg='PaleTurquoise', fg='black', command=hz_channelorder)
        b1.pack(side=tk.LEFT)
        b = tk.Button(tab1, text='浩泽结单', activebackground='blue', activeforeground='Black', bg='PaleTurquoise', fg='black', command=hz_finishorder)
        b.pack(side=tk.RIGHT)


class HYCancelView():
    pass


class EamilTaskView():
    pass


class ReplaceView():
    pass


class DumpDataView():
    pass


class AbnormalOrderView():
    pass


class HelpView():
    pass
