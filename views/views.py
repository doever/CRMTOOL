import tkinter as tk
from tkinter import ttk

from setting import *


class BaseView:
    def __init__(self, menu_name: dict):
        pass


class MainView():
    def init_ui(self, title: str):
        top = tk.Tk()
        top.title(title)
        tab_control = ttk.Notebook(top)
        tab_control.pack(expand=5, fill="both")
        # tabControl.add(,text='首页')
        IndexView().init_ui(tab_control)


# 各个页面
class IndexView:
    def init_ui(self, tab_control):
        tab = ttk.Frame(tab_control)
        tab_control.add(tab, text=MENU_NAME['page_one'])
        canvas = tk.Canvas(tab, height=PAGE_HEIGHT, width=PAGE_WIDTH)
        canvas.pack()
        canvas.create_rectangle(0, 0, PAGE_WIDTH, PAGE_HEIGHT, fill="white")
        canvas.create_text(290, 80, text=INDEX_CONFIG['logo'], fill='skyblue', font=('Comic Sans MS', FONT_SIZE['title_lg']))
        canvas.create_text(290, 130, text=INDEX_CONFIG['title'], fill='skyblue', font=('Arial', FONT_SIZE['title_sm']))
        # 导入背景图片
        try:
            img = tk.PhotoImage(file=INDEX_CONFIG['background'])
        except:
            img = tk.PhotoImage(file=INDEX_CONFIG['default_background'])
        canvas.create_image(0, 165, anchor=tk.NW, image=img)
        # canvas.create_rectangle(0,290,202,294,fill='white',outline='Khaki')
        # canvas.create_rectangle(198,294,202,580,fill='white',outline='Khaki')


class HZCancelView:
    def init_ui(self, tab_control):
        tab1 = ttk.Frame(tab_control)
        tab_control.add(tab1, text=MENU_NAME['page_two'])
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
