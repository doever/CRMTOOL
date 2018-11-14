import tkinter as tk
from tkinter import ttk

from setting import *


class BaseView:
    # def __init__(self, menu_name: dict):
    def __init__(self, tab_control):
        self.tab = ttk.Frame(tab_control)


class MainView:
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
        tab = ttk.Frame(tab_control)
        tab_control.add(tab, text=MENU_NAME['page_two'])
        l_title = tk.Label(tab, text="浩  泽", fg='blue', font=("Symbol", "15"))
        l_title.pack()
        tk.Label(tab, text='----------------------------------------------------------------------', font=('', 10)).pack()
        l_warning = tk.Label(tab, text="注:请输入XX000000-0000格式的单号", fg='DeepPink', font=("Comic Sans MS", "8"))
        l_warning.pack()
        tk.Label(tab, text='-----------开户退换单据撤单请注意同时作废浩优以及WMS单据--------------', fg='DeepPink', font=("Comic Sans MS", "8")).pack()
        t_workarea = tk.Text(tab, height=20, width=80)
        t_workarea.pack()
        t_showarea = tk.Text(tab, height=15, width=80)
        t_showarea.pack()
        # Text颜色实现
        '''
        #第一个参数为自定义标签的名字
        #第二个参数为设置的起始位置，第三个参数为结束位置
        #第四个参数为另一个位置
        showbill.tag_add('tag1','1.0','end')
        #用tag_config函数来设置标签的属性
        showbill.tag_config('tag1',background='LightCyan',foreground='red')
        '''
        b_cancel = tk.Button(tab, text='浩泽撤单', activebackground='blue', activeforeground='Black', bg='PaleTurquoise', fg='black', command=hz_channelorder)
        b_cancel.pack(side=tk.LEFT)
        b_cancel.bind()
        b_finish = tk.Button(tab, text='浩泽结单', activebackground='blue', activeforeground='Black', bg='PaleTurquoise', fg='black', command=hz_finishorder)
        b_finish.pack(side=tk.RIGHT)


class HYCancelView():
    def init_ui(self, tab_control):
        tab = ttk.Frame(tab_control)
        tab_control.add(tab, text=MENU_NAME['page_three'])
        l_title = tk.Label(tab, text="灏  优", fg='blue', font=("Symbol", "15"))
        l_title.pack()
        l_warning = tk.Label(tab, text="注:浩优单据需选择模式,模式一对应尾数4位数单据,模式二对应尾数5位数单据,请勿混用", fg='DeepPink', font=("Comic Sans MS", "8"))
        l_warning.pack()
        model = tk.StringVar()
        tk.Radiobutton(tab, text='模式1:GD000000-0000 ', variable=model, value='a', command=model_select, font=('', 10)).pack()
        tk.Radiobutton(tab, text='模式2:GD000000-00000', variable=model, value='b', command=model_select, font=('', 10)).pack()
        t_workarea = tk.Text(tab, height=20, width=80)
        t_workarea.pack()
        t_showarea = tk.Text(tab, height=15, width=80)
        t_showarea.pack()
        b_cancel = tk.Button(tab, text='浩优撤单', activebackground='yellow', activeforeground='Black', bg='BlanchedAlmond', fg='black',
                       command=hy_channelorder)
        b_cancel.pack(side=tk.LEFT)
        b_finish = tk.Button(tab, text='浩优结单', activebackground='yellow', activeforeground='Black', bg='BlanchedAlmond', fg='black',
                       command=hy_finishorder)
        b_finish.pack(side=tk.RIGHT)


class EamilTaskView():
    def init_ui(self, tab_control):
        tab = ttk.Frame(tab_control)
        canvas_email = tk.Canvas(tab, height=585, width=580)
        canvas_email.pack()
        # canvas_email.create_rectangle(0,0,580,585,fill="white")
        email_background = tk.PhotoImage(file="image/email1.png")
        canvas_email.create_image(0, 200, anchor=tk.NW, image=email_background)
        canvas_email.create_rectangle(0, 0, 580, 200, fill="#FFF")
        canvas_email.create_text(290, 35, text="ETL脚本文件", fill='black', font=('Comic Sans MS', 25))
        canvas_email.create_rectangle(0, 0, 6, 200, fill="skyblue")
        canvas_email.create_rectangle(0, 0, 580, 6, fill="skyblue")
        canvas_email.create_rectangle(576, 0, 580, 200, fill="skyblue")
        canvas_email.create_rectangle(0, 196, 580, 200, fill="skyblue")
        b_hz_report = tk.Button(tab, text='浩泽报表', width=12, fg='#999', bg='Gold', activebackground='#00F',
                          command=lambda: os.system("call ETLbat/OznerReport.bat"))
        b_hz_report.place(x=100, y=100)
        b_financial = tk.Button(tab, text='财务数据', width=12, fg='#999', bg='Gold', activebackground='#00F',
                          command=lambda: os.system("call ETLbat/FinanceData.bat"))
        b_financial.place(x=250, y=100)
        b_water_chip = tk.Button(tab, text='水芯片', width=12, fg='#999', bg='Gold', activebackground='#00F', command=lambda: os.system("call ETLbat/SXP.bat"))
        b_water_chip.place(x=400, y=100)
        b_five_day = tk.Button(tab, text='每月五号', width=12, fg='#999', bg='Gold', activebackground='#00F',
                          command=lambda: os.system("call ETLbat/FiveDay.bat"))
        b_five_day.place(x=100, y=150)
        b_clear = tk.Button(tab, text='清理数据', width=12, fg='#999', bg='Gold', activebackground='#00F',
                          command=lambda: os.system("call ETLbat/clean.bat"))
        b_clear.place(x=250, y=150)
        b_reboot = tk.Button(tab, text='重启电脑', width=12, fg='#999', bg='Gold', activebackground='#00F',
                          command=lambda: os.system("call ETLbat/Restart.bat"))
        b_reboot.place(x=400, y=150)
        # bat_2=tk.Button(tab3,text='',width=8,fg='red',activebackground='green',command=lambda: foo(x))


class ReplaceView():
    def init_ui(self, tab_control):
        tab = ttk.Frame(tab_control)
        l_replace = tk.Label(tab, text="单据替换", fg='blue', font=("Symbol", "15"))
        l_replace.pack()
        l_remark = tk.Label(tab, text="注：模式一用于XX111111-1111格式", font=("Times", "8", "bold italic"), fg="red")
        l_remark.pack()
        l_remark_t = tk.Label(tab, text="注：模式二用于XX111111-11111格式", font=("Times", "8", "bold italic"), fg="red")
        l_remark_t.pack()
        # 模式选择
        replace_type = tk.StringVar()
        tk.Radiobutton(tab, text='模式一', variable=replace_type, value='a', command=model_select_2, font=('', 8)).pack(side=tk.TOP)
        tk.Radiobutton(tab, text='模式二', variable=replace_type, value='b', command=model_select_2, font=('', 8)).pack(side=tk.TOP)
        t_in_area = tk.Text(tab, height=15, width=80)
        t_in_area.pack()
        t_out_area = tk.Text(tab, height=15, width=80)
        t_out_area.pack()
        b_replace = tk.Button(tab, text='替换', width=8, fg='red', activebackground='green', command=changetext)
        b_replace.pack()


class DumpDataView():
    pass


class AbnormalOrderView():
    pass


class HelpView():
    pass
