import pyodbc
import os
from tkinter import messagebox as mes
class Connconfig():
    def __init__(self,server,database,uid,pwd):
        self.server=server
        self.database=database
        self.uid=uid
        self.pwd=pwd

    def getcursor(self):
        # global conn,cur
        try:
            self.conn=pyodbc.connect(r'DRIVER={SQL Server};SERVER=%s;DATABASE=%s;UID=%s;PWD=%s' % (self.server,self.database,self.uid,self.pwd))
            self.cur=self.conn.cursor()
            return self.cur
        except pyodbc.OperationalError as err:
            print('aerror')
            mes.showerror('错误提示','找不到服务器：%s' % str(err))
        except pyodbc.InterfaceError as err:
            print('berror')
            mes.showerror('错误提示','数据库连接失败：%s' % str(err))

    def closeconn(self):
        self.cur.close()
        self.conn.close()
