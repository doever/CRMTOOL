import pyodbc
import os
from tkinter import messagebox as mes
class Connconfig():
    def __init__(self,server,database,uid,pwd):
        self.server=server
        self.database=database
        self.uid=uid
        self.pwd=pwd

    def getcursor():
        global conn,cur
        try:
            conn=pyodbc.connect(r'DRIVER={SQL Server};SERVER=self.server;DATABASE=self.database;UID=self.uid;PWD=self.pwd')
            cur=conn.cursor()
            return cur
        except pyodbc.OperationalError as err:
            mes.showerror('错误提示','找不到服务器：%s' % str(err))
        except pyodbc.InterfaceError as err:
            mes.showerror('错误提示','数据库连接失败：%s' % str(err))

    def closedb():
        cur.close()
        conn.close()

        
