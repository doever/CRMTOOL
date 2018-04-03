# CRMTOOL

```
This's Tools for Customer Relation System
```
### main function
- [x] Cancel order
- [x] Finish order
- [x] Send ETL email
- [X] Export data
- [x] Monitor abnormal order
- [X] help document
***
### myblog
[visit it now!](http://www.cnblogs.com/chilo/)
***
## 注意事项

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;使用pyodbc需要安装微软官方的Native Client（没有安装会报错IM002)，安装SQL server management studio会自动附带安装（控制面板里可以看到安装的版本）。如果没有安装过需要点此处：
[驱动链接](https://msdn.microsoft.com/en-us/data/ff658533.aspx)
下载安装(sqlncli.msi)。建议选择与远程数据库版本相对应的Native Client。如果本地安装的Native Client是高版本，则DRIVER={SQL Server Native Client 11.0}需要填写的是本地的高版本。  
找到对应的DRIVER，按此DRIVER设置pyodbc的connect的DRIVER  
方法：控制面板-管理工具-数据源odbc-系统DSN--点击添加即可查看安装的数据源名称，即为pyodbc.connect的DRIVER
<br/>其它模块直接pip即可</br>
