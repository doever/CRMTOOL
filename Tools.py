#!/usr/bin/python3
# -*- coding: UTF-8 -*-
import tkinter as tk
import pyodbc
from tkinter import messagebox as mes
import re
import os
import datetime
from tkinter import ttk
import webbrowser
import xlsxwriter
import requests
import json
####浩泽撤单####
#未try catch
def conn_test():
	try:
		conn=pyodbc.connect(r'DRIVER={SQL Server};SERVER=localhost;DATABASE=KFT_MAIN;UID=chan;PWD=password')
		conn.close()
	except pyodbc.OperationalError as reason1:
		mes.showerror("错误提示","找不到服务器,原因是:%s" % str(reason1))
	except pyodbc.InterfaceError as reason:
		mes.showerror("错误提示","数据连接失败,原因是:%s" % str(reason))
	else:
		mes.showinfo("连接提示","连接成功")


#选择模式
def model_select(func):
	global mo
	mo=model.get()
	return func
		


def hz_channelorder():
	showbill_hz.delete(1.0,"end")
	try:
		#正式库
		conn=pyodbc.connect(r'DRIVER={SQL Server};SERVER=localhost;DATABASE=KFT_MAIN;UID=chan;PWD=password')
		#测试库
		#conn=pyodbc.connect(r'DRIVER={SQL Server};SERVER=localhost;DATABASE=KFT_MAIN_TEST;UID=chan;PWD=password')
		cur=conn.cursor()
	except pyodbc.OperationalError as reason1:
		mes.showerror("错误提示","找不到服务器,原因是:%s" % str(reason1))
	except pyodbc.InterfaceError as reason:
		mes.showerror("错误提示","数据连接失败,原因是:%s" % str(reason))
#获取输入值;  try在此处少个else
	cd=str(t_hz.get("1.0","end"))
	cd=cd.strip()
	if cd == '':
		mes.showerror("系统提示","未输入单号")
	else:
		answer=mes.askyesno('tips','是否确定撤单？')
		if answer is False:
			pass
		else:
			billnos=re.findall(r'\w{2}\d{6}-\d{4}',cd)
			billnos = [cd] if billnos == [] else billnos
			for billno in billnos:
				types=billno[0:2]
				if types == 'KH':
					select_sql="select billstate from kft_applynew where billno='%s'" % (str(billno))
				elif types == 'WX':
					select_sql="select billstate from kft_repair where billno='%s'" % (str(billno))
				else:
					select_sql="select billstate from kft_exchange where billno='%s'" % (str(billno))
				#print(select_sql)
				cur.execute(select_sql)
				test=cur.fetchone()
				print(test)
				if test is None:
					info="%s单据在系统未找到，请检查单据是否输入错误" % (str(billno))
					hz_print_info(info)
					save_log(info)
				else:
					state_id=str(test[0])
					#if (billstate=='150' or billstate=='330' or billstate=='430' or billstate=='130'):
					if state_id in ('130','330','-100','140','-300','-400'):
						#可能会中断循环操作
						#mes.showerror("拒绝操作","%s单据是已完成或已撤单状态,无法撤单" % (str(billno)))
						info="%s单据是已完成或已撤单状态,无法撤单" % (str(billno))
						hz_print_info(info)
						save_log(info)
					else:
						update_sql="exec cd '%s'" % (str(billno))
						#print(update_sql)
						try:
							cur.execute(update_sql)
							conn.commit()
							info = "%s单据已撤单" % (str(billno))
							hz_print_info(info)
							save_log(info)
						except:
							conn.rollback()
							mes.showerror("发生错误","已回滚数据")


		cur.close()
		conn.close()
#####浩优撤单######
@model_select
def hy_channelorder():
	showbill.delete(1.0,"end")
	#mes.showinfo("浩优撤单","敬请期待")
	#model_select()
	#正式库
	con=pyodbc.connect(r'DRIVER={SQL Server};SERVER=localhost;DATABASE=S60_MAIN;UID=chan;PWD=password')
	curo=con.cursor()
	if mo == '':
		mes.showwarning("系统提示", "请选择模式")
	elif mo == 'a':
		cd=str(t.get("1.0","end"))
		cd=cd.strip()
		if cd == '':
			mes.showerror("系统提示","未输入单号")
		else:
			answer = mes.askyesno('浩优操作tips', '是否确定撤单？')
			if answer is False:
				pass
			else:
				billnos=re.findall(r'\w{2}\d{6}-\d{4}',cd)
				# print(cd)
				billnos = [cd] if billnos == [] else billnos
				for billno in billnos:
					select_sql="select billstate from dbo.mx_repair where billno='%s'" % (str(billno))
					curo.execute(select_sql)
					test=curo.fetchone()
					if test is None:
						info="%s单据未找到，请确认是否输入正确" % (str(billno))
						print_info(info)
						save_log(info)
					else:
						state_id=str(test[0])
						states=('130000009','130000010','130000011')
						if state_id in states:
							info="%s单据状态已完成无法撤单" % str(billno)
							print_info(info)
							save_log(info)
						else:
							update_sql="update mx_repair set billstate=130000005,isdel=1,remark=remark+'小工具撤单' where billno='%s'" % str(billno)
							update_sql2="update Family_Base_Service_Order set CRMCheck=110000006,CRMRemark=CRMRemark+'小工具撤单' where CRMID='%s'" % str(billno)
							curo.execute(update_sql)
							curo.execute(update_sql2)
							con.commit()
							info="%s单据浩优&&app已撤单" % str(billno)
							print_info(info)
							save_log(info)
	else:
		cd=str(t.get("1.0","end"))
		cd=cd.strip()
		if cd == '':
			mes.showerror("系统提示","未输入单号")
		else:
			answer = mes.askyesno('浩优操作tips', '是否确定撤单？')
			if answer is False:
				pass
			else:
				billnos=re.findall(r'\w{2}\d{6}-\d{5}',cd)
				billnos=[cd] if billnos==[] else billnos
				for billno in billnos:
					select_sql="select billstate from mx_repair where billno='%s'" % (str(billno))
					curo.execute(select_sql)
					test=curo.fetchone()
					if test is None:
						info="%s单据未找到，请确认是否输入正确" % (str(billno))
						print_info(info)
						save_log(info)
					else:
						state_id = str(test[0])
						states=('130000008','130000009','130000010','130000011','130000012','130000013','130000014','130000015')
						if state_id in states:
							info="%s单据状态已完成无法撤单" % str(billno)
							print_info(info)
							save_log(info)
						else:
							update_sql="update mx_repair set billstate=130000005,isdel=1 where billno='%s'" % str(billno)
							update_sql2="update Family_Base_Service_Order set CRMCheck=110000006 where CRMID='%s'" % str(billno)
							curo.execute(update_sql)
							curo.execute(update_sql2)
							con.commit()
							info="%s单据浩优&&app已撤单" % str(billno)
							print_info(info)
							save_log(info)
#函数结束后关闭数据库连接
	curo.close()
	con.close()
###############浩泽结单#################
def hz_finishorder():
	showbill_hz.delete(1.0, "end")
	HZ_billno=str(t_hz.get("1.0","end"))
	HZ_billno=HZ_billno.strip()
	if HZ_billno=='':
		mes.showerror("错误提示","未输入单号！！")
	else:
		answer = mes.askyesno('tips', '是否确定结单？')
		if answer is False:
			pass
		else:
			billnos=re.findall(r'\w{2}\d{6}-\d{4}',HZ_billno)
			urls=['http://192.168.1.106/api/ApplyNew_Completed.do',
				  'http://192.168.1.106/api/Repair_Completed.do',
				  'http://192.168.1.106/api/Exchange_Completed.do']
			# print(billnos)
			# print(type(billnos))
			infos=''
			for billno in billnos:
				bill_type=billno[0:2]
				if bill_type=='KH':
					post_url = urls[0]
				elif bill_type=='WX':
					post_url = urls[1]
				else:
					post_url = urls[2]
				post_billno={'BillNo':billno}
				request=requests.post(url=post_url,data=post_billno)
				# print(request.status_code)
				rest=request.text
				# print(rest)
				# print(type(rest))
				##使用loads方法将response的str转换成dict
				rest=json.loads(rest)
				# print(type(rest))
				if str(rest['BillState']) in ('150', '330', '430'):
					state = '已完成'
				else:
					state = '未完成'
				if str(rest['Result'])=='0':
					info=str(rest['BillNo'])+'的状态为'+state +' '+str(rest['BillState'])+'：状态同步success！ '+str(rest['Result'])+'\n'
				else:
					info=str(rest['BillNo'])+'的状态为'+ state+' '+str(rest['BillState'])+'：状态同步failed！ '+str(rest['Result'])+'\n'
				infos=infos+info

			hz_print_info(infos)
			save_log(infos)

			# print(rests)
			# print(type(rests[0]))
			# # ##rests=[{"Result":0,"BillNo":"KH171018-1254","BillState":150},{"Result":0,"BillNo":"KH171018-1254","BillState":150}]
			# for rest in rests:
			# 	if rest['Result']==0:
			# 		info=str(rest['BillNo'])+'：success！'+'\n'
			# 	else:
			# 		info=str(rest['BillNo'])+'：failed！'+'\n'
			# 	infos=infos+info


# @model_select #hy_finishorder=model_select(model_select)
@model_select
def hy_finishorder():
	showbill.delete(1.0, "end")
	#model_select()
	# 正式库
	con = pyodbc.connect(r'DRIVER={SQL Server};SERVER=localhost;DATABASE=S60_MAIN;UID=chan;PWD=password')
	curo = con.cursor()
	if mo == '':
		mes.showwarning("系统提示", "请选择模式")
	elif mo == 'a':
		cd = str(t.get("1.0", "end"))
		cd = cd.strip()
		if cd == '':
			mes.showerror("系统提示", "未输入单号")
		else:
			answer = mes.askyesno('浩优结单tips', '是否确定结单？')
			if answer is False:
				pass
			else:

				billnos = re.findall(r'\w{2}\d{6}-\d{4}', cd)
				# 因为billnos若为空的列表，会导致for运行没任何结果，因此给此处billnos一个默认的值
				billnos = [cd] if billnos==[] else billnos
				for billno in billnos:
					select_sql = "select billstate from dbo.mx_repair where billno='%s'" % (str(billno))
					curo.execute(select_sql)
					test = curo.fetchone()
					if test is None:
						info = "%s单据未找到，请确认是否输入正确" % (str(billno))
						print_info(info)
						save_log(info)
					else:
						state_id = str(test[0])
						states = ('130000008')
						if state_id in states:
							info = "%s单据状态已完成无法结单" % str(billno)
							print_info(info)
							save_log(info)
						else:
							update_sql = "update mx_repair set billstate=130000008,remark=remark+'tools结单' where billno='%s'" % str(billno)
							# app不需要我们结单
							# update_sql2 = "update Family_Base_Service_Order set CRMCheck=110000006,CRMRemark=CRMRemark+'小工具撤单' where CRMID='%s'" % str(billno)
							curo.execute(update_sql)
							# curo.execute(update_sql2)
							con.commit()
							info = "%s单据浩优已结单" % str(billno)
							print_info(info)
							save_log(info)

	else:
		cd = str(t.get("1.0", "end"))
		cd = cd.strip()
		if cd == '':
			mes.showerror("系统提示", "未输入单号")
		else:
			answer = mes.askyesno('浩优操作tips', '是否确定结单？')
			if answer is False:
				pass
			else:
				billnos = re.findall(r'\w{2}\d{6}-\d{5}', cd)
				billnos = [cd] if billnos == [] else billnos
				for billno in billnos:
					select_sql = "select billstate from mx_repair where billno='%s'" % (str(billno))
					curo.execute(select_sql)
					test = curo.fetchone()
					if test is None:
						info = "%s单据未找到，请确认是否输入正确" % (str(billno))
						print_info(info)
						save_log(info)
					else:
						state_id = str(test[0])
						states = ('130000008')
						if state_id in states:
							info = "%s单据状态已完成无需结单" % str(billno)
							print_info(info)
							save_log(info)
						else:
							update_sql = "update mx_repair set billstate=130000008,remark=remark+'tools结单' where billno='%s'" % str(billno)
							# update_sql2 = "update Family_Base_Service_Order set CRMCheck=110000006 where CRMID='%s'" % str(billno)
							curo.execute(update_sql)
							# curo.execute(update_sql2)
							con.commit()
							info = "%s单据浩优已结单" % str(billno)
							print_info(info)
							save_log(info)
	# 函数结束后关闭数据库连接
	curo.close()
	con.close()
	# mes.showinfo("浩优结单","敬请期待")
#选择模式
def model_select():
	global mo
	mo=model.get()

#使用装饰器,这里加上装饰器等于运行了model_select，然后func=func，没什么改变
# def model_select(func):
# 	global mo
# 	mo=model.get()
# 	return func

#打印消息
def print_info(info):
	# showbill.delete(1.0,"end")
	showbill.insert('end',"\n" + info + "\n------------------------------------------------------------")
	showbill.see("end")
def hz_print_info(info):
	# showbill_hz.delete(1.0,"end")
	showbill_hz.insert('end',"\n" + info + "\n------------------------------------------------------------")
	showbill_hz.see("end")
#输出日志
def save_log(info):
	#已追加方式打开文件,需要在此路径下有log目录才可生成日志
	f=open("D:/CRMTOOL/log/CRMlog.txt","a")
	now=str(datetime.datetime.now())
	f.write(now+' : '+info+"\n")
	f.close()

#######sql导出记录日志########

def save_sqllog(sqlinfo):
	###最多保存36个月的txt，可使用bat脚本实现删除,先检查sqllog文件数量是否过多
	# i=0
	# for root, dirs, files in os.walk('D:/CRMTOOL/log'):
	# 	for f in files:
	# 		# print(f)
	# 		i += 1
	#
	# if i >= 37:
	# 	answer = mes.askyesno("清理提示","sqllog文件已经超过60个月没清理了，是否需要清理？")
	# 	if answer is True:
	# 		#使用call命令乱码，暂时不加自动清理功能
	# 		os.system("call D:/CRMTOOL/clean.bat")
	# 		mes.showinfo("success","成功清理")
	# 	else:
	# 		pass
	# else:
	# 	pass
	###每月创建一个年月日期开头的sqllog.txt
		#为了适应bat的日期格式加入此判断条件，为一位字符的月份补0
	if len(str(datetime.datetime.now().month)) == 1:
		log_head=str(datetime.datetime.now().year)+"0"+str(datetime.datetime.now().month)
	else:
		log_head = str(datetime.datetime.now().year) + str(datetime.datetime.now().month)
	f=open("D:/CRMTOOL/log/"+log_head+"SQLlog.txt","a")
	now=str(datetime.datetime.now())
	f.write(now+ ' : ' + sqlinfo+"\n")
	f.close()
######单据替换功能#######
def changetext():
	model_select_2()
#需要return mo的值 mode才能被赋值
#	mode=model_select()
#	print(mode)
#	print(mo)
	#mode=mode.strip()
	if mo2 == '':
		mes.showwarning("贴心小提示", "请选择模式")
	else:
		#print(mo)
		if mo2 == 'a':
			cd=str(entryCd.get("1.0","end"))
			cd=cd.strip()
			#print(cd)
			if cd == '':
				mes.showerror("贴心小提示", "未输入有效单号")
			else:
				entryCd2.delete(1.0,"end")
				m=re.findall(r'\w{2}\d{6}-\d{4}',cd)
				m1=re.findall(r'\w{2}\d{11}',cd)
				#列表数据拼接
				res=m+m1
				#将数据格式化输出，去掉开头结尾的[]加上()
				res1='('+str(res)[1:len(str(res))-1]+')'
#				for i in m:
#					s='\''+i+'\''+','
				entryCd2.insert(1.0,str(res1))
		else:
			cd=str(entryCd.get("1.0","end"))
			cd=cd.strip()
			#print(cd)
			if cd == '':
				mes.showerror("贴心小提示", "未输入有效单号")
			else:
				entryCd2.delete(1.0,"end")
				m=re.findall(r'\w{2}\d{6}-\d{5}',cd)
				m1=re.findall(r'\w{2}\d{11}',cd)
				#列表数据拼接
				res=m+m1
				#将数据格式化输出，去掉开头结尾的[]加上()
				res1='('+str(res)[1:len(str(res))-1]+')'
#				for i in m:
#					s='\''+i+'\''+','
				entryCd2.insert(1.0,str(res1))
def model_select_2():
	global mo2
	mo2=model2.get()
	#print(mo)
	#return mo
#############装机案例导出################

def dump_example():
	sql_city=str(city.get())
	sql_customer=str(customer.get())
	sql_area= str(bill_area.get())
	if sql_area == '省市区？':
		mes.showerror("错误提示", "省市不可为空!")
	else:
		if sql_city == '' and sql_customer == '':
			mes.showerror("错误提示", "案例名称不可为空!")
		elif sql_city != '' and sql_customer == '':
			mes.showerror("错误提示","案例名称不可为空!")
		elif sql_city == '' and sql_customer != '':
			customers=sql_customer.split("，")
			sql_data=''
			for customer_name in customers:
				sql_data = sql_data + " a.name like '%"+ customer_name + "%' or"
			sql_data="("+sql_data[0:len(sql_data)-3]+")"
		else:
			customers=sql_customer.split("，")
			sql_data1=''
			sql_data2=''
			for customer_name in customers:
				sql_data1= sql_data1 + " a.name like '%"+ customer_name + "%' or"
			sql_data1="("+sql_data1[0:len(sql_data1)-3]+")"
			# print("sql_data1:"+sql_data1)
			select_city=sql_city.split("，")
			if sql_area =='省':
				for pro_name in select_city:
					sql_data2 = sql_data2 +" b.province like '%" + pro_name + "%' or"
			else:
				for city_name in select_city:
					sql_data2 = sql_data2 +" c.city like '%" + city_name + "%' or"
			sql_data2="("+sql_data2[0:len(sql_data2)-3]+")"
			sql_data = sql_data1 + ' and ' + sql_data2
			# print("sql_data2:" + sql_data2)
	sql_head = '''select  a.Name as 客户名称,max(b.Province) as 省,max(c.city) as 市  from kft_applynew a
					left join KFT_Province b on a.Province=b.provinceID
					left join KFT_City c on a.city=c.cityID
					left join KFT_Area d on a.Area=d.areaID
					where a.isspecial <> 1  and (a.A3 is NULL or a.A3 = '')  
					and a.isdel <> 1  and a.billstate=150 and  '''
	final_sql = sql_head + sql_data + " group by a.Name  order by a.name "
	# print(final_sql)
	anli_dumpexecl(final_sql)

###########退换移数据导出功能###########
def dump_exchange():
	mes.showinfo("tips","敬请期待")
############开户数据导出功能############
def dump_applynew():
	date_type = str(datetype.get())
	state = str(billstate.get())
	startdate = str(syear.get()) + str(smonth.get()) + str(sday.get()) + ' 00:00:00'
	enddate = str(eyear.get()) + str(emonth.get()) + str(eday.get()) + ' 23:59:59'
	code = str(areacode.get())
	pid = str(machine.get())
	##时间类型跟状态有一条未填写
	if date_type == '请选择' or state == '请选择':
		mes.showwarning("贴心提示：不影响导出", "时间类型&状态真的不选吗")
		##分为三种情况：时间类型有状态没有，时间类型没有状态有,二者都没有
		if date_type != '请选择' and state == '请选择':
			date_type = 'a.createdate' if str(datetype.get()) == '制单日期' else 'a.InstalDate'
			#若选择时间类型，则必须选择时间，否则报错，函数结束,未给sql_data赋值，会引发 赋值前引用 错误
			if len(startdate) != 17 or len(enddate) != 17:
				# mes.showerror("错误","未选择时间")
				##对选择的时间还要进行判断，也分三种情况
				if len(startdate) == 17 and len(enddate) != 17:
					sql_data = ' and '+date_type+'>='+"'"+startdate+"'"
				elif len(startdate) != 17 and len(enddate) == 17:
					sql_data = ' and ' + date_type + '<=' + "'" + enddate + "'"
				else:
					# 若用户至少填写了区域代码或者机型，则不应该限制导出
					if code != '' or pid != '':
						sql_data = ' '
					# 用户未填写区域代码和机型，且时间只选了开始或结束时间，却被提示单据量过多，不够严谨，最好由len(rows)判断，但不知道
					# rows能否装载大量的数据
					else:
						mes.showinfo("小提示:不影响导出","因单据量可能过大，仅导出往前一天的单据参考")
						# sql_data = ' and a.createdate between getdate()-1 and getdate()'
						sql_data = ' '
			else:
				sql_data = ' and '+date_type + '>' + "'" + startdate + "'" + ' and ' + date_type + '<' + "'" + enddate + "'"
		elif date_type == '请选择' and state != '请选择':
			state = 'and a.billstate=150' if str(billstate.get()) == '已完成' else ' '
			if code != '' or pid != '':
				sql_data = state
			else:
				mes.showinfo("tips:不影响导出","因单据量可能过大，仅导出往前一天的单据参考")
				sql_data = state + ' and a.createdate between getdate()-1 and getdate()'
		else:
			# 若用户至少填写了区域代码或者机型，则不应该限制导出
			if code != '' or pid != '':
				sql_data = ' '
			# 用户未加任何限制，限制导出
			else:
				mes.showinfo("小提示:不影响导出","因数据量过大，仅导出往前一天的单据参考")
				sql_data = ' and a.createdate between getdate()-1 and getdate()'
	##时间类型跟状态都填了
	else:
		#三目运算，加好sql条件
		date_type = 'a.createdate' if str(datetype.get()) == '制单日期' else 'a.InstalDate'
		state=' and a.billstate=150' if str(billstate.get()) == '已完成' else ' '
		if len(startdate) != 17 or len(enddate) != 17:
			mes.showwarning("贴心提示：不影响导出", "真的不把时间选全吗？")
			#两个都不选
			if len(startdate) != 17 and len(enddate) != 17:
				sql_data = ' '+state
			#只选了开始日期
			elif len(startdate) == 17 and len(enddate) != 17:
				sql_data = ' and '+date_type+'>'+"'"+startdate+"'"+state
			#只选了结束日期
			else:
				sql_data = ' and '+date_type+'<'+"'"+enddate+"'"+state
		else:
			sql_data=' and '+date_type+'>'+"'"+startdate+"'"+' and '+date_type+'<'+"'"+enddate+"'"+state

	# print("sql_data:"+sql_data)
	#####判断是否输入代码或机型，然后拼接条件语句sql_data
	if code == '' and pid == '':
		pass
	elif code != '' and pid == '':
		codes = code.split(',')
		acode = ''
		for co in codes:
			acode= acode + " areacode like '%" + co + "%' or "
		acode=' and ('+acode[0:len(acode)-3] + ')'
		sql_data=sql_data+acode
		# print(acode)
		# print(sql_data)
	elif code == '' and pid != '':
		pids = pid.split(',')
		apid = ''
		for co in pids:
			apid= apid + " machinesid ='" + co + "' or "
			# print(apid)
		apid=' and ('+apid[0:len(apid)-3] + ')'
		sql_data = sql_data + apid
		# print(apid)
		# print(sql_data)
	else:
		codes = code.split(',')
		acode = ''
		for co in codes:
			acode= acode + " areacode like '%" + co + "%' or "
		acode = ' and (' + acode[0:len(acode) - 3] + ')'
		pids = pid.split(',')
		apid = ''
		for co in pids:
			apid = apid + " machinesid ='" + co + "' or "
		apid = ' and (' + apid[0:len(apid) - 3] + ')'
		sql_data = sql_data+acode + apid
	# print(sql_data)
	sql_head='''select 
		a.billno         开户单号,
		a.cardcode       开户卡号,
		a.areaCode       区域代码,
		jx.ProName       机器型号,
		a.name           客户名称,
		a.A5             代理商名称,
		--fl.ClassName     客户分类,
		--a.remark       备注,
		br.branchname    网点,
		--u1.name          业务员ID,
		u2.name          代理商ID,
		--u3.name          安装工程师ID,
		--a.tel            电话,
		--a.mobile	     手机,
		--a.address        地址,
		u10.province     省,
		u11.city         市,
		u12.area         区,
		--u4.name          维修工程师,
		a.machinecode    机器编号,
		--d1.name          渠道平台,
		--a.y_tds          源水TDS,
		--a.z_tds          活水TD,
		--a.shuiya         水压,
		--fw.[name]        服务满意度,
		--a.dissatisfied   不满意事项,
		a.instaldate     安装时间,
		--a.deliverydate   送货日期,
		--u5.name          发货人,
		bi.statename     单据状态,
		a.createdate     制单日期,
		--a.activadate     激活日期,
		--a.extend1        安装确认单号,
		--a.machineversion 机器版本,
		--a.BillSort       单据类型,
		--a.openaudit      开箱情况,
		fl.ClassName     客户分类
		from kft_applynew a
		left join kft_users u1 on a.seller = u1.userid
		left join kft_users u2 on a.agent = u2.userid
		left join kft_users u3 on a.installer = u3.userid
		left join kft_users u4 on a.repairer = u4.userid
		left join kft_users u5 on a.delivery = u5.userid
		left join kft_users u7 on a.cashier = u7.userid
		left join kft_province u10 on a.province = u10.provinceid
		left join kft_city u11 on a.city = u11.cityid
		left join kft_area u12 on a.area = u12.areaid
		left join kft_billstate bi on a.billstate = bi.stateid
		left join kft_branch br on a.branchno=br.branchno 
		left join KFT_Dictionary fw on a.satisfaction=fw.[id]
		left join KFT_Products jx on a.MachineSID = jx.ProSID
		left join KFT_Customers KE on a.CustomerNo = KE.CustomerNo
		left JOin KFT_CusClass fl on KE.CusClass = fl.ClassID
		left join  kft_dictionary d1 on a.A7=d1.id
		where a.isspecial <> 1 and (a.A3 is NULL or a.A3 = '')
			  and a.isdel <> 1	
	'''
	final_sql = sql_head + sql_data
	# print(final_sql)
	dumpexecl(final_sql)

############导出开户数据到execl功能#########
def dumpexecl(final_sql):
	###导出数据连接镜像数据库
	try:
		conn=pyodbc.connect(r'DRIVER={SQL Server Native Client 10.0};SERVER=localhost;DATABASE=KFT_SNAP;UID=chan;PWD=password')
		cur=conn.cursor()
		cur.execute(final_sql)
		rows=cur.fetchall()
		cur.close()
		conn.close()
	except pyodbc.OperationalError:
		mes.showerror("错误提示","无法连接248镜像服务器")
	except pyodbc.InterfaceError:
		mes.showerror("错误提示二","无法连接数据库")
	except pyodbc.ProgrammingError:
		mes.showerror("错误提示三","查询语句有错误，请在日志查看SQLlog")
		sqlinfo="\n" + '出错语句：' + final_sql +"\n"
		save_sqllog(sqlinfo)
	else:
		###若数量大于5000，限制导出
		if len(rows) >= 5000:
			mes.showerror("错误提示","导出数据数量大于5000，禁止导出,请联系管理员")
		else:
			#格式化日期，取纯数字做execl表名
			now1 = str(datetime.datetime.now().year)+str(datetime.datetime.now().month)+str(datetime.datetime.now().day)
			now2 = str(datetime.datetime.now().hour)+str(datetime.datetime.now().minute)+str(datetime.datetime.now().second)
			now = now1+'_'+now2
			os.chdir("C:/Users/Administrator/Desktop")
			workbook=xlsxwriter.Workbook(now+'KH.xlsx')
			worksheet=workbook.add_worksheet("applynew")
			worksheet.write(0, 0, "开户卡")
			worksheet.write(0, 1, "开户卡号")
			worksheet.write(0, 2, "区域代码")
			worksheet.write(0, 3, "机器型号")
			worksheet.write(0, 4, "客户名称")
			worksheet.write(0, 5, "代理商名称")
			worksheet.write(0, 6, "网点")
			worksheet.write(0, 7, "代理商ID")
			worksheet.write(0, 8, "省")
			worksheet.write(0, 9, "市")
			worksheet.write(0, 10, "区")
			worksheet.write(0, 11, "机器编号")
			worksheet.write(0, 12, "安装时间")
			worksheet.write(0, 13, "单据状态")
			worksheet.write(0, 14, "制单日期")
			worksheet.write(0, 15, "客户分类")
			# worksheet.write(0, 15, "")
			row = 1
			col = 0
			for billno, cardcode, code, macname, name, agent, branch, agentid, province, city, area, macode, insdate, bstate, createdate, ClassName in rows:
				worksheet.write(row, col, billno)
				worksheet.write(row, col + 1, cardcode)
				worksheet.write(row, col + 2, code)
				worksheet.write(row, col + 3, macname)
				worksheet.write(row, col + 4, name)
				worksheet.write(row, col + 5, agent)
				worksheet.write(row, col + 6, branch)
				worksheet.write(row, col + 7, agentid)
				worksheet.write(row, col + 8, province)
				worksheet.write(row, col + 9, city)
				worksheet.write(row, col + 10, area)
				worksheet.write(row, col + 11, macode)
				worksheet.write(row, col + 12, insdate)
				worksheet.write(row, col + 13, bstate)
				worksheet.write(row, col + 14, createdate)
				worksheet.write(row, col + 15, ClassName)
				row += 1
			workbook.close()
			mes.showinfo("success","成功导出数据至桌面")
			save_sqllog(final_sql)
###########导出案例到execl############
def anli_dumpexecl(final_sql):
	try:
		conn = pyodbc.connect(r'DRIVER={SQL Server Native Client 10.0};SERVER=localhost;DATABASE=KFT_SNAP;UID=chan;PWD=password')
		cur = conn.cursor()
		cur.execute(final_sql)
		rows = cur.fetchall()
		cur.close()
		conn.close()
	except pyodbc.OperationalError:
		mes.showerror("错误提示", "无法连接248镜像服务器")
	except pyodbc.InterfaceError:
		mes.showerror("错误提示二", "无法连接数据库")
	except pyodbc.ProgrammingError:
		mes.showerror("错误提示三", "查询语句有错误，请在日志查看SQLlog")
		sqlinfo = "\n" + '出错语句：' + final_sql + "\n"
		save_sqllog(sqlinfo)
	else:
		###若数量很庞大，限制导出
		if len(rows) >= 5000:
			mes.showerror("错误提示", "导出数据数量大于5000，禁止导出,请联系管理员")
		else:
			now1 = str(datetime.datetime.now().year) + str(datetime.datetime.now().month) + str(datetime.datetime.now().day)
			now2 = str(datetime.datetime.now().hour) + str(datetime.datetime.now().minute) + str(datetime.datetime.now().second)
			now = now1 + '_' + now2
			os.chdir("C:/Users/Administrator/Desktop")
			workbook = xlsxwriter.Workbook(now + 'AL.xlsx')
			worksheet = workbook.add_worksheet("installexample")
			worksheet.write(0, 0, '客户名称')
			worksheet.write(0, 1, '省')
			worksheet.write(0, 2, '市')
			row=1
			col=0
			for name,province,cityname in rows:
				worksheet.write(row, col, name)
				worksheet.write(row, col + 1, province)
				worksheet.write(row, col + 2, cityname)
				row+=1
			workbook.close()
			mes.showinfo("success","成功导出数据至桌面")
			save_sqllog(final_sql)
############异常单据功能############
####为画布取数据
def get_bad_order():
	# bad_order=[]
	sql=''
	# global bad_order
	try:
		conn=pyodbc.connect(r'DRIVER={SQL Server Native Client 10.0};SERVER=192.168.172.206;DATABASE=wmsinput;UID=chan;PWD=123Qwe')
		cur=conn.cursor()
		sql='''select top 5 a.BillNo,a.BillSort,b.statename,DATEDIFF(mi,a.createtime,getdate()) as memo
				from [syncdata] a left join dictionary b on a.BillState=b.stateid where memo like '%距%' '''
		cur.execute(sql)
		bad_order=cur.fetchall()
		# print(bad_order)
		cur.close()
		conn.close()
	except pyodbc.OperationalError:
		mes.showerror("错误提示", "无法连接206报表服务器")
	except pyodbc.InterfaceError:
		mes.showerror("错误提示二", "无法连接数据库")
	except pyodbc.ProgrammingError:
		mes.showerror("错误提示三", "查询语句有错误，请在日志查看SQLlog")
		sqlinfo = "\n" + '监控出错语句：' + sql + "\n"
		save_sqllog(sqlinfo)
	return bad_order
##填充数据功能

def output_bad_order():
	orders = get_bad_order()
	# canvas_8.create_text(92, 142, text=orders[0][0],fill="orange",font=("Verdana",10))
	# canvas_8.create_text(216, 142, text=orders[0][1], fill="orange", font=("Verdana", 10))
	# canvas_8.create_text(349, 142, text=orders[0][2], fill="orange", font=("Verdana", 10))
	# canvas_8.create_text(482, 142, text=orders[0][3], fill="orange", font=("Verdana", 10))
	# canvas_8.create_text(92, 187, text=orders[1][0], fill="orange", font=("Verdana", 10))
	# canvas_8.create_text(216, 187, text=orders[1][1], fill="orange", font=("Verdana", 10))
	# canvas_8.create_text(349, 187, text=orders[1][2], fill="orange", font=("Verdana", 10))
	# canvas_8.create_text(482, 187, text=orders[1][3], fill="orange", font=("Verdana", 10))
	y=142
	try:
		for i in range(5):
			canvas_8.create_text(92, y, text=orders[i][0], fill="orange", font=("Verdana", 10))
			canvas_8.create_text(216, y, text=orders[i][1], fill="orange", font=("Verdana", 10))
			canvas_8.create_text(349, y, text=orders[i][2], fill="orange", font=("Verdana", 10))
			canvas_8.create_text(482, y, text=str(orders[i][3])+'分钟', fill="orange", font=("Verdana", 10))
			y=y+44
		raise IndexError
	except IndexError:
		pass
	finally:
		num=len(orders)
		global bad_billnos
		bad_billnos = tk.StringVar()
		bad_billno_chose = ttk.Combobox(tab8, width=14, textvariable=bad_billnos)
		if num == 0:
			pass
		elif num == 1:
			bad_billno_chose['values'] = ('选择单据',orders[0][0])
		elif num == 2:
			bad_billno_chose['values'] = ('选择单据',orders[0][0],orders[1][0])
		elif num == 3:
			bad_billno_chose['values'] = ('选择单据',orders[0][0],orders[1][0],orders[2][0])
		elif num == 4:
			bad_billno_chose['values'] = ('选择单据',orders[0][0],orders[1][0],orders[2][0],orders[3][0])
		else:
			bad_billno_chose['values'] = ('选择单据',orders[0][0],orders[1][0],orders[2][0],orders[3][0],orders[4][0])
		bad_billno_chose.place(x=410,y=362)
		bad_billno_chose.current(0)
#####查看异常原因
def select_reason():
	bad_billno=bad_billnos.get()
	if bad_billno[0:2] == 'KH':
		conn = pyodbc.connect(r'DRIVER={SQL Server Native Client 10.0};SERVER=localhost;DATABASE=KFT_SNAP;UID=chan;PWD=password')
		cur = conn.cursor()
		sql="select billstate,BillSort,B1,tel  from kft_applynew where BillNo='%s'" % bad_billno
		cur.execute(sql)
		rows=cur.fetchall()
		cur.close()
		conn.close()
		if len(rows)==0:
			info = "error:billno not find"
		elif rows[0][2]==0:
			solution_id = 1
			info = "Q:单据渠道平台配置了无需服务！\n"   \
				   "A:同步推荐方案，请修改单据渠道平台后，再变更中间库同步状态！"
		elif rows[0][1] in ("智能柜标配","补单","换机开户"):
			solution_id = 2
			info = "Q:单据类型是不需要服务的类型！\n"  \
				   "A:同步推荐方案，请修改单据类型后,再变更中间库同步状态！"
		elif rows[0][0] in (-100,150):
			if rows[0][0]==-100:
				solution_id = 3
				info="Q:单据已撤单无需同步！\n"   \
					 "A:同步推荐方案，可重新开户！"
			else:
				solution_id = 4
				info="Q:单据已完成无法同步！\n"   \
					 "A:同步推荐方案，请先补单退机，再重新开户！"
		else:
			info = "Q:已排查不是由于状态、单据类型、渠道平台问题引起！\n" \
				   "A:同步推荐方案，请联系管理员查看服务器日志！"
	elif bad_billno[0:2] == 'GD':
		conn = pyodbc.connect(r'DRIVER={SQL Server Native Client 10.0};SERVER=localhost;DATABASE=S60_MAIN;UID=chan;PWD=password')
		cur = conn.cursor()
		sql="select a.billstate,a.isinterfaceDesc,a.BranchNo,a.Warehouse,a.Tel,b.ID from MX_Repair a left join MX_Products b on a.proname=b.extend1 where billno='%s'" % bad_billno
		cur.execute(sql)
		rows=cur.fetchall()
		cur.close()
		conn.close()
		if str(rows[0][0]) in ('130000002','130000005'):
			if rows[0][0] == '130000002':
				solution_id = 5
				info="Q:单据尚未接单，无法同步！\n"   \
					 "A:同步推荐方案，请先接单！"
			else:
				solution_id = 6
				info="Q:单据是已撤单状态，无需同步！\n"  \
					 "A:同步推荐方案，请将浩泽单据也同步撤单，重新开户！"
		elif str(rows[0][1]) is None:
			if rows[0][2] is None:
				solution_id = 7
				info="Q:单据网点为空，无法同步！\n"  \
					 "A:同步推荐方案，请将浩泽单据也同步撤单，重新开户！"
			elif rows[0][3] is None:
				solution_id = 8
				info="Q:单据仓库为空，无法同步！\n"  \
					 "A:同步推荐方案，请将浩泽单据也同步撤单，重新开户！"
			elif rows[0][4] is None:
				solution_id = 9
				info="Q:单据电话为空，无法同步！\n"  \
					 "A:同步推荐方案，请将浩泽单据也同步撤单，重新开户！"
			elif rows[0][5] is None:
				solution_id = 10
				info="Q:单据机器物料在浩优未配置，无法同步！\n"  \
					 "A:同步推荐方案，请配置浩优产品表！！"
			else:
				solution_id = 11
				info="Q:已排查不是由于状态、电话、网点、物料仓库问题引起！\n"  \
					 "A:同步推荐方案，请联系管理员查看服务器日志！"
		else:
			if str(rows[0][1])== '成功':
				info='单据已同步！'
			else:
				solution_id = 12
				info="Q:"+str(rows[0][1])+'\n' \
				 "A:同步推荐方案，请联系管理员！"
	else:
		info="单据未选择或输入错误！"
	print(info)
	canvas_8.create_text(290, 480, text=info, fill="orange", font=("Verdana", 13))
	mes.showinfo("提示",info)
	return solution_id
	# if count == 1:
	# 	canvas_8.create_text(290,465,text=info,fill="orange",font=("Verdana", 13))
	# elif count ==2:
	# 	canvas_8.create_text(290, 490, text=info, fill="orange", font=("Verdana", 13))
	# else:
	# 	canvas_8.create_text(290, 525, text=info, fill="orange", font=("Verdana", 13))
	# count+=1

def solution_way():
	mes.showinfo("成功","success!")
	# solution = select_reason()
	solution=1
	if solution == 1:
		pass
	elif solution == 2:
		pass
	elif solution == 3:
		pass
	elif solution == 4:
		pass
	elif solution == 5:
		pass
	elif solution == 6:
		pass
	elif solution == 7:
		pass
	elif solution == 8:
		pass
	elif solution == 9:
		pass
	elif solution == 10:
		pass
	elif solution == 11:
		pass
	elif solution == 12:
		pass
	else:
		pass
###############退出################
def _quit():
	top.quit()
	exit()
def restart():
	top.quit()
	os.system("python.exe D:/CRMTOOL/cdd.py")
#####淘宝链接######
def taobao_link():
# 	url='http://'
	url="https://yingxism.tmall.com/?spm=a220o.1000855.1997427721.d4918089.90ebce3OjTBZA"
	webbrowser.open(url)
def tools_link():
	url='https://www.baidu.com/s?wd=%E5%B7%A5%E5%85%B7&rsv_spt=1&rsv_iqid=0xfc57dbee00022872&issp=1&f=8&rsv_bp=0&rsv_idx=2&ie=utf-8&' \
		'tn=baiduhome_pg&rsv_enter=1&rsv_sug3=1&rsv_sug1=1&rsv_sug7=100&rsv_sug2=0&inputT=1320&rsv_sug4=1321'
	webbrowser.open(url)
def openfile_log():
	try:
		os.system("start log/CRMlog.txt")
	except OSError:
		mes.showerror("错误提示","未找到CRMlog文件")
#####
def openfile_sqllog():
	os.system("start log")
#####菜单功能实现######
def msgBox1():  
	mes.showinfo('系统版本','当前版本:CrmTools_2.0.0')
def msgBox2():  
	mes.showwarning('Python Message Warning Box', '警告：程序出现错误，请检查！')
def msgBox3():  
	mes.showwarning('Python Message Error Box', '错误：程序出现严重错误，请退出！')
def msgBox4():  
	answer = mes.askyesno("问卷小调查", "你喜欢这个工具吗？\n您的选择是：")
	if answer is True:
		mes.showinfo('谢谢参与', '您选择了“是”，谢谢参与！')
	else:
		mes.showinfo('谢谢参与', '您选择了“否”，谢谢参与！')


'''
win = tk.Tk() # Create instance
win.title("Python GUI") # Add a title
tabControl = ttk.Notebook(win) # Create Tab Control
tab1 = ttk.Frame(tabControl) # Create a tab
tabControl.add(tab1, text='Tab 1') # Add the tab
tabControl.pack(expand=1, fill="both") # Pack to make visible
tab2 = ttk.Frame(tabControl) # Add a second tab
tabControl.add(tab2, text='Tab 2') # Make second tab visible
'''
############分离视图
top=tk.Tk()
top.title('CRM TOOLS 3.0.1')
#创建分离控制器
tabControl=ttk.Notebook(top)
tabControl.pack(expand=5, fill="both")
#创建视图零
tab0=ttk.Frame(tabControl)
tabControl.add(tab0, text='首页')
#创建视图一
tab1=ttk.Frame(tabControl)
tabControl.add(tab1, text='浩泽撤单')
#创建视图二
tab2 = ttk.Frame(tabControl)
tabControl.add(tab2, text='浩优单据')
#创建视图三
tab3=ttk.Frame(tabControl)
tabControl.add(tab3,text='自动邮件')
#创建视图四
tab4=ttk.Frame(tabControl)
tabControl.add(tab4,text='替换工具')
#创建视图七
tab7=ttk.Frame(tabControl)
tabControl.add(tab7,text='数据导出')
#创建视图八
tab8=ttk.Frame(tabControl)
tabControl.add(tab8,text='异常单据')
#创建视图五
tab5=ttk.Frame(tabControl)
tabControl.add(tab5,text='每日运势')
#创建视图六
tab6=ttk.Frame(tabControl)
tabControl.add(tab6,text='帮助说明')

######首页########
canvas=tk.Canvas(tab0,height=585,width=580)
canvas.pack()
canvas.create_rectangle(0,0,580,585,fill="black")
canvas.create_text(290,80,text="OZNER",fill='White',font=('Comic Sans MS',25))
canvas.create_text(290,130,text="浩泽小工具",fill='White',font=('Arial',18))

###画星星####
points = [100, 140, 110, 110, 140, 100, 110, 90, 100, 60, 90, 90, 60, 100, 90, 110]
canvas.create_polygon(points, outline="#476042",
			fill='yellow', width=3)
points2 = [480, 340, 490, 310, 520, 300, 490, 290, 480, 260, 470, 290, 440, 300, 470, 310]
canvas.create_polygon(points2, outline="#476042",
			fill='Orange', width=3)
points3 = [100, 540, 110, 510, 140, 500, 110, 490, 100, 460, 90, 490, 60, 500, 90, 510]
canvas.create_polygon(points3, outline="#476042",
			fill='white', width=3)
###导入图片###
img=tk.PhotoImage(file="backpic.png")
canvas.create_image(65,150, anchor=tk.NW, image=img)
# canvas.create_rectangle(0,290,202,294,fill='white',outline='Khaki')
# canvas.create_rectangle(198,294,202,580,fill='white',outline='Khaki')

'''
points = [100, 140, 110, 110, 140, 100, 110, 90, 100, 60, 90, 90, 60, 100, 90, 110]

w.create_polygon(points, outline=python_green, 
			fill='yellow', width=3)
###导入图片
canvas_width = 300
canvas_height =300

master = Tk()

canvas = Canvas(master, 
		   width=canvas_width, 
		   height=canvas_height)
canvas.pack()

img = PhotoImage(file="rocks.ppm")
canvas.create_image(20,20, anchor=NW, image=img)
'''

###########浩泽撤单界面###########

la_hz=tk.Label(tab1,text="浩  泽",fg='blue',font=("Symbol", "15"))
la_hz.pack()
tk.Label(tab1,text='----------------------------------------------------------------------',font=('', 10)).pack()
la1_hz=tk.Label(tab1,text="注:请输入XX000000-0000格式的单号",fg='DeepPink',font=("Comic Sans MS","8"))
la1_hz.pack()
tk.Label(tab1,text='-----------开户退换单据撤单请注意同时作废浩优以及WMS单据--------------',fg='DeepPink',font=("Comic Sans MS","8")).pack()
t_hz=tk.Text(tab1,height=20,width=80)
t_hz.pack()
showbill_hz=tk.Text(tab1,height=15,width=80)
showbill_hz.pack()
#Text颜色实现
'''
#第一个参数为自定义标签的名字
#第二个参数为设置的起始位置，第三个参数为结束位置
#第四个参数为另一个位置
showbill.tag_add('tag1','1.0','end')
#用tag_config函数来设置标签的属性
showbill.tag_config('tag1',background='LightCyan',foreground='red')
'''
b1=tk.Button(tab1,text='浩泽撤单',activebackground='blue',activeforeground='Black',bg='PaleTurquoise',fg='black',command=hz_channelorder)
b1.pack(side=tk.LEFT)
b=tk.Button(tab1,text='浩泽结单',activebackground='blue',activeforeground='Black',bg='PaleTurquoise',fg='black',command=hz_finishorder)
b.pack(side=tk.RIGHT)


###########浩优撤单界面###########

la=tk.Label(tab2,text="灏  优",fg='blue',font=("Symbol", "15"))
la.pack()
la1=tk.Label(tab2,text="注:浩优单据需选择模式,模式一对应尾数4位数单据,模式二对应尾数5位数单据,请勿混用",fg='DeepPink',font=("Comic Sans MS","8"))
la1.pack()
model=tk.StringVar()
tk.Radiobutton(tab2,text='模式1:GD000000-0000 ',variable=model,value='a',command=model_select,font=('', 10)).pack()
tk.Radiobutton(tab2,text='模式2:GD000000-00000',variable=model,value='b',command=model_select,font=('', 10)).pack()
t=tk.Text(tab2,height=20,width=80)
t.pack()
showbill=tk.Text(tab2,height=15,width=80)
showbill.pack()
#Text颜色实现
'''
#第一个参数为自定义标签的名字
#第二个参数为设置的起始位置，第三个参数为结束位置
#第四个参数为另一个位置
showbill.tag_add('tag1','1.0','end')
#用tag_config函数来设置标签的属性
showbill.tag_config('tag1',background='LightCyan',foreground='red')
'''
b3=tk.Button(tab2,text='浩优撤单',activebackground='yellow',activeforeground='Black',bg='BlanchedAlmond',fg='black',command=hy_channelorder)
b3.pack(side=tk.LEFT)
b4=tk.Button(tab2,text='浩优结单',activebackground='yellow',activeforeground='Black',bg='BlanchedAlmond',fg='black',command=hy_finishorder)
b4.pack(side=tk.RIGHT)

###########自动邮件###########
canvas_3=tk.Canvas(tab3,height=585,width=580)
canvas_3.pack()
#canvas_3.create_rectangle(0,0,580,585,fill="white")
img_3=tk.PhotoImage(file="image/email1.png")
canvas_3.create_image(0,200,anchor=tk.NW,image=img_3)
canvas_3.create_rectangle(0,0,580,200,fill="#FFF")
canvas_3.create_text(290,35,text="ETL脚本文件",fill='black',font=('Comic Sans MS',25))
canvas_3.create_rectangle(0,0,6,200,fill="skyblue")
canvas_3.create_rectangle(0,0,580,6,fill="skyblue")
canvas_3.create_rectangle(576,0,580,200,fill="skyblue")
canvas_3.create_rectangle(0,196,580,200,fill="skyblue")
bat_1=tk.Button(tab3,text='浩泽报表',width=12,fg='#999',bg='Gold',activebackground='#00F',command=lambda: os.system("call ETLbat/OznerReport.bat"))
bat_1.place(x=100,y=100)
bat_2=tk.Button(tab3,text='财务数据',width=12,fg='#999',bg='Gold',activebackground='#00F',command=lambda: os.system("call ETLbat/FinanceData.bat"))
bat_2.place(x=250,y=100)
bat_3=tk.Button(tab3,text='水芯片',width=12,fg='#999',bg='Gold',activebackground='#00F',command=lambda: os.system("call ETLbat/SXP.bat"))
bat_3.place(x=400,y=100)
bat_4=tk.Button(tab3,text='每月五号',width=12,fg='#999',bg='Gold',activebackground='#00F',command=lambda: os.system("call ETLbat/FiveDay.bat"))
bat_4.place(x=100,y=150)
bat_5=tk.Button(tab3,text='清理数据',width=12,fg='#999',bg='Gold',activebackground='#00F',command=lambda: os.system("call ETLbat/clean.bat"))
bat_5.place(x=250,y=150)
bat_6=tk.Button(tab3,text='重启电脑',width=12,fg='#999',bg='Gold',activebackground='#00F',command=lambda: os.system("call ETLbat/Restart.bat"))
bat_6.place(x=400,y=150)
###########单据替换###########
la_2=tk.Label(tab4,text="单据替换",fg='blue',font=("Symbol", "15"))
la_2.pack()
la2_2=tk.Label(tab4,text="注：模式一用于XX111111-1111格式",font=("Times", "8", "bold italic"),fg="red")
la2_2.pack()
la3_2=tk.Label(tab4,text="注：模式二用于XX111111-11111格式",font=("Times", "8", "bold italic"),fg="red")
la3_2.pack()
#模式选择
model2=tk.StringVar()
tk.Radiobutton(tab4,text='模式一',variable=model2,value='a',command=model_select_2,font=('', 8)).pack(side=tk.TOP)
tk.Radiobutton(tab4,text='模式二',variable=model2,value='b',command=model_select_2,font=('', 8)).pack(side=tk.TOP)

entryCd=tk.Text(tab4,height=15,width=80)
entryCd.pack()
entryCd2=tk.Text(tab4,height=15,width=80)
entryCd2.pack()
la_3=tk.Button(tab4,text='替换',width=8,fg='red',activebackground='green',command=changetext)
la_3.pack()
##################################################数据导出#############################################
'''
# 创建一个下拉列表
number = tk.StringVar()
numberChosen = ttk.Combobox(win, width=12, textvariable=number)
numberChosen['values'] = (1, 2, 4, 42, 100)     # 设置下拉列表的值
numberChosen.grid(column=1, row=1)      # 设置其在界面中出现的位置  column代表列   row 代表行
numberChosen.current(0)    # 设置下拉列表默认显示的值，0为 numberChosen['values'] 的下标值
-------------------------------------------复选框--------------------
typeBlod = tk.IntVar()
typeItalic = tk.IntVar()
tk.Checkbutton(top, text = "Blod", variable = typeBlod, onvalue = 1, offvalue = 0, command = typeChecked).pack(side = tk.LEFT)
tk.Checkbutton(top, text = "Italic", variable = typeItalic, onvalue = 2, offvalue = 0, command = typeChecked).pack(side = tk.LEFT)
'''
canvas_7=tk.Canvas(tab7,height=585,width=580)
canvas_7.pack()
canvas_7.create_rectangle(0,0,580,585,fill="skyblue")
canvas_7.create_text(290,35,text="数据导出",fill='white',font=('Comic Sans MS',25))
canvas_7.create_rectangle(0,75,580,78,fill="white")
canvas_7.create_rectangle(0,251,580,254,fill="white")
canvas_7.create_rectangle(0,428,580,431,fill="white")
##时间类型下拉框
canvas_7.create_text(50,111,text="日期类型：",fill='black',font=('Comic Sans MS',10))
datetype=tk.StringVar()
datetype_chose=ttk.Combobox(tab7,width=12,textvariable=datetype)
datetype_chose['values'] = ('请选择','安装时间','制单日期')
datetype_chose.place(x=90,y=100)
###只读
datetype_chose["state"] = "readonly"
datetype_chose.current(0)
##年下拉框
canvas_7.create_text(220,111,text="S年:",fill='black',font=('Comic Sans MS',10))
syear=tk.StringVar()
syear_chose=ttk.Combobox(tab7,width=8,textvariable=syear)
syear_chose['values'] = ('开始年份S','2010','2011','2012','2013','2014','2015','2016','2017','2018','2019','2020')
syear_chose.place(x=240,y=100)
syear_chose["state"] = "readonly"
syear_chose.current(0)
##月下拉框
canvas_7.create_text(345,111,text="S月:",fill='black',font=('Comic Sans MS',10))
smonth=tk.StringVar()
smonth_chose=ttk.Combobox(tab7,width=8,textvariable=smonth)
smonth_chose['values'] = ('开始月份','01','02','03','04','05','06','07','08','09','10','11','12')
smonth_chose.place(x=365,y=100)
smonth_chose["state"] = "readonly"
smonth_chose.current(0)
##日下拉框
canvas_7.create_text(465,111,text="S日:",fill='black',font=('Comic Sans MS',10))
sday=tk.StringVar()
sday_chose=ttk.Combobox(tab7,width=8,textvariable=sday)
sday_chose['values'] = ('开始日','01','02','03','04','05','06','07','08','09','10','11','12','13','14',
						'15','16','17','18','19','20','21','22','23','24','25','26','27','28','29','30','31')
sday_chose.place(x=485,y=100)
sday_chose["state"] = "readonly"
sday_chose.current(0)

##状态下拉框
canvas_7.create_text(50,151,text="单据状态：",fill='black',font=('Comic Sans MS',10))
billstate=tk.StringVar()
billstate_chose=ttk.Combobox(tab7,width=12,textvariable=billstate)
billstate_chose['values'] = ('请选择','已完成','所有状态')
billstate_chose.place(x=90,y=140)
###只读
billstate_chose["state"] = "readonly"
billstate_chose.current(0)
##年下拉框
canvas_7.create_text(220,151,text="E年:",fill='black',font=('Comic Sans MS',10))
eyear=tk.StringVar()
eyear_chose=ttk.Combobox(tab7,width=8,textvariable=eyear)
eyear_chose['values'] = ('结束年份E','2010','2011','2012','2013','2014','2015','2016','2017','2018','2019','2020')
eyear_chose.place(x=240,y=140)
eyear_chose["state"] = "readonly"
eyear_chose.current(0)
##月下拉框
canvas_7.create_text(345,151,text="E月:",fill='black',font=('Comic Sans MS',10))
emonth=tk.StringVar()
emonth_chose=ttk.Combobox(tab7,width=8,textvariable=emonth)
emonth_chose['values'] = ('结束月份','01','02','03','04','05','06','07','08','09','10','11','12')
emonth_chose.place(x=365,y=140)
emonth_chose["state"] = "readonly"
emonth_chose.current(0)
##日下拉框
canvas_7.create_text(465,151,text="E日:",fill='black',font=('Comic Sans MS',10))
eday=tk.StringVar()
eday_chose=ttk.Combobox(tab7,width=8,textvariable=eday)
eday_chose['values'] = ('结束日','01','02','03','04','05','06','07','08','09','10','11','12','13','14',
						'15','16','17','18','19','20','21','22','23','24','25','26','27','28','29','30','31')
eday_chose.place(x=485,y=140)
eday_chose["state"] = "readonly"
eday_chose.current(0)

canvas_7.create_text(90,190,text="区域代码",fill='black',font=('Comic Sans MS',10))
areacode=tk.Entry(tab7)
areacode.place(x=120,y=180)
canvas_7.create_text(290,190,text="机器P码",fill='black',font=('Comic Sans MS',10))
machine=tk.Entry(tab7)
machine.place(x=320,y=180)

b_7=tk.Button(tab7,text="开户单导出",activebackground="gold",fg='black',bg='beige',command=dump_applynew)
b_7.place(x=60,y=210)
b_8=tk.Button(tab7,text="退机单导出",activebackground="gold",fg='black',bg='beige',command=dump_exchange)
b_8.place(x=175,y=210)
b_9=tk.Button(tab7,text="换机单导出",activebackground="gold",fg='black',bg='beige',command=dump_exchange)
b_9.place(x=290,y=210)
b_10=tk.Button(tab7,text="维修单导出",activebackground="gold",fg='black',bg='beige',command=dump_exchange)
b_10.place(x=405,y=210)

canvas_7.create_text(90,290,text="案例城市",fill='black',font=('Comic Sans Ms',10))
city=tk.Entry(tab7)
city.place(x=120,y=280)

canvas_7.create_text(290,290,text="案例名称",fill='black',font=('Comic Sans Ms',10))
customer=tk.Entry(tab7)
customer.place(x=320,y=280)

bill_area=tk.StringVar()
bill_area_chose=ttk.Combobox(tab7,width=8,textvariable=bill_area)
bill_area_chose['values'] = ('省市区？','省','市')
bill_area_chose.place(x=490,y=280)
bill_area_chose["state"] = "readonly"
bill_area_chose.current(0)

b_11=tk.Button(tab7,text="装机案例",activebackground="gold",fg='black',bg='beige',command=dump_example)
b_11.place(x=260,y=310)

canvas_7.create_text(290,460,text="Tips:多个代码机型(注意是P码)请用英文逗号分开,",fill='white',font=('Comic Sans MS',15))
canvas_7.create_text(290,490,text="为便于输入，案例城市，案例名称请用中文逗号分开",fill='white',font=('Comic Sans MS',15))
canvas_7.create_text(290,520,text="出于信息安全，数据量限制5000条，请不要超过此数量",fill='white',font=('Comic Sans MS',15))
#####################################################异常数据第一版###############################################
# canvas_8=tk.Canvas(tab8,height=585,width=580)
# canvas_8.pack()
# canvas_8.create_rectangle(0,0,580,385,fill="Khaki")
# canvas_8.create_rectangle(0,385,580,585,fill="DarkKhaki")
# canvas_8.create_text(290,35,text="Bad Order",fill='White',font=('Comic Sans MS',25))
# tk.Button(tab8,text="异常单据",activebackground="skyblue",bg="Orange",fg="white").place(x=520,y=55)
# t_8=tk.Text(tab8,height=20,width=70)
# t_8.place(x=50,y=90)
# entry_8=tk.Entry(tab8,font=('', 15))
# entry_8.place(x=50,y=400)
# model3=tk.StringVar()
# tk.Radiobutton(tab8,text='浩泽单据',variable=model3,value='a',command=model_select_2,font=('Times', 10),bg='DarkKhaki').place(x=50,y=440)
# tk.Radiobutton(tab8,text='浩优单据',variable=model3,value='b',command=model_select_2,font=('Times', 10),bg='DarkKhaki').place(x=175,y=440)
# tk.Button(tab8,text="查看异常原因",activebackground="skyblue",bg="Orange",fg="white").place(x=110,y=480)
# tk.Button(tab8,text="生成解决方案",activebackground="skyblue",bg="Orange",fg="white").place(x=110,y=530)
# points = [100, 140, 110, 110, 140, 100, 110, 90, 100, 60, 90, 90, 60, 100, 90, 110]
# points_8 = [470, 140, 480, 110, 510, 100, 480, 90, 470, 60, 460, 90, 430, 100, 460, 110]
# canvas_8.create_polygon(points, outline="#476042",
# 			fill='yellow', width=3)
# canvas_8.create_polygon(points_8, outline="#476042",
# 			fill='yellow', width=3)
#####################################################异常数据第二版###############################################
canvas_8=tk.Canvas(tab8,height=585,width=580)
canvas_8.pack()
dmimg=tk.PhotoImage(file="D:/CRMTOOL/model/orderDM.png")
canvas_8.create_image(0,0,anchor=tk.NW,image=dmimg)
canvas_8.create_text(290,35,text="数据监控",fill="#F4FFEC",font=('Verdana',25))
##画横线
canvas_8.create_rectangle(30,75,550,78,fill="white")
canvas_8.create_rectangle(30,120,550,122,fill="white")
canvas_8.create_rectangle(30,165,550,167,fill="white")
canvas_8.create_rectangle(30,210,550,212,fill="white")
canvas_8.create_rectangle(30,255,550,257,fill="white")
canvas_8.create_rectangle(30,300,550,302,fill="white")
canvas_8.create_rectangle(30,345,550,348,fill="white")
##画纵线
canvas_8.create_rectangle(150,78,152,348,fill="white")
canvas_8.create_rectangle(283,78,285,348,fill="white")
canvas_8.create_rectangle(416,78,418,348,fill="white")
#画首行
canvas_8.create_text(92,97.5,text="单据编号",fill="white",font=("Verdana",15))
canvas_8.create_text(216,97.5,text="单据类型",fill="white",font=("Verdana",15))
canvas_8.create_text(349,97.5,text="单据状态",fill="white",font=("Verdana",15))
canvas_8.create_text(482,97.5,text="超时时间",fill="white",font=("Verdana",15))
# tk.Button(tab8,text='查看刷新',fg='white',bg='#03193B',command=output_bad_order).place(x=120,y=360)
# tk.Button(tab8,text='异常原因',fg='white',bg='#03193B',command=output_bad_order).place(x=253,y=360)
# tk.Button(tab8,text='一键解决',fg='white',bg='#03193B',command=solution_way).place(x=386,y=360)
tk.Button(tab8,text='点击查看',fg='white',bg='#03193B',command=output_bad_order).place(x=60,y=360)
tk.Button(tab8,text='异常原因',fg='white',bg='#03193B',command=select_reason).place(x=193,y=360)
tk.Button(tab8,text='一键解决',fg='white',bg='#03193B',command=solution_way).place(x=336,y=360)
######################################################每日运势###################################################

###########关于###########

###########菜单实现###########
menuBar=tk.Menu(top)
top.config(menu=menuBar)
#####增加主题
##第一个菜单
fileMenu=tk.Menu(menuBar,tearoff=0)
fileMenu.add_command(label="新建")
#菜单选项的下划线
fileMenu.add_separator()
fileMenu.add_command(label="操作日志",command=openfile_log)
fileMenu.add_separator()
fileMenu.add_command(label="sql日志",command=openfile_sqllog)
fileMenu.add_command(label="生成日志路径")
fileMenu.add_separator()
fileMenu.add_command(label="连接数据库测试",command=conn_test)
fileMenu.add_separator()
fileMenu.add_command(label="退出程序",command=_quit)
fileMenu.add_separator()
fileMenu.add_command(label="重新启动",command=restart)
menuBar.add_cascade(label="系统",menu=fileMenu)
##第二个菜单
fileMenu_s=tk.Menu(menuBar,tearoff=0)
fileMenu_s.add_command(label="导入")
fileMenu_s.add_separator()
fileMenu_s.add_command(label="没想好")
fileMenu_s.add_separator()
fileMenu_s.add_command(label="待确定")
menuBar.add_cascade(label="设置",menu=fileMenu_s)
##第三个菜单
fileMenu_s=tk.Menu(menuBar,tearoff=0)
fileMenu_s.add_command(label="中文")
fileMenu_s.add_separator()
fileMenu_s.add_command(label="汉语")
fileMenu_s.add_separator()
fileMenu_s.add_command(label="Chinese")
menuBar.add_cascade(label="语言",menu=fileMenu_s)
##最后个菜单
fileMenu_f=tk.Menu(menuBar,tearoff=0)
fileMenu_f.add_command(label="系统版本",command=msgBox1)
fileMenu_f.add_separator()
fileMenu_f.add_command(label="意见反馈",command=msgBox4)
fileMenu_f.add_separator()
fileMenu_f.add_command(label="淘宝链接",command=taobao_link)
fileMenu_f.add_separator()
fileMenu_f.add_command(label="更多工具",command=tools_link)
fileMenu_f.add_separator()
fileMenu_f.add_command(label="开发选项")
menuBar.add_cascade(label="关于",menu=fileMenu_f)


'''
win = tk.Tk()     
# Add a title         
win.title("Python 图形用户界面")  
menuBar = tkinter.Menu(win)  
win.config(menu=menuBar)  

# Add menu items  
fileMenu = tkinter.Menu(menuBar, tearoff=0)  
fileMenu.add_command(label="新建")  
fileMenu.add_separator()  
fileMenu.add_command(label="退出")  
menuBar.add_cascade(label="文件", menu=fileMenu)

msgMenu = Menu(menuBar, tearoff=0)  
msgMenu.add_command(label="通知 Box", command=_msgBox1)  
msgMenu.add_command(label="警告 Box", command=_msgBox2)  
msgMenu.add_command(label="错误 Box", command=_msgBox3)  
msgMenu.add_separator()  
msgMenu.add_command(label="判断对话框", command=_msgBox4)  
menuBar.add_cascade(label="消息框", menu=msgMenu)  
win.mainloop()
'''

top.mainloop()
