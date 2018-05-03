import re
import os
import pyodbc
import xlsxwriter
class Dataconn():
    def __init__(self,server,database,uid,pwd):
        self.conn=pyodbc.connect(r'DRIVER={SQL Server};SERVER='+server

+';DATABASE='+database+';UID='+uid+';PWD='+pwd)
        self.cur=self.conn.cursor()

    def close_conn(self):
        self.cur.close()
        self.conn.close()


#得到初始化基础数据
def get_dict(init_sql):
    hzdb=Dataconn('localhost','KFT_MAIN','admin','admin')
    hzdb.cur.execute(init_sql)
    init_res=hzdb.cur.fetchall()
    dict={}
    for data in range(len(init_res)-1):
        dict[init_res[data][0]]=init_res[data][1]
    hzdb.close_conn()
    return dict

def get_res(data_sql):
    hydb=Dataconn('localhost','S60_MAIN','admin','admin')
    hydb.cur.execute(data_sql)
    res=hydb.cur.fetchall()
    # print(len(res))
    # for i in range(len(res)-1):
    #     res[i]=list(res[i])
        # keylist=

['cdate','customerNo','billno','name','tel','mobile','province','city','area','address'

,'remark','createdate','username','finshdatetime','A8']
        # order=dict(zip(keylist,res[i]))
        # res_li.append(order)
    # print(res)
    hydb.close_conn()
    return res

def get_card(num):
    agent_db=Dataconn('localhost','AgentSystem','admin','admin')
    card_sql="select top "+str(num)+ " cardcode,password from HM_AccountCards where 

AreaCode='999021' and IsOpen<>1 and IsCommited<>1"
    # print(card_sql)
    agent_db.cur.execute(card_sql)
    cards=agent_db.cur.fetchall()
    str_card = ""
    for i in range(len(cards)):
        str_card = str_card + "'" + cards[i][0] + "',"

    update_sql = "update HM_AccountCards set 

IsOpen=1,IsCommited=1,OpenTimes=1,OpenDate=getdate() where CardCode in " + "(" + 

str_card[0:len(str_card) - 1] + ")"
    #更改卡状态
    agent_db.cur.execute(update_sql)
    agent_db.close_conn()
    return cards


def create_order(dict,res):
    li=[]
    customer_dict={}
    old_cusno_1=172509126
    old_cusno_2=181012807
    i=0
    card = get_card(len(res))
    print(len(res))
    for j in res:
        j=list(j)
        if j[0] in list(dict.keys()):
            #找到对应的日期，在基数后面加1
            dict[j[0]]=int(dict[j[0]])+1
            # j=list(j)
            #构造开户单
            if len(str(dict[j[0]]))==3:
                kh_billno = 'KH' + j[2][2:8] + '-' +'0'+ str(dict[j[0]])
            else:
                kh_billno='KH'+j[2][2:8]+'-'+str(dict[j[0]])
            j.append(kh_billno)

            # 构造客户编号
            if j[0][0:4] == '2017':
                #若之前出现过同样的客户编号，则使用之前的客户编号
                if j[1] in customer_dict.keys():
                    j.append(customer_dict[j[1]])
                else:
                    j.append('C' + str(old_cusno_1))
                    customer_dict[j[1]] = j[21]
                    old_cusno_1 += 1
            else:
                if j[1] in customer_dict.keys():
                    j.append(customer_dict[j[1]])
                else:
                    j.append('C' + str(old_cusno_2))
                    customer_dict[j[1]] = j[21]
                    old_cusno_2 += 1

            #构造开户卡、密码
            j.append(card[i][0])
            j.append(card[i][1])

            #构造其它字段
            # ['cdate', 'customerNo', 'billno', 'name', 'tel', 'mobile', 'province', 

'city', 'area', 'address', 'remark', 'createdate', 'username', 'finshdatetime', 'A8']
            j=['',j[20],1001,'',j[22],j[23],'P507441','JZY-A2B-X(T1)深圳伊泉',0,j

[21],j[3],j[3],j[4],j[5],j[6],j[7],j[8],j[9],
               '',j[10],'WD0053','CK0014',0,21410,j[12],j

[12],0,'contractcode','effectivedate','',0.00,0.00,0.00,
               0,'','',j[16],j[17],j[18],'','',j[13],j[13],'','','','','','','',j

[15],210,0,20974,j[11],j[11],
               20974,j[11],'999021',0,'','','',0,'','',0,0,0,'','',0,j[19],0,0,0,'正

常','',0,'',1,0,'','','','','深圳伊泉净品','',
               4521,j[14],'',20700,0,'',0,'',''
               ]
            li.append(j)
        i+=1
    return li


def dump_execl(list):
    os.chdir('C:/users/administrator/desktop')
    workbook=xlsxwriter.Workbook('yq-applynew.xlsx')
    worksheet = workbook.add_worksheet("applynew")
    worksheet.write(0, 1, 'ID')
    worksheet.write(0, 2, 'BillNo')
    worksheet.write(0, 3, 'BillClass')
    worksheet.write(0, 4, 'CustomCode')
    worksheet.write(0, 5, 'CardCode')
    worksheet.write(0, 6, 'CardPwd')
    worksheet.write(0, 7, 'MachineSID')
    worksheet.write(0, 8, 'MachineName')
    worksheet.write(0, 9, 'IsTrial')
    worksheet.write(0, 10, 'CustomerNo')
    worksheet.write(0, 11, 'Name')
    worksheet.write(0, 12, 'Contact')
    worksheet.write(0, 13, 'Tel')
    worksheet.write(0, 14, 'Mobile')
    worksheet.write(0, 15, 'Province')
    worksheet.write(0, 16, 'City')
    worksheet.write(0, 17, 'Area')
    worksheet.write(0, 18, 'Address')
    worksheet.write(0, 19, 'PayMethod')
    worksheet.write(0, 20, 'Remark')
    worksheet.write(0, 21, 'BranchNo')
    worksheet.write(0, 22, 'StoreHouseNo')
    worksheet.write(0, 23, 'Seller')
    worksheet.write(0, 24, 'Agent')
    worksheet.write(0, 25, 'Installer')
    worksheet.write(0, 26, 'Repairer')
    worksheet.write(0, 27, 'Filler')
    worksheet.write(0, 28, 'ContractCode')
    worksheet.write(0, 29, 'EffectiveDate')
    worksheet.write(0, 30, 'RemindState')
    worksheet.write(0, 31, 'ApplyFee')
    worksheet.write(0, 32, 'Total')
    worksheet.write(0, 33, 'FollowPrice')
    worksheet.write(0, 34, 'IsContract')
    worksheet.write(0, 35, 'MachineModel')
    worksheet.write(0, 36, 'MachineCode')
    worksheet.write(0, 37, 'Y_TDS')
    worksheet.write(0, 38, 'Z_TDS')
    worksheet.write(0, 39, 'ShuiYa')
    worksheet.write(0, 40, 'Satisfaction')
    worksheet.write(0, 41, 'Dissatisfied')
    worksheet.write(0, 42, 'InstalDate')
    worksheet.write(0, 43, 'DeliveryDate')
    worksheet.write(0, 44, 'Delivery')
    worksheet.write(0, 45, 'DeliveryTel')
    worksheet.write(0, 46, 'DeliveryInfo')
    worksheet.write(0, 47, 'DeliveryCost')
    worksheet.write(0, 48, 'Paid')
    worksheet.write(0, 49, 'Cashier')
    worksheet.write(0, 50, 'PaidDate')
    worksheet.write(0, 51, 'BillState')
    worksheet.write(0, 52, 'PayState')
    worksheet.write(0, 53, 'IsPaid')
    worksheet.write(0, 54, 'Creator')
    worksheet.write(0, 55, 'CreateDate')
    worksheet.write(0, 56, 'ActivaDate')
    worksheet.write(0, 57, 'Updater')
    worksheet.write(0, 58, 'UpdateDate')
    worksheet.write(0, 59, 'AreaCode')
    worksheet.write(0, 60, 'IsDel')
    worksheet.write(0, 61, 'oldSeller')
    worksheet.write(0, 62, 'oldRepairer')
    worksheet.write(0, 63, 'oldCreator')
    worksheet.write(0, 64, 'IsRemind')
    worksheet.write(0, 65, 'Reminder')
    worksheet.write(0, 66, 'RemindDate')
    worksheet.write(0, 67, 'IsSpecial')
    worksheet.write(0, 68, 'IsHousehold')
    worksheet.write(0, 69, 'IsConfirm')
    worksheet.write(0, 70, 'Agreement')
    worksheet.write(0, 71, 'Source')
    worksheet.write(0, 72, 'CardSeller')
    worksheet.write(0, 73, 'Extend1')
    worksheet.write(0, 74, 'IsEndDeliver')
    worksheet.write(0, 75, 'IsKft')
    worksheet.write(0, 76, 'MachineVersion')
    worksheet.write(0, 77, 'BillSort')
    worksheet.write(0, 78, 'OpenAudit')
    worksheet.write(0, 79, 'FillInterval')
    worksheet.write(0, 80, 'Type')
    worksheet.write(0, 81, 'A1')
    worksheet.write(0, 82, 'ContractRecycle')
    worksheet.write(0, 83, 'InstallRemark')
    worksheet.write(0, 84, 'A2')
    worksheet.write(0, 85, 'A3')
    worksheet.write(0, 86, 'A4')
    worksheet.write(0, 87, 'A5')
    worksheet.write(0, 88, 'A6')
    worksheet.write(0, 89, 'A7')
    worksheet.write(0, 90, 'A8')
    worksheet.write(0, 91, 'A9')
    worksheet.write(0, 92, 'JDState')
    worksheet.write(0, 93, 'B1')
    worksheet.write(0, 94, 'B2')
    worksheet.write(0, 95, 'B0')
    worksheet.write(0, 96, 'WLYDH')
    worksheet.write(0, 97, 'CKBH')
    row=1
    for i in list:
        worksheet.write(row, 1, i[0])
        worksheet.write(row, 2, i[1])
        worksheet.write(row, 3, i[2])
        worksheet.write(row, 4, i[3])
        worksheet.write(row, 5, i[4])
        worksheet.write(row, 6, i[5])
        worksheet.write(row, 7, i[6])
        worksheet.write(row, 8, i[7])
        worksheet.write(row, 9, i[8])
        worksheet.write(row, 10, i[9])
        worksheet.write(row, 11, i[10])
        worksheet.write(row, 12, i[11])
        worksheet.write(row, 13, i[12])
        worksheet.write(row, 14, i[13])
        worksheet.write(row, 15, i[14])
        worksheet.write(row, 16, i[15])
        worksheet.write(row, 17, i[16])
        worksheet.write(row, 18, i[17])
        worksheet.write(row, 19, i[18])
        worksheet.write(row, 20, i[19])
        worksheet.write(row, 21, i[20])
        worksheet.write(row, 22, i[21])
        worksheet.write(row, 23, i[22])
        worksheet.write(row, 24, i[23])
        worksheet.write(row, 25, i[24])
        worksheet.write(row, 26, i[25])
        worksheet.write(row, 27, i[26])
        worksheet.write(row, 28, i[27])
        worksheet.write(row, 29, i[28])
        worksheet.write(row, 30, i[29])
        worksheet.write(row, 31, i[30])
        worksheet.write(row, 32, i[31])
        worksheet.write(row, 33, i[32])
        worksheet.write(row, 34, i[33])
        worksheet.write(row, 35, i[34])
        worksheet.write(row, 36, i[35])
        worksheet.write(row, 37, i[36])
        worksheet.write(row, 38, i[37])
        worksheet.write(row, 39, i[38])
        worksheet.write(row, 40, i[39])
        worksheet.write(row, 41, i[40])
        worksheet.write(row, 42, i[41])
        worksheet.write(row, 43, i[42])
        worksheet.write(row, 44, i[43])
        worksheet.write(row, 45, i[44])
        worksheet.write(row, 46, i[45])
        worksheet.write(row, 47, i[46])
        worksheet.write(row, 48, i[47])
        worksheet.write(row, 49, i[48])
        worksheet.write(row, 50, i[49])
        worksheet.write(row, 51, i[50])
        worksheet.write(row, 52, i[51])
        worksheet.write(row, 53, i[52])
        worksheet.write(row, 54, i[53])
        worksheet.write(row, 55, i[54])
        worksheet.write(row, 56, i[55])
        worksheet.write(row, 57, i[56])
        worksheet.write(row, 58, i[57])
        worksheet.write(row, 59, i[58])
        worksheet.write(row, 60, i[59])
        worksheet.write(row, 61, i[60])
        worksheet.write(row, 62, i[61])
        worksheet.write(row, 63, i[62])
        worksheet.write(row, 64, i[63])
        worksheet.write(row, 65, i[64])
        worksheet.write(row, 66, i[65])
        worksheet.write(row, 67, i[66])
        worksheet.write(row, 68, i[67])
        worksheet.write(row, 69, i[68])
        worksheet.write(row, 70, i[69])
        worksheet.write(row, 71, i[70])
        worksheet.write(row, 72, i[71])
        worksheet.write(row, 73, i[72])
        worksheet.write(row, 74, i[73])
        worksheet.write(row, 75, i[74])
        worksheet.write(row, 76, i[75])
        worksheet.write(row, 77, i[76])
        worksheet.write(row, 78, i[77])
        worksheet.write(row, 79, i[78])
        worksheet.write(row, 80, i[79])
        worksheet.write(row, 81, i[80])
        worksheet.write(row, 82, i[81])
        worksheet.write(row, 83, i[82])
        worksheet.write(row, 84, i[83])
        worksheet.write(row, 85, i[84])
        worksheet.write(row, 86, i[85])
        worksheet.write(row, 87, i[86])
        worksheet.write(row, 88, i[87])
        worksheet.write(row, 89, i[88])
        worksheet.write(row, 90, i[89])
        worksheet.write(row, 91, i[90])
        worksheet.write(row, 92, i[91])
        worksheet.write(row, 93, i[92])
        worksheet.write(row, 94, i[93])
        worksheet.write(row, 95, i[94])
        worksheet.write(row, 96, i[95])
        worksheet.write(row, 97, i[96])
        row+=1

def get_wxorder(base_dict,res,map):
    li=[]
    for j in res:
        j=list(j)
        if j[0] in list(base_dict.keys()):
            #找到对应的日期，在基数后面加1
            base_dict[j[0]]=int(base_dict[j[0]])+1
            #构造单号
            wx_billno='WX'+j[2][2:8]+'-'+str(base_dict[j[0]])
            j.append(wx_billno)
            #根据客户编号映射开户单号
            if j[1] in list(map.keys()):
                j.append(map[j[1]])
            else:
                j.append('NO MAPPING')
            j=[j[1],j[11],1003,'',j[12],'P507441','customerno',j[3],j

[3],'tel','province','city','area','address',j[5],j[5],j[10][3:6],j[10][0:3],j[4],
               'branch','store',j[6],'','','','','','',0.00,'',j[9],20974,j[5],20974,j

[5],0,0,0,'正常','','','','','','','','','','','']
            li.append(j)
    return li

def dump_wxorder(res):
    os.chdir('C:/users/administrator/desktop')
    workbook=xlsxwriter.Workbook('yq-repair.xlsx')
    worksheet = workbook.add_worksheet("repair")
    worksheet.write(0, 0, 'ID')
    worksheet.write(0, 1, 'BillNo')
    worksheet.write(0, 2, 'BillClass')
    worksheet.write(0, 3, 'CustomCode')
    worksheet.write(0, 4, 'ApplyNewNo')
    worksheet.write(0, 5, 'MachineSID')
    worksheet.write(0, 6, 'CustomerNo')
    worksheet.write(0, 7, 'Name')
    worksheet.write(0, 8, 'Contact')
    worksheet.write(0, 9, 'Tel')
    worksheet.write(0, 10, 'Province')
    worksheet.write(0, 11, 'City')
    worksheet.write(0, 12, 'Area')
    worksheet.write(0, 13, 'Address')
    worksheet.write(0, 14, 'ReportDate')
    worksheet.write(0, 15, 'PeriodDate')
    worksheet.write(0, 16, 'RepairType')
    worksheet.write(0, 17, 'RepairClass')
    worksheet.write(0, 18, 'Remark')
    worksheet.write(0, 19, 'BranchNo')
    worksheet.write(0, 20, 'StoreHouseNo')
    worksheet.write(0, 21, 'Repairer')
    worksheet.write(0, 22, 'RepairDate')
    worksheet.write(0, 23, 'RepairLevel')
    worksheet.write(0, 24, 'ResidualSZ')
    worksheet.write(0, 25, 'TDS')
    worksheet.write(0, 26, 'Fault')
    worksheet.write(0, 27, 'Solution')
    worksheet.write(0, 28, 'RepairCost')
    worksheet.write(0, 29, 'DistanceWay')
    worksheet.write(0, 30, 'BillState')
    worksheet.write(0, 31, 'Creator')
    worksheet.write(0, 32, 'CreateDate')
    worksheet.write(0, 33, 'Updater')
    worksheet.write(0, 34, 'UpdateDate')
    worksheet.write(0, 35, 'IsDel')
    worksheet.write(0, 36, 'IsSpecial')
    worksheet.write(0, 37, 'MachineVersion')
    worksheet.write(0, 38, 'BillSort')
    worksheet.write(0, 39, 'A0')
    worksheet.write(0, 40, 'A1')
    worksheet.write(0, 41, 'A2')
    worksheet.write(0, 42, 'A3')
    worksheet.write(0, 43, 'A4')
    worksheet.write(0, 44, 'A5')
    worksheet.write(0, 45, 'A6')
    worksheet.write(0, 46, 'A7')
    worksheet.write(0, 47, 'A8')
    worksheet.write(0, 48, 'A9')
    worksheet.write(0, 49, 'JDState')
    row=1
    for i in res:
        worksheet.write(row, 0, i[0])
        worksheet.write(row, 1, i[1])
        worksheet.write(row, 2, i[2])
        worksheet.write(row, 3, i[3])
        worksheet.write(row, 4, i[4])
        worksheet.write(row, 5, i[5])
        worksheet.write(row, 6, i[6])
        worksheet.write(row, 7, i[7])
        worksheet.write(row, 8, i[8])
        worksheet.write(row, 9, i[9])
        worksheet.write(row, 10, i[10])
        worksheet.write(row, 11, i[11])
        worksheet.write(row, 12, i[12])
        worksheet.write(row, 13, i[13])
        worksheet.write(row, 14, i[14])
        worksheet.write(row, 15, i[15])
        worksheet.write(row, 16, i[16])
        worksheet.write(row, 17, i[17])
        worksheet.write(row, 18, i[18])
        worksheet.write(row, 19, i[19])
        worksheet.write(row, 20, i[20])
        worksheet.write(row, 21, i[21])
        worksheet.write(row, 22, i[22])
        worksheet.write(row, 23, i[23])
        worksheet.write(row, 24, i[24])
        worksheet.write(row, 25, i[25])
        worksheet.write(row, 26, i[26])
        worksheet.write(row, 27, i[27])
        worksheet.write(row, 28, i[28])
        worksheet.write(row, 29, i[29])
        worksheet.write(row, 30, i[30])
        worksheet.write(row, 31, i[31])
        worksheet.write(row, 32, i[32])
        worksheet.write(row, 33, i[33])
        worksheet.write(row, 34, i[34])
        worksheet.write(row, 35, i[35])
        worksheet.write(row, 36, i[36])
        worksheet.write(row, 37, i[37])
        worksheet.write(row, 38, i[38])
        worksheet.write(row, 39, i[39])
        worksheet.write(row, 40, i[40])
        worksheet.write(row, 41, i[41])
        worksheet.write(row, 42, i[42])
        worksheet.write(row, 43, i[43])
        worksheet.write(row, 44, i[44])
        worksheet.write(row, 45, i[45])
        worksheet.write(row, 46, i[46])
        worksheet.write(row, 47, i[47])
        worksheet.write(row, 48, i[48])
        worksheet.write(row, 49, i[49])
        row+=1


if __name__=="__main__":
    ###开户单
    kh_sql='''select convert(varchar(10),a.createdate,120) as cdate,right(max

(billno),4) as num,SUBSTRING(max(billno),0,9) as head from kft_applynew  a
            where CreateDate>='2017-9-30' group by convert(varchar(10),createdate,120)
            order by convert(varchar(10),createdate,120)
            '''
    data_sql='''select convert(varchar(10),a.createdate,120) as 

cdate,CustomerNo,billno,a.Name,yq.mobile_MXexport,yq.mobile_MXexport,
				b.province,c.city,d.area,yq.address_MXexport,convert

(varchar(20),getdate(),120)+':补录订单.'+isnull(yq.Remark,''),convert(varchar

(30),a.createdate,121) as createdate,
				u.UserName,convert(varchar(30),a.FinshDateTime,121) as 

FinshDateTime,a.A8, 
            yq.billstate as bstate,
            Y_TDS,Z_TDS,ShuiYa,yq.device_id
            from MX_Repair a join yq_order yq on a.A8=yq.serialno
            left join MX_Province b on a.Province=b.provinceID
            left join MX_city c on a.city=c.cityID
            left join MX_area d on a.area=d.areaID
            left join MX_User u on a.Engineer=u.ID
            --where a.CompanyID=100 and BillState >= 130000006 
            --and serviceitem=123000006
            order by a.CreateDate'''
    kh_base=get_dict(kh_sql)
    res=get_res(data_sql)
    applynew_order=create_order(kh_base,res)
    # 输出execl
    dump_execl(applynew_order)
    print('开户单转换任务完成！')

    ###############退换
    wx_sql='''select convert(varchar(10),a.createdate,120) as cdate,right(max

(billno),4) as num,SUBSTRING(max(billno),0,9) as head from kft_repair a
            where CreateDate>='2017-9-30' group by convert(varchar(10),createdate,120)
            order by convert(varchar(10),createdate,120)
            '''
    wx_data_sql='''select convert(varchar(10),a.createdate,120) as 

cdate,CustomerNo,billno,a.Name,Remark,
        convert(varchar(30),a.createdate,121) as createdate,u.UserName,convert(varchar

(30),a.FinshDateTime,121) as FinshDateTime,a.A8, 
            case a.billstate when 130000002 then 100 when 130000006 then 120 when 

130000008 then 150 else 99 end as bstate,
            case serviceitem when 123000008 then '414892' else '896843' end as fuwuitem
            from MX_Repair a
            left join MX_Province b on a.Province=b.provinceID
            left join MX_city c on a.city=c.cityID
            left join MX_area d on a.area=d.areaID
            left join MX_User u on a.Engineer=u.ID
            where a.CompanyID=100 and BillState >= 130000006 
            and serviceitem in(123000007,123000008)
            order by a.CreateDate
        '''

    map_apply={}  #得到维修开户单映射
    for i in applynew_order:
        map_apply[i[0]]=i[1]  #{'C172712853':'KH171101-1101','C172736566':'KH171101-1102'}

    wx_base=get_dict(wx_sql) #得到一个维修基数对照字典
    wx_res=get_res(wx_data_sql)#得到维修数据
    repair_order=get_wxorder(wx_base,wx_res,map_apply)
    dump_wxorder(repair_order)
    print('维修单转换任务完成！')






