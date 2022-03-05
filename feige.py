# -*- coding: utf-8 -*-
"""
Created on Fri Feb 18 17:59:14 2022

@author: Administrator

#导出飞鸽数据
"""
import requests,time,datetime,xlrd,glob,os
from datetime import date, timedelta



Yesterday = (date.today() + timedelta(days=-1)).strftime("%Y-%m-%d")
LastWeek = (date.today() + timedelta(days=-7)).strftime("%Y-%m-%d")

year = datetime.datetime.now().year
month = datetime.datetime.now().month
nowday = datetime.datetime.now().day
if nowday == 1:
    month = month-1
first_day = str(year)+'-'+str(month)+'-'+'1' + ' 00:00:00'
first_day = time.strptime(first_day, "%Y-%m-%d %H:%M:%S")
first_day = int(time.mktime(first_day))*1000
yesterday = Yesterday + ' 23:59:59'
yesterday = time.strptime(yesterday, "%Y-%m-%d %H:%M:%S")
yesterday = int(time.mktime(yesterday))*1000+999
Nowtime = time.strftime("%Y-%m-%d_%H_%M_%S", time.localtime())
MonthStart = str(year)+'-'+str(month).rjust(2,'0')+'-'+ '01' #本月1号


#创建文件夹
def mkdir(path):
    if not os.path.exists(path):
        os.makedirs(path) 

#获取Cookie
def read_cookie_file():
    CookieList = []
    f = open('cookie.txt',encoding = 'utf-8')
    f = f.readlines()
    for i in f:
        if len(i) >1000:
            CookieList.append(i)
    return CookieList

def list_to_string(row):
    text_row = str(row).replace('[', '')
    text_row = str(text_row).replace(']', '')
    text_row = str(text_row).replace("'", '')
    text_row = text_row.replace(' ', '')
    return text_row



#获取昨日店铺数据
def get_yesterday():
    print('正在保存店铺数据-历史数据-近1天')
    yesterday_url = 'https://pigeon.jinritemai.com/backstage/exportShopHistoryConversationData?endTime=%s&startTime=%s'%(Yesterday,Yesterday)
    res = requests.get(yesterday_url,headers=headers)
    with open(ShopName+'_'+Yesterday +'_' +'店铺数据(日表).xlsx','wb') as f:
        f.write(res.content)
#获取近7天店铺数据数据
def get_weekly():
    print('正在保存店铺数据-历史数据-近7天')
    weekly_url = 'https://pigeon.jinritemai.com/backstage/exportShopHistoryConversationData?endTime=%s&startTime=%s'%(Yesterday,LastWeek)
    res = requests.get(weekly_url,headers=headers)
    with open(ShopName+'_'+LastWeek+'-'+Yesterday +'_' +'店铺数据(周表).xlsx','wb') as f:
        f.write(res.content)
#获取本月店铺数据
def get_monthly():
    print('正在保存店铺数据-历史数据-1号至%s号'%(Yesterday.split('-')[-1]))
    monthly_url = 'https://pigeon.jinritemai.com/backstage/exportShopHistoryConversationData?endTime=%s&startTime=%s'%(Yesterday,MonthStart)
    res = requests.get(monthly_url,headers=headers)
    with open(ShopName+'_'+MonthStart+'-'+Yesterday +'_' +'店铺数据(月表).xlsx','wb') as f:
        f.write(res.content)

#获取昨日客服数据
def get_yesterday_kefu():
    print('正在保存客服数据-历史数据-近1天')
    strattime = yesterday-86399999
    yesterday_url = 'https://pigeon.jinritemai.com/backstage/exportHistoricalStaffStatistics?endTime=%s&queryStaffName=&sortTag=0&sortType=0&startTime=%s'%(yesterday,strattime)
    res = requests.get(yesterday_url,headers=headers)
    with open(ShopName+'_'+Yesterday +'_' +'客服数据(日表).xlsx','wb') as f:
        f.write(res.content)
# 获取近7天客服数据
def get_weekly_kefu():
    print('正在保存客服数据-历史数据-近7天')
    strattime = yesterday-604799999
    weekly_url = 'https://pigeon.jinritemai.com/backstage/exportHistoricalStaffStatistics?endTime=%s&queryStaffName=&sortTag=0&sortType=0&startTime=%s'%(yesterday,strattime)
    res = requests.get(weekly_url,headers=headers)
    with open(ShopName+'_'+LastWeek+'-'+Yesterday +'_' +'客服数据(周表).xlsx','wb') as f:
        f.write(res.content)
#获取本月客服数据
def get_monthly_kefu():
    print('正在保存客服数据-历史数据-1号至%s号'%(Yesterday.split('-')[-1]))
    monthly_url = 'https://pigeon.jinritemai.com/backstage/exportHistoricalStaffStatistics?endTime=%s&queryStaffName=&sortTag=0&sortType=0&startTime=%s'%(yesterday,first_day)
    res = requests.get(monthly_url,headers=headers)
    with open(ShopName+'_'+MonthStart+'-'+Yesterday +'_' +'客服数据(月表).xlsx','wb') as f:
        f.write(res.content)

#根据文件名获取店铺信息
def getShopName(Filename):
    if 'realme手机旗舰店' in Filename:
        shopname = '真我'
    elif '一加官方旗舰店' in Filename:
        shopname = '一加'
    elif 'OPPO手机旗舰店' in Filename:
        shopname = 'OPPO'
    elif 'OPPO商城旗舰店' in Filename:
        shopname = '欢太'
    else :
        shopname = '/'
    return shopname

def hebing():
    #读取所有.xlsx文件
    fileList = glob.glob(os.path.join('','*.xlsx'))
    #店铺数据
    yesterday_dianpu_xlxs = [i for i in fileList if '店铺数据(日表)' in i]
    weekly_dianpu_xlxs =[i for i in fileList if '店铺数据(周表)' in i]
    monthly_dianpu_xlxs =[i for i in fileList if '店铺数据(月表)' in i]
    #客服数据
    yesterday_kefu_xlxs = [i for i in fileList if '客服数据(日表)' in i]
    weekly_kefu_xlxs =[i for i in fileList if '客服数据(周表)' in i]
    monthly_kefu_xlxs =[i for i in fileList if '客服数据(月表)' in i]
    #抖音店铺数据源（店铺数据-日表合并）
    for yesterday_dianpu in yesterday_dianpu_xlxs:
        book = xlrd.open_workbook(yesterday_dianpu)
        sheet = book.sheet_by_index(0)
        #获取日表第三行内容
        shopname = getShopName(yesterday_dianpu)
        row = sheet.row_values(2) 
        row.append(shopname)
        text_row = list_to_string(row)
        with open(floderName+'\\'+'合并_店铺数据-日表.csv','a') as f :
            f.write(text_row+'\n')
    print('店铺数据-日表数据已合并！')
    #抖音店铺数据源（店铺数据-周表合并）
    for weekly_dianpu in weekly_dianpu_xlxs:
        shopname = getShopName(weekly_dianpu)
        book = xlrd.open_workbook(weekly_dianpu)
        sheet = book.sheet_by_index(0)
        row = sheet.row_values(1)
        row.append(shopname)
        text_row = list_to_string(row)
        with open(floderName+'\\'+'合并_店铺数据-周表.csv','a') as f :
            f.write(text_row+'\n')
    print('店铺数据-周表数据已合并！')        
    #抖音店铺数据源（店铺数据-月表合并）
    for monthly_dianpu in monthly_dianpu_xlxs:
        shopname = getShopName(monthly_dianpu)
        book = xlrd.open_workbook(monthly_dianpu)
        sheet = book.sheet_by_index(0)
        row = sheet.row_values(1)
        row.append(shopname)
        text_row = list_to_string(row)
        with open(floderName+'\\'+'合并_店铺数据-月表.csv','a') as f :
            f.write(text_row+'\n')
    print('店铺数据-月表数据已合并！')
    #抖音店铺数据源（客服数据-日表合并）
    for yesterday_kefu in yesterday_kefu_xlxs:
        shopname = getShopName(yesterday_kefu)
        book = xlrd.open_workbook(yesterday_kefu)
        sheet = book.sheet_by_index(0)
        #行数
        rows = sheet.nrows
        for i in range(rows):
            row = sheet.row_values(i+1)
            if row[2] =='0':
                break
            row.append(shopname)
            text_row = list_to_string(row)
            with open(floderName+'\\'+'合并_客服数据-日表.csv','a') as f :
                f.write(text_row+'\n')
    print('客服数据-日表数据已合并！')
    #客服数据周表
    for weekly_kefu in weekly_kefu_xlxs:
        shopname = getShopName(weekly_kefu)
        book = xlrd.open_workbook(weekly_kefu)
        sheet = book.sheet_by_index(0)
        #行数
        rows = sheet.nrows
        for i in range(rows):
            row = sheet.row_values(i+1)
            if row[2] =='0':
                break
            row.append(shopname)
            text_row = list_to_string(row)
            with open(floderName+'\\'+'合并_客服数据-周表.csv','a') as f :
                f.write(text_row+'\n')
    print('店铺数据-月周表数据已合并！')
    #客服数据月表
    for monthly_kefu in monthly_kefu_xlxs:
        shopname = getShopName(monthly_kefu)
        book = xlrd.open_workbook(monthly_kefu)
        sheet = book.sheet_by_index(0)
        #行数
        rows = sheet.nrows
        for i in range(rows):
            row = sheet.row_values(i+1)
            if row[2] =='0':
                break
            row.append(shopname)
            text_row = list_to_string(row)
            with open(floderName+'\\'+'合并_客服数据-月表.csv','a') as f :
                f.write(text_row+'\n')
    print('客服数据-月表数据已合并！')
floderName = str(year)+'-'+str(month)+'-'+str(nowday)+'处理'
mkdir(floderName)
if   __name__ == '__main__':
    CookieList = read_cookie_file()
    print('***即将下载店铺数据***')
    for cookie in CookieList:
        cookie=cookie.replace('\n', '')
        headers = {
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/95.0.4638.69 Safari/537.36',
            'cookie' : cookie}
        dpurl = 'https://pigeon.jinritemai.com/backstage/currentuser?_ts=1645164132807&biz_type=4&_pms=1'
        personalData = requests.get(dpurl,headers=headers)
        personalJson = personalData.json()
        ShopName = personalJson['data']['ShopName']
        print('*'*22)
        print('当前店铺名称为%s：'%ShopName)
        get_yesterday()
        time.sleep(2)
        get_weekly()
        time.sleep(2)
        get_monthly()
        time.sleep(2)
        get_yesterday_kefu()
        time.sleep(2)
        get_weekly_kefu()
        time.sleep(2)
        get_monthly_kefu()
    print('~'*10)
    print('正在合并数据！请保持导出表头设置相同！')
    hebing()