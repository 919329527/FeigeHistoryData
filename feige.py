# -*- coding: utf-8 -*-
"""
Created on Fri Feb 18 17:59:14 2022

@author: Administrator

#导出飞鸽数据
"""
import requests,time,datetime


year = datetime.datetime.now().year
month = datetime.datetime.now().month
nowday = datetime.datetime.now().day
first_day = str(year)+'-'+str(month)+'-'+'1' + ' 00:00:00'
first_day = time.strptime(first_day, "%Y-%m-%d %H:%M:%S")
first_day = int(time.mktime(first_day))*1000
yesterday = str(year)+'-'+str(month)+'-'+str(nowday-1) + ' 23:59:59'
yesterday = time.strptime(yesterday, "%Y-%m-%d %H:%M:%S")
yesterday = int(time.mktime(yesterday))*1000+999
Nowtime = time.strftime("%Y-%m-%d_%H_%M_%S", time.localtime())
Yesterday = str(year)+'-'+str(month).rjust(2,'0')+'-'+str(nowday-1).rjust(2,'0') #昨天（一天前）
LastWeek = str(year)+'-'+str(month).rjust(2,'0')+'-'+str(nowday-7).rjust(2,'0') #昨天（七天前）
MonthStart = str(year)+'-'+str(month).rjust(2,'0')+'-'+ '01' #本月1号

#获取Cookie
def read_cookie_file():
    CookieList = []
    f = open('cookie.txt',encoding = 'utf-8')
    f = f.readlines()
    for i in f:
        if len(i) >1000:
            CookieList.append(i)
    return CookieList

#获取昨日店铺数据
def get_yesterday():
    print('正在保存店铺数据-历史数据-近1天')
    yesterday_url = 'https://pigeon.jinritemai.com/backstage/exportShopHistoryConversationData?endTime=%s&startTime=%s'%(Yesterday,Yesterday)
    res = requests.get(yesterday_url,headers=headers)
    with open(ShopName+'_'+Yesterday +'_' +'店铺数据.xlsx','wb') as f:
        f.write(res.content)
#获取近7天店铺数据数据
def get_weekly():
    print('正在保存店铺数据-历史数据-近7天')
    weekly_url = 'https://pigeon.jinritemai.com/backstage/exportShopHistoryConversationData?endTime=%s&startTime=%s'%(Yesterday,LastWeek)
    res = requests.get(weekly_url,headers=headers)
    with open(ShopName+'_'+LastWeek+'-'+Yesterday +'_' +'店铺数据.xlsx','wb') as f:
        f.write(res.content)
#获取本月店铺数据
def get_monthly():
    print('正在保存店铺数据-历史数据-1号至%s号'%(Yesterday.split('-')[-1]))
    monthly_url = 'https://pigeon.jinritemai.com/backstage/exportShopHistoryConversationData?endTime=%s&startTime=%s'%(Yesterday,MonthStart)
    res = requests.get(monthly_url,headers=headers)
    with open(ShopName+'_'+MonthStart+'-'+Yesterday +'_' +'店铺数据.xlsx','wb') as f:
        f.write(res.content)

#获取昨日客服数据
def get_yesterday_kefu():
    print('正在保存客服数据-历史数据-近1天')
    strattime = yesterday-86399999
    yesterday_url = 'https://pigeon.jinritemai.com/backstage/exportHistoricalStaffStatistics?endTime=%s&queryStaffName=&sortTag=0&sortType=0&startTime=%s'%(yesterday,strattime)
    res = requests.get(yesterday_url,headers=headers)
    with open(ShopName+'_'+Yesterday +'_' +'客服数据.xlsx','wb') as f:
        f.write(res.content)
# 获取近7天客服数据
def get_weekly_kefu():
    print('正在保存客服数据-历史数据-近7天')
    strattime = yesterday-604799999
    weekly_url = 'https://pigeon.jinritemai.com/backstage/exportHistoricalStaffStatistics?endTime=%s&queryStaffName=&sortTag=0&sortType=0&startTime=%s'%(yesterday,strattime)
    res = requests.get(weekly_url,headers=headers)
    with open(ShopName+'_'+LastWeek+'-'+Yesterday +'_' +'客服数据.xlsx','wb') as f:
        f.write(res.content)
#获取本月客服数据
def get_monthly_kefu():
    print('正在保存客服数据-历史数据-1号至%s号'%(Yesterday.split('-')[-1]))
    monthly_url = 'https://pigeon.jinritemai.com/backstage/exportHistoricalStaffStatistics?endTime=%s&queryStaffName=&sortTag=0&sortType=0&startTime=%s'%(yesterday,first_day)
    res = requests.get(monthly_url,headers=headers)
    with open(ShopName+'_'+MonthStart+'-'+Yesterday +'_' +'客服数据.xlsx','wb') as f:
        f.write(res.content)

if   __name__ == '__main__':
    CookieList = read_cookie_file()
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
        time.sleep(5)
        get_weekly()
        time.sleep(5)
        get_monthly()
        time.sleep(5)
        get_yesterday_kefu()
        time.sleep(5)
        get_weekly_kefu()
        time.sleep(5)
        get_monthly_kefu()




