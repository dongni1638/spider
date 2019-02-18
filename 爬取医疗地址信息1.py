#!/usr/bin/env python 
# -*- coding:utf-8 -*-
import requests
import json
import xlwt

Cookie ="SERVERID=e4a97ef54647b4f0bdd28627b8b00e45|1550219064|1550218752"

headers = {
    'User-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_3) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/72.0.3626.109 Safari/537.36',
    'Cookie': Cookie,
    'Connection': 'keep-alive',
    'Accept': '*/*',
    'Accept-Encoding': 'gzip, deflate',
    'Accept-Language': 'zh-CN,zh;q=0.9',
    'Host': 'dlxxbs.zjzwfw.gov.cn',
    'Referer': 'http://map.zjzwfw.gov.cn/map/map/index.html?_random=0.29481452514730344&webId=1&flag=0'
}

# 存入excel表格
book = xlwt.Workbook()
sheet1 = book.add_sheet('sheet1', cell_overwrite_ok=True)
data_list_content=[]
i= 0
page = 1
#5978条数据，每页10条，循环600次，取完所有数据
while page <=600 :
    pageIndex = str(page)
    # url = "http://dlxxbs.zjzwfw.gov.cn/ReportServer/rest/columns/hotmap/poi/list?callback=jQuery19107838412944351099_1548925767191&key=&areacode=330226000&resource_code=CH0000071&pageIndex="+pageIndex+"&pageSize=10&orderBy=-1&bounds=POLYGON((120.97663879394531+29.07806396484375%2C121.90292358398438+29.07806396484375%2C121.90292358398438+29.60540771484375%2C120.97663879394531+29.60540771484375%2C120.97663879394531+29.07806396484375))&isbounds=1&ecod=false&_=1548925767197"
    #行政区划，页码，边界
    url1 = "http://dlxxbs.zjzwfw.gov.cn/ReportServer/rest/columns/hotmap/poi/list?callback=jQuery19107838412944351099_1548925767191&key=&areacode="
    url2 = "&resource_code=CH0000071&pageIndex="
    url3 = "&pageSize=10&orderBy=-1&bounds=POLYGON(("
    url4 = "))&isbounds=1&ecod=false&_=1548925767197"

    #行政区划
    areacode = str(330000000)
    polygon = str('115.97663879394531+25.07806396484375'
                  '%2C125.90292358398438+25.07806396484375'
                  '%2C125.90292358398438+35.60540771484375'
                  '%2C115.97663879394531+35.60540771484375'
                  '%2C115.97663879394531+25.07806396484375')
    url = url1+areacode+url2+pageIndex+url3+polygon+url4
    print url
    r = requests.get(url, headers=headers)
    json_string = r.text
    print json_string
    # json_string.find('{') 即返回”{“在json_string字符串中的索引位置。
    json_string = json_string[json_string.find('{'):-1]
    json_data = json.loads(json_string)
    comment_list = json_data['result']['list']
    print  comment_list

    for eachone in comment_list:
        message0 = eachone['SZS']
        message1 = eachone['SZQX']
        message2 = eachone['NAME']
        message3 = eachone['DZ']
        data_list_content=[message0,message1,message2,message3]
        j = 0
        for data in data_list_content:
            sheet1.write(i, j, data)
            j += 1
        i += 1
        # if message1 == '宁海县':
        #     data_list_content=[message1,message2,message3]
        #     j = 0
        #     for data in data_list_content:
        #         sheet1.write(i, j, data)
        #         j += 1
        #     i += 1
    page += 1
    print(pageIndex + "写入完成！")
print("全部完成")
book.save('医疗机构地址信息1.xls')