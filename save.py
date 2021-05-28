import requests
from bs4 import BeautifulSoup
import xlwt
import xlrd
from xlutils.copy import copy
import pymysql
import os
import re

def totxt(biz,table):
    try:
        db = pymysql.connect('localhost', 'root', '123456', 'wechat_biz')
    except Exception as e:
        print(e)
    else:
        print("开始将公众号：{} 的数据转化为txt".format(biz))
        cursor = db.cursor()
        search_sql = f"select * from {str(table)}"
        try:
            cursor.execute(search_sql)
        except Exception as e:
            print(f'select {biz} error:', e)
        else:
            position = 0
            if not os.path.exists(str(biz)):
                os.makedirs(str(biz))
            search_result = cursor.fetchall()
            for i in search_result:
                url = str(i[5])
                #totxt_result = 0
                try:
                    respond = requests.get(url)
                except Exception as e:
                    print(e)
                else:
                    html = respond.text
                    bf = BeautifulSoup(html,'lxml')
                    title = bf.find_all('h2',class_ = 'rich_media_title')
                    content = bf.find_all('div',class_ = 'rich_media_content') 
                    if title:
                        t = re.sub(r'\s|\t|\n','',title[0].text)
                    else:
                        t = ''
                    if content:
                        c = re.sub(r'\s|\t|\n','',content[0].text)
                    else:
                        c = ''
                    if t == '' and c == '':
                        continue
                    position = position + 1
                    f = open(r"{}/{}.txt".format(str(biz),position),"w+",encoding='utf-8')
                    f.write(t)
                    f.write('\n')
                    f.write(c)
                    f.close()
                    print('第{}篇文章{}已转化为txt'.format(position,t))

                    
    #return totxt_result

def one_toexcel(biz, table, start_date, end_date):
    try:
        readbook = xlrd.open_workbook('{}.xls'.format(str(biz)), formatting_info=True)
    except Exception as e:
        print(e)
        workbook = xlwt.Workbook(encoding = 'utf-8')
        worksheet = workbook.add_sheet(str(biz))
    else:
        workbook = copy(readbook)
        try:
            worksheet = workbook.get_sheet(str(biz))
        except Exception as e:
            print(e)
            worksheet = workbook.add_sheet(str(biz))

    worksheet.write(0,0, label = 'biz_name')
    worksheet.write(0,1, label = 'app_msg_title')
    worksheet.write(0,2, label = 'msg_create_time')
    worksheet.write(0,3, label = 'app_msg_digest')
    worksheet.write(0,4, label = 'app_msg_url')
    try:
        db = pymysql.connect('localhost', 'root', '123456', 'wechat_biz')
    except Exception as e:
        print(e)
    else:
        print("开始读取公众号：{} 的数据".format(biz))
        print("start_date,end_date: ", start_date,end_date)
        cursor = db.cursor()
        # print(type(infos['aid']))
        search_sql = f"select * from {str(table)}"
        # print(search_sql)
        try:
            cursor.execute(search_sql)
        except Exception as e:
            print(f'select {biz} error:', e)
        else:
            position = 1
            width = [0,0,0,0,0]
            style = xlwt.XFStyle() # 初始化样式
            font = xlwt.Font() # 为样式创建字体
            font.underline = True # 下划线
            style.font = font # 设定样式
            search_result = cursor.fetchall()
            for i in search_result:
                if (i[3] <= end_date and i[3] >= start_date) or start_date == None:
                    print(i)
                    worksheet.write(position,0, label = i[1])
                    worksheet.write(position,1, label = i[6])
                    worksheet.write(position,2, label = i[3])
                    worksheet.write(position,3, label = i[4])
                    worksheet.write(position,4, label = xlwt.Formula('HYPERLINK("%s";"文章永久链接")' %i[5]), style = style)
                else:
                    continue
                width[1] = max(width[1],len(i[6].encode('gb18030')))
                width[3] = max(width[3],len(i[4].encode('gb18030')))
                position += 1
            width[0] = len(biz.encode('gb18030'))
            width[2] = len(search_result[0][3].encode('gb18030'))
            width[4] = len("文章永久链接".encode('gb18030'))
            for k in range(0,5):
                worksheet.col(k).width = 256*width[k]
    db.close()
    if str(biz) == 'XYSTRATEGY':
        workbook.save('兴业策略.xls')
    else:
        workbook.save('{}.xls'.format(str(biz)))
    return


 
