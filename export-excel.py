#coding:utf8
import pymysql.cursors
import xlwt
import sys
import importlib
importlib.reload(sys)

#建立一个MySQL连接
conn = pymysql.connect(host='localhost',
                                      user='root',
                                      password='root',
                                      db='homestead',
                                      charset='utf8mb4',
                                      cursorclass=pymysql.cursors.DictCursor)
cursor=conn.cursor()

sql="select * from orders"

count = cursor.execute(sql)
print(count)

fileds = [filed[0] for filed in cursor.description]  # 列表生成式，所有字段
all_data = cursor.fetchall() #所有数据
#写excel
book = xlwt.Workbook() #先创建一个book
sheet = book.add_sheet('sheet1') #创建一个sheet表
# col = 0
# for field in fileds: #写表头的
#     sheet.write(0, col, field)
#     col += 1
#enumerate自动计算下标
for col, field in enumerate(fileds): #跟上面的代码功能一样
    sheet.write(0, col, field)

#从第一行开始写
row = 1 #行数
for data in all_data:  #二维数据，有多少条数据，控制行数
    for col, field in enumerate(data):  #控制列数
        sheet.write(row, col, data[field])
    row += 1 #每次写完一行，行数加1
book.save(r'./test.xls')







conn.commit()
conn.close()
