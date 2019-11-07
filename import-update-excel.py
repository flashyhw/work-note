import  pandas  as pd
import pymysql
import xlrd
# Open the workbook and define the worksheet
book = xlrd.open_workbook("users.xlsx")
sheet = book.sheet_by_index(1)


#建立一个MySQL连接
database = pymysql.connect(host='localhost',
                                      user='root',
                                      password='root',
                                      db='homestead',
                                      charset='utf8mb4',
                                      cursorclass=pymysql.cursors.DictCursor)


# 获得游标对象, 用于逐行遍历数据库数据
cursor = database.cursor()

# 创建插入SQL语句
query = "update orders set order_weight_shentong=%s where order_id=%s"


# 创建一个for循环迭代读取xls文件每行数据的, 从第二行开始是要跳过标题
for r in range(1, sheet.nrows):
      order_id_shentong     = sheet.cell(r,0).value
      order_weight_shentong  = sheet.cell(r,1).value
      values = (order_weight_shentong,order_id_shentong)
      # 执行sql语句
      print(order_weight_shentong)
      print(order_id_shentong)
      cursor.execute(query, values)
      #print(query)


# 关闭游标
cursor.close()

# 提交
database.commit()

# 关闭数据库连接
database.close()

# 打印结果

columns = str(sheet.ncols)
rows = str(sheet.nrows)
print( '我刚导入了')

