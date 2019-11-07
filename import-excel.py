import  pandas  as pd
import pymysql
import xlrd
# Open the workbook and define the worksheet
book = xlrd.open_workbook("users.xlsx")
sheet = book.sheet_by_index(0)


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
query = "INSERT INTO `orders` (`name`,`address`,`order_id`,`order_weight`,`order_time`) VALUES (%s,%s,%s,%s,%s)"


# 创建一个for循环迭代读取xls文件每行数据的, 从第二行开始是要跳过标题
for r in range(1, sheet.nrows):
      name      = sheet.cell(r,0).value
      address   = sheet.cell(r,1).value
      order_id  = sheet.cell(r,2).value
      order_weight = sheet.cell(r,3).value
      order_time = sheet.cell(r,4).value

      values = (name,address,order_id.strip(),order_weight,order_time)

      # 执行sql语句
      cursor.execute(query, values)

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

