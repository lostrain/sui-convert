"""
@File    :   convert_to_excel.py
@Time    :   2024/08/10 20:22:44
@Author  :   lostrain
@Version :   1.0
@Contact :   guodingli@qq.com
@Desc    :   转换SQLite数据到Excel
"""

import sqlite3
import pandas as pd
import xlwt

# 连接到SQLite数据库，如果数据库文件不存在，会自动在当前目录创建一个
# 数据库文件名为example.db
conn = sqlite3.connect("record_decrypt.sqlite")

# 创建一个游标对象，用于执行SQL语句
cursor = conn.cursor()

sql = """SELECT strftime('%Y/%m/%d %H:%M', a.tradeTime / 1000 + 8 * 3600, 'unixepoch') as 日期,
 case
   when a.type = 1 then '收入'
     when a.type = 0 then '支出'
     when a.type = 2 then '转账'
 end as 收支类型,
 case
   when a.type = 1 then (select case 
                             when (select b.currencyType from t_account b where b.accountPOID = a.sellerAccountPOID) = 'CNY' then a.buyerMoney
                                                     else (select round(a.buyerMoney * d.rate, 2) from 
                                                                    (select b.currencyType from t_account b where b.accountPOID = a.sellerAccountPOID) c, 
                                                                    t_exchange d where c.currencyType = d.sell)
                                                     end)
   when a.type = 0 then (select case 
                             when (select b.currencyType from t_account b where b.accountPOID = a.buyerAccountPOID) = 'CNY' then a.buyerMoney
                                                     else (select round(a.buyerMoney * d.rate, 2) from 
                                                                    (select b.currencyType from t_account b where b.accountPOID = a.buyerAccountPOID) c, 
                                                                    t_exchange d where c.currencyType = d.sell)
                                                     end)
  when a.type = 2 then (select case 
                             when (select b.currencyType from t_account b where b.accountPOID = a.buyerAccountPOID) = 'CNY' then a.buyerMoney
                                                     else (select round(a.buyerMoney * d.rate, 2) from 
                                                                    (select b.currencyType from t_account b where b.accountPOID = a.buyerAccountPOID) c, 
                                                                    t_exchange d where c.currencyType = d.sell)
                                                     end)
 end as 金额,
 case 
   when a.type = 1 then (select d.name from (select b.parentCategoryPOID from t_category b 
                              where b.categoryPOID = a.buyerCategoryPOID) c, t_category d
                                                            where c.parentCategoryPOID = d.categoryPOID)
     when a.type = 0 then (select d.name from (select b.parentCategoryPOID from t_category b 
                              where b.categoryPOID = a.sellerCategoryPOID) c, t_category d
                                                            where c.parentCategoryPOID = d.categoryPOID)
 end as 类别,
 case 
   when a.type = 1 then (select b.name from t_category b where b.categoryPOID = a.buyerCategoryPOID)
     when a.type = 0 then (select b.name from t_category b where b.categoryPOID = a.sellerCategoryPOID)
 end as 子类,
 '日常账本' as 所属账本,
 case 
   when a.type = 1 then (select b.name from t_account b where b.accountPOID = a.sellerAccountPOID)
     when a.type = 0 then (select b.name from t_account b where b.accountPOID = a.buyerAccountPOID)
     when a.type = 2 then (select b.name from t_account b where b.accountPOID = a.buyerAccountPOID)
 end as 账户1,
  case 
     when a.type = 2 then (select b.name from t_account b where b.accountPOID = a.sellerAccountPOID)
 end as 账户2,
 a.memo as 备注,
 (select c.name 
     from t_transaction_projectcategory_map b, t_tag c 
     where b.transactionPOID = a.transactionPOID 
     and b.projectCategoryPOID = c.tagPOID
     and b.type = 2) as 标签,
 '' as 地址
 FROM t_transaction a
 WHERE a.type IN (0,1,2)
 order by a.tradetime desc;"""

# 运行一个查询语句
select_sql = sql

# 使用游标执行查询
cursor.execute(select_sql)

# 获取所有查询结果
results = cursor.fetchall()

# 打印结果
for row in results:
    print(row)

# 将结果转换为DataFrame
columns = [column[0] for column in cursor.description]  # 获取列名
df = pd.DataFrame(results, columns=columns)

# 关闭游标和连接
cursor.close()
conn.close()

# 使用xlwt库将DataFrame写入Excel文件
excel_filename = "账单导入.xls"

# 创建一个Excel工作簿
wb = xlwt.Workbook()
# 添加一个工作表
ws = wb.add_sheet("Sheet 1")

# 将DataFrame数据写入Excel工作表
for col_num, col_data in enumerate(df.columns):
    ws.write(0, col_num, col_data)  # 写入列名

for row_num, row_data in enumerate(df.values):
    for col_num, col_data in enumerate(row_data):
        ws.write(row_num + 1, col_num, col_data)  # 写入单元格数据

# 保存Excel文件
wb.save(excel_filename)

print(f"数据已成功写入到 {excel_filename}")
