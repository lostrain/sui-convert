"""
@File    :   sui.py
@Time    :   2024/08/10 20:07:04
@Author  :   lostrain
@Version :   1.0
@Contact :   guodingli@qq.com
@Desc    :   导出随手记账单
"""

import zipfile
import os

import sqlite3
import pandas as pd


def convert_to_excel(sqlite_file, output_file):
    """
    转换SQLite数据到Excel
    """
    conn = sqlite3.connect(sqlite_file)

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

    df.to_excel(output_file, index=False)

    print(f"数据已成功写入到 {output_file}")


def unzip_kbf(input_files=None):
    zf = zipfile.ZipFile(input_files)
    zf.extractall(path="output/")


def ssj_kbf_sqlite_convert(input_file, output_file):
    """
    convert ssj data, after kbf unzip to sqlite,convert it to normal sqlite database file
    :param input_file: the mymoney.sqlite file path
    :param output_file: the convert mymoney.sqlite file path
    :return:
    """
    sqlite_header = (
        0x53,
        0x51,
        0x4C,
        0x69,
        0x74,
        0x65,
        0x20,
        0x66,
        0x6F,
        0x72,
        0x6D,
        0x61,
        0x74,
        0x20,
        0x33,
        0x0,
    )
    if os.path.exists(output_file):
        os.remove(output_file)
    with open(input_file, mode="rb") as f:
        with open(output_file, mode="wb") as fw:
            data_buffer = f.read()
            write_buffer = bytearray(data_buffer)
            index = 0
            while index < len(sqlite_header):
                write_buffer[index] = sqlite_header[index]
                index = index + 1
            fw.write(write_buffer)
    print(f"convert done, output file: {output_file}")


if __name__ == "__main__":
    os.makedirs("output", exist_ok=True)
    # 执行kbf文件解密
    unzip_kbf("record.kbf")
    ssj_kbf_sqlite_convert("output/mymoney.sqlite", "output/record_decrypt.sqlite")
    convert_to_excel("output/record_decrypt.sqlite", "账单记录.xlsx")
