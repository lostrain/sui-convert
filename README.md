# sui-convert

随手记APP数据转换工具，将官方的kbf文件转换为可读的Excel文件

## 使用方法

- 将`kbf`文件重命名为`record.kbf`，放在项目根目录下
- 运行`sui.py`，生成解密后的`record_decrypt.sqlite`文件
- 运行`convert_to_excel.py`，从`sqlite`中取出数据，得到最终的`账单导入.xls`文件
