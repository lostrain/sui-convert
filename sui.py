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


def unzip_kbf(input_files=None):
    zf = zipfile.ZipFile(input_files)
    zf.extractall(path="./")


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
    print("convert done")


if __name__ == "__main__":
    # 执行kbf文件解密
    unzip_kbf("record.kbf")
    ssj_kbf_sqlite_convert("mymoney.sqlite", "record_decrypt.sqlite")
