# # -*- coding: utf-8 -*-
# # @file : write_book_robot_send_wechat.py
from openpyxl import Workbook

class WriteBook:
    def __init__(self):
        self.book = Workbook()

    def insert_row(self, value_list, sheet_value):
        """
        行数据要与表头数据对应
        :param value_list: 一行对应的值
        :param sheet_value: 表数据（默认索引从0开始）
        :return: 整表填充完整数据
        """
        sheet_value.append(value_list)

    def create_new_sheet(self, table_name, table_index, insert_values):
        """
        :param table_name: 新建表名
        :param table_index: 新建表索引
        :return: 添加表头数据
        """
        sheet = self.book.create_sheet(table_name, table_index)
        sheet.append(insert_values)

    def save_file(self, table_name):
        self.book.save(table_name)


# 新建excel，并创建多个sheet
if __name__ == "__main__":
    write_book = WriteBook()
    write_book.create_new_sheet("JS", 0, ['businessLicenseNumber', 'businessPerson', 'certStr', 'cityCode', 'countyCode', 'creatUser', 'createTime', 'endTime', 'epsAddress', 'epsName', 'epsProductAddress', 'id', 'isimport', 'legalPerson', 'offDate', 'offReason', 'parentid', 'preid', 'processid', 'productSn', 'provinceCode', 'qfDate', 'qfManagerName', 'qualityPerson', 'rcManagerDepartName', 'rcManagerUser', 'startTime', 'warehouseAddress', 'xkCompleteDate', 'xkDate', 'xkDateStr', 'xkName', 'xkProject', 'xkRemark', 'xkType', 'ID', 'EPS_NAME', 'PRODUCT_SN', 'CITY_CODE', 'XK_COMPLETE_DATE', 'XK_DATE', 'QF_MANAGER_NAME', 'BUSINESS_LICENSE_NUMBER', 'XC_DATE', 'NUM_', 'md5_value'])
    sheets = write_book.book.get_sheet_names()
    for i in range(3):
        write_book.insert_row(['yaopin'] * 45, write_book.book.get_sheet_by_name(sheets[0]))
    write_book.save_file("JS.xlsx")



import os
import requests

# 传入文件
def send_file_body(keys, file_path):
    data = {'file': open(file_path, 'rb')}
    id_url = f'https://qyapi.weixin.qq.com/cgi-bin/webhook/upload_media?key={keys}&type=file'
    response = requests.post(url=id_url, files=data)
    json_res = response.json()
    media_id = json_res['media_id']
    wx_url = f'https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key={keys}'
    data = {"msgtype": "file", "file": {"media_id": media_id}}
    result = requests.post(url=wx_url, json=data)
    return result


abs_path = os.path.abspath(__file__)
file_path = abs_path + f'/data.xlsx'
keys = 'xxxxxxxxxx'
# file_path = 'JD.xlsx'
send_file_body(keys, file_path)
