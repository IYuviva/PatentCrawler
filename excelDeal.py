# 本文处理Excel的格式
import pandas as pd
import xlrd
import xlwt
from xlutils.copy import copy

class ExcelDeal:
    def __init__(self):
        pass

    def deal(self, raw_excel, out_excel):
        # Excel读取
        frame = pd.DataFrame(pd.read_excel(raw_excel))

        # Excel去除重复行标题
        frame = frame.drop_duplicates(keep=False)

        # 删除第1列
        frame = frame.drop(frame.columns[0], axis=1)

        # 重新建立索引, 从1开始
        frame = frame.reset_index(drop=True)
        frame.index += 1

        # 写入新的Excel
        frame.to_excel(out_excel)

    def format_deal(self, raw_excel, out):
        # 打开读取文件
        data = xlrd.open_workbook(raw_excel, formatting_info=False)
        sheet = data.sheet_by_index(0)
        print("总行数：" + str(sheet.nrows))

        # 写文件
        workbook = xlwt.Workbook()
        worksheet = workbook.add_sheet('sheet_name')
        # 调整列宽
        worksheet.col(0).width = 256 * 20
        worksheet.write(0,0,'raw_data')
        workbook.save(out)


# 测试函数
if __name__ == '__main__':
    excel = ExcelDeal()
    excel.deal('patentInfo.xlsx', 'patentInfo3.xlsx')
    excel.format_deal('patentInfo3.xlsx', 'patentInfo4.xlsx')
