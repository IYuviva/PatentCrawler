# 本文处理Excel的格式
import pandas as pd
import xlrd
import xlwt
from xlutils.copy import copy
import os

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

    def list_files(self):
        filelist = []
        # 查找当前目录下的文件
        for root, dirs, files in os.walk(".", topdown=False):
            for name in files:
                str = os.path.join(root, name)
                if str.split('.')[-1] == 'xlsx':
                    filelist.append(str.split('./')[-1])
        filelist.remove('patentInfoTmp.xlsx')
        filelist.remove('patentInfo-Merge.xlsx')
        print(filelist)
        filelist.sort()
        print(filelist)
        print(len(filelist))
        return filelist

    def multi_excels_merge(self, file_list, out):
        tmp_file = []

        for i in file_list:
            tmp_file.append(pd.read_excel(i))

        writer = pd.ExcelWriter(out)
        pd.concat(tmp_file).to_excel(writer, 'Sheet1')

        writer.save()
        print('Merge ok.')

    def multi_excels_deal(self, raw_excel, out_excel):
        # Excel读取
        frame = pd.DataFrame(pd.read_excel(raw_excel))

        # Excel去除重复行标题
        frame = frame.drop_duplicates(keep=False)

        # 删除第1列和第2列
        frame = frame.drop(frame.columns[0:2], axis=1)

        # 重新建立索引, 从1开始
        frame = frame.reset_index(drop=True)
        frame.index += 1

        # 写入新的Excel
        frame.to_excel(out_excel)
        print('Deal ok.')

# 测试函数
if __name__ == '__main__':
    excel = ExcelDeal()

    files_list = excel.list_files()
    excel.multi_excels_merge(files_list, 'patentInfoTmp.xlsx')
    excel.multi_excels_deal('patentInfoTmp.xlsx', 'patentInfo-Merge.xlsx')