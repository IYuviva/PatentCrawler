# 解析列表式数据
from lxml import etree
from pandas.core.frame import DataFrame


# 封装成class
class PatentParse:
    def __init__(self):
        pass

    def parse(self, html_data, writer, row=0):
        html = etree.HTML(html_data)

        # 数据处理部分
        list_rq_id = []  # 申请号
        list_public_id = []  # 公开号
        list_patent_name = []  # 专利名称
        list_rq_people = []  # 申请人
        list_rq_date = []  # 申请日
        list_public_date = []  # 公开日

        content = html.xpath('//*[@id="tbody"]')  # 定位

        # 1.申请号
        list_rq_id_tmp = []
        list_tmp = content[0].xpath('tr/td[2]/span')
        for i in range(0, len(list_tmp)):
            list_rq_id_tmp.append(list_tmp[i].get('title'))
            list_rq_id.append(list_rq_id_tmp[i].replace('</FONT>', '').replace('<FONT>', ''))
        # print(list_rq_id)
        # print(len(list_rq_id))

        # 2.申请日列表
        info_list_year = content[0].xpath('tr/td[3]/font/text()')
        info_list_date = content[0].xpath('tr/td[3]/text()')
        # print(info_list_year)
        # print(info_list_date)

        if len(info_list_year) != 0:
            for i in range(0, len(info_list_date)):
                info_list_date[i] = info_list_year[i] + info_list_date[i]

        # 3.公开号
        list_public_id = content[0].xpath('tr/td[4]/span/text()')
        # print(list_public_id)
        # print(len(list_public_id))

        # 4.公开日列表
        info_list3 = content[0].xpath('tr/td[5]/text()')

        # 5.专利名称
        list_patent_name_tmp = []
        list_tmp = content[0].xpath('tr/td[6]/span')
        for i in range(0, len(list_tmp)):
            list_patent_name_tmp.append(list_tmp[i].get('title'))
            list_patent_name.append(list_patent_name_tmp[i].replace('</FONT>', '').replace('<FONT>', ''))
        # print(list_patent_name)
        # print(len(list_patent_name))

        # 6.申请人
        list_rq_people_tmp = []
        list_tmp = content[0].xpath('tr/td[7]/span')
        for i in range(0, len(list_tmp)):
            list_rq_people_tmp.append(list_tmp[i].get('title'))
            list_rq_people.append(list_rq_people_tmp[i].replace('</FONT>', '').replace('<FONT>', ''))
        # print(list_rq_people)
        # print(len(list_rq_people))

        # 对日期进行处理
        for i in info_list_date:
            list_rq_date.append(i.rstrip())

        for i in info_list3:
            list_public_date.append(i.rstrip())

        # print(list_rq_date)
        # print(list_public_date)

        dict_patent_info = {
            "专利名称": list_patent_name,
            "申请号": list_rq_id,
            "申请日": list_rq_date,
            "申请人": list_rq_people
        }
        # print(dict_patent_info)
        # 将字典转换成为数据框 # 多个DataFrame组成1个Excel
        DataFrame(dict_patent_info).to_excel(writer, startrow=row)
        print('Save ok.')

    def getPageNum(self, html_data):
        html = etree.HTML(html_data)
        content = html.xpath('//*[@id="tbody"]')  # 定位
        tmp = content[0].xpath("//*[@class='page_bottom']/p[1]/text()")
        # 总页数
        page_num = tmp[0].split('\xa0')[2]
        return int(page_num)

    def getCurPage(self, html_data):
        html = etree.HTML(html_data)
        content = html.xpath('//*[@id="tbody"]')  # 定位
        tmp = content[0].xpath("//*[@class='input_bottom']/input")
        # 当前页数
        page_cur = tmp[0].get("value")
        return int(page_cur)

    # 检查首页是否加载ok
    def getFirstPage(self, html_data):
        html = etree.HTML(html_data)
        content = html.xpath('//*[@id="result_view_mode"]')  # 定位
        return content[0]

# 测试函数
if __name__ == '__main__':
    f = open('new2.shtml', 'r')
    data = f.read()
    f.close()

    aPatent = PatentParse()
    aPatent.parse(data, 'd.xlsx')
    print(aPatent.getPageNum(data))
    print(aPatent.getCurPage(data))