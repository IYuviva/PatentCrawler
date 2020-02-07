# 使用Firefox和selenium完成登录界面
import time
import pandas as pd
from selenium import webdriver
from time import strftime, localtime

from parseHtml import PatentParse
from excelDeal import ExcelDeal

# 配置区
user_name = 'thisisdaming'
user_password = '19911231a'
url_login = 'http://pss-system.cnipa.gov.cn/sipopublicsearch/portal/uilogin-forwardLogin.shtml'
page_rq = 0  # 需要前几页, 如果=0, 那么为最大页数
page_turn_delay = 2  # 自动翻页延时，单位:s

# 打印当前时间
def printTime():
    print(strftime("%Y-%m-%d %H:%M:%S", localtime()))
    return

# get访问登录界面
driver = webdriver.Firefox()
driver.get(url_login)

# 延时，确保登录网页出现
time.sleep(1)

# 找到用户名和密码输入框
driver.find_element_by_id('j_username').send_keys(user_name)
driver.find_element_by_id('j_password_show').send_keys(user_password)

# 建立解析对象
aPatent = PatentParse()

# 网页内容加载完成标志
input('加载完成后按任意键继续...')
html_flag = False
while not html_flag:
    try:
        tmp = aPatent.getFirstPage(driver.page_source)
        if tmp is not None:
            html_flag = True
    except:
        time.sleep(1)

    print('wait...')
    time.sleep(1)


# 切换到列表式
time.sleep(2)
driver.find_element_by_link_text('列表式').click()
time.sleep(2)

printTime()
# 保存首页
writer = pd.ExcelWriter('patentInfoTmp.xlsx')  # 一个Excel中
aPatent.parse(driver.page_source, writer)
print(1, 1)

# 确认总页数
page_num = aPatent.getPageNum(driver.page_source)
if page_rq > page_num or page_rq == 0:
    page_rq = page_num

#print(page_num)

# 自动保存5页
new_page = 0
old_page = 0
# 翻页
page_in = aPatent.getCurPage(driver.page_source)
driver.find_element_by_link_text('下一页').click()
while new_page < page_rq:
    time.sleep(page_turn_delay)
    new_page = aPatent.getCurPage(driver.page_source)
    if (new_page - old_page) >= 1 or (new_page - page_in) == 1:
        old_page = new_page
        aPatent.parse(driver.page_source, writer, (13 * (new_page - 1)))  # 保存数据
        try:
            driver.find_element_by_link_text('下一页').click()
        except:
            old_page -= 1  # 返回，继续翻页
            time.sleep(page_turn_delay)
        # 翻页
    else:
        pass

    print(new_page, old_page)

# 保存到硬盘
writer.save()
printTime()

# excel格式处理
excel = ExcelDeal()
excel.deal('patentInfoTmp.xlsx', 'patentInfoNew.xlsx')

print('Finish.')
# 关闭浏览器
driver.quit()
