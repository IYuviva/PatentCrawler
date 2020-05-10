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
page_turn_delay = 3  # 自动翻页延时，单位:s

# 打印当前时间
def printTime():
    start_time = strftime("%Y-%m-%d %H:%M:%S", localtime())
    print(start_time)
    return start_time

# get访问登录界面
driver = webdriver.Firefox()
driver.get(url_login)

# 设置超时时间
driver.set_page_load_timeout(20)

# 延时，确保登录网页出现
input('登录界面加载完成后按任意键继续...')
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

run_start_time = printTime()

# 保存首页
writer = pd.ExcelWriter('patentInfoTmp.xlsx')  # 一个Excel中
aPatent.parse(driver.page_source, writer)
start_page = aPatent.getCurPage(driver.page_source)
print(start_page, start_page)

# 确认总页数
page_num = aPatent.getPageNum(driver.page_source)
if page_rq > page_num or page_rq == 0:
    page_rq = page_num

#print(page_num)

new_page = 0
old_page = 0
page_turn_try_cnt = 0  # 翻页尝试次数
page_get_cnt = 0  # 查找异常计数
# 翻页
page_in = aPatent.getCurPage(driver.page_source)
driver.find_element_by_link_text('下一页').click()
while new_page < page_rq:
    time.sleep(page_turn_delay)
    try:
        new_page = aPatent.getCurPage(driver.page_source)
        page_get_cnt = 0
    except:
        page_get_cnt += 1
        if page_get_cnt == 10:
            break

    if (new_page - old_page) >= 1 or (new_page - page_in) == 1:
        old_page = new_page
        try:
            aPatent.parse(driver.page_source, writer, (13 * (new_page - 1)))  # 保存数据
            driver.find_element_by_link_text('下一页').click()
            page_turn_try_cnt = 0
        except:
            old_page -= 1  # 返回，继续翻页
            time.sleep(page_turn_delay)

    else:
        page_turn_try_cnt += 1
        print('no change, ' + str(page_turn_try_cnt))
        time.sleep(page_turn_try_cnt)
        if page_turn_try_cnt == 5:
            break

    print(new_page, old_page)

# 保存到硬盘
writer.save()
print(run_start_time)
printTime()

# excel格式处理
excel = ExcelDeal()
excel.deal('patentInfoTmp.xlsx', 'patentInfo-' + str(start_page) + '-' + str(new_page) + '.xlsx')

print('Finish.')
# 关闭浏览器
driver.quit()
