import time

from selenium.common.exceptions import TimeoutException
from selenium.webdriver import Chrome, ActionChains
from selenium.webdriver import ChromeOptions
from lxml import etree
from openpyxl import Workbook, load_workbook
import properties_read


chrome_options = ChromeOptions()
# 修改windows.navigator.webdriver，防机器人识别机制，selenium自动登陆判别机制
chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])
# 隐藏"Chrome正在受到自动软件的控制"
chrome_options.add_argument('disable-infobars')

driver = Chrome(chrome_options=chrome_options)
#窗口最大化
driver.maximize_window()
#隐式等待
#driver.set_page_load_timeout (15)

# CDP执行JavaScript 代码  重定义windows.navigator.webdriver的值
driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
    "source": """
        Object.defineProperty(navigator, 'webdriver', {
          get: () => undefined
        })
      """
})
#起始页面
start_url = 'https://auth.cantonfair.org.cn/'
exhibition_url = 'https://www.cantonfair.org.cn/exhibition'

p = properties_read.Properties('config.properties')
try:
    #检索类目
    item = p.getProperties('craw_comms').split(",")
    #用户名
    userName = p.getProperties('userName')
    # 用户名
    password = p.getProperties('password')
except:
    print("config文件不正确")

try:
    driver.get(start_url)
except TimeoutException:
    print("超时，停止加载当前页")
    driver.execute_script("window.stop()")
time.sleep(1)
# 输入账号
driver.find_element_by_xpath("//input[@type='text']").send_keys(userName)
# 输入密码
driver.find_element_by_xpath("//input[@type='password']").send_keys(password)
# 点击搜索
driver.find_element_by_xpath("//input[@class='ivu-checkbox-input']").click()
# 点击登录
driver.find_element_by_xpath("//button[@class='btn__red btn__red--long ivu-btn ivu-btn-default ivu-btn-long']").click()
time.sleep(5)
driver.execute_script("window.stop()")
try:
    driver.get(exhibition_url)
except TimeoutException:
    print("超时，停止加载当前页")
time.sleep(2)
itemDom =  driver.find_element_by_xpath("//a[@class='item'][.='{}']".format(item[0]))
if len(item)==2:
    #鼠标悬停到一级菜单
    ActionChains(driver).move_to_element(itemDom).perform()
    time.sleep(0.5)
    #点击二级菜单
    driver.find_element_by_xpath("//a[@target='_self'][.='{}']".format(item[1])).click()
else:
    # 点击一级菜单
    itemDom.click()
    pass
time.sleep(1)
#切换到企业
driver.find_element_by_xpath("//ul[@class='tabs-bar']/li[1]").click()
time.sleep(1)
#字母排序按钮
sort_btn = driver.find_element_by_xpath("//ul[@class='filter']//li[3]")
#点击字母排序
driver.execute_script("arguments[0].click();", sort_btn)
time.sleep(1)
#存放所有企业link
link_temps = [];
#//div[@class='exhibitor-item']//a//@href
links = driver.find_elements_by_xpath("//div[@class='exhibitor-item']//a")

for link in links:
    s_href = link.get_attribute("href")
    link_temps.append(s_href)

while True:
    #下一页DOM
    next_page_btn = driver.find_element_by_xpath("//div[@class='company']//li[@title='下一页']")
    # 是否是最后一页
    if "ivu-page-disabled" in next_page_btn.get_attribute("class"):
        break
    #点击下一页
    driver.execute_script("arguments[0].click();", next_page_btn)
    time.sleep(2)
    #取得当前页面所有link
    links = links = driver.find_elements_by_xpath("//div[@class='exhibitor-item']//a")
    for link in links:
        s_href = link.get_attribute("href")
        link_temps.append(s_href)
print(link_temps)

# 创建文件对象
wb = Workbook()
# 获取第一个sheet
ws = wb.active
# 标题
result_tittle = ["企业名称", "企业类型", "成立年份", "注册资本", "企业规模", "主要目标客户", "主营展品", "地址", "所在地区", "网址"
    , "业务联系人", "邮箱", "电话", "手机", "传真", "邮编"]
first_Flag = True
# 输出数据
for link_temp in link_temps:
    #正文
    result_value = []
    #页面跳转
    js_url = "window.location.href = '{}{}';".format(link_temp,"/company")
    driver.execute_script(js_url)
    time.sleep(1)
    try:
        # 企业联系方式DOM
        contact_btn = driver.find_element_by_xpath("// a[ @class ='ex-60__contact-view']")
        # 点击企业联系方式
        driver.execute_script("arguments[0].click();", contact_btn)
        time.sleep(0.5)
        html = etree.HTML(driver.page_source)
        cell_items1 = html.xpath("//div[@class='ex-60__baseinfo']//div[@class='cell-item']")
        cell_items2 = html.xpath("//div[@class='ex-60__contact']//div[@class='cell-item']")
        #结果标题
        if first_Flag:
            ws.append(result_tittle)
            first_Flag = False
        print("---------------------------------爬取："+link_temp + "------------------------------------------")
        for cell_item1 in cell_items1:
            result_value.append(cell_item1[1].text)
        for cell_item2 in cell_items2:
            result_value.append(cell_item2[1].text)
        ws.append(result_value)
    except:
        print("---------------------------------"+link_temp + "爬取失败--------------------------------------")

wb.save("result_{}.xlsx".format(time.strftime("%Y%m%d", time.localtime())))
print("---------------------------------爬取完毕--------------------------------------")

