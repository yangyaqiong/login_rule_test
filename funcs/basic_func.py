#encoding=utf-8
from selenium import webdriver
import xlrd
import xlwt
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException,TimeoutException
from xlutils.copy import copy
import xlutils
import os
import time
#current_time=time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()))
import matplotlib.mlab as mlab
import matplotlib.pyplot as plt
import xlsxwriter
#启动浏览器，跳转至url
def setUp(url):
    global driver
    driver=webdriver.Firefox(executable_path="D:\\geckodriver")
    driver.get(url)
#获取表格
def dataGetExcel(path):
    data=xlrd.open_workbook(path)
    table=data.sheets()[0]
    rows=table.nrows
    cols=table.ncols
    #print rows,cols
    return table._cell_values
#获取测试数据Excel表格行数
def getRows():
    os.chdir(r'E:\\pythonProject\\login_rule_test\\test_result\\')
    result_excel = xlrd.open_workbook(r'test_result.xlsx')
    return result_excel.sheet_by_index(0).nrows
#输出测试结果
def test_result(test_data,result,error,i,j):
    os.chdir(r'E:\\pythonProject\\login_rule_test\\test_result\\')
    result_excel=xlrd.open_workbook(r'test_result.xlsx')
    table=copy(result_excel)
    sheet1=table.get_sheet(0)
    sheet2=table.get_sheet(1)
    rows=getRows()
    sheet1.write(rows, 0, test_data)
    sheet1.write(rows,1,result)
    sheet1.write(rows,2,error)
    sheet1.write(rows,3,current_time)
    sheet2.write(1, 0, i)
    sheet2.write(1, 1, j)
    table.save(r'test_result.xlsx')
#生成饼图结果

def add_chart():
    os.chdir(r'E:\\pythonProject\\login_rule_test\\test_result\\')
    table1=xlrd.open_workbook(r'test_result.xlsx')
    sheet1=table1.sheet_by_index(1)
    sucess = sheet1.cell_value(1, 0)
    fail = sheet1.cell_value(1, 1)
    workbook = xlsxwriter.Workbook("expense01.xlsx")
    worksheet = workbook.add_worksheet()
    worksheet.write(0,0,u"成功")
    worksheet.write(0,1,u"失败")
    worksheet.write(1,0,sucess)
    worksheet.write(1,1,fail)
    chart=workbook.add_chart({"type":"pie"})
    chart.add_series({
        # "name":"饼形图",
        "categories": "=Sheet1!$A$1:$B$1",
        "values": "=Sheet1!$A$2:$B$2",
        # 定义各饼块的颜色
        "points": [
            {"fill": {"color": "green"}},
            {"fill": {"color": "red"}}
        ]
    })

    chart.set_title({"name": "test_result"})
    chart.set_style(3)
    worksheet.insert_chart("B7", chart)
    workbook.close()
'''
def add_chart():
    os.chdir(r'E:\\pythonProject\\login_rule_test\\test_result\\')
    table1 = xlrd.open_workbook(r'test_result.xlsx')
    workbook=xlutils.copy.copy(table1)
    print table1,type(workbook)
    sheet1 = workbook.get_sheet(1)
    chart = workbook.add_chart({"type": "pie"})
    chart.add_series({
        # "name":"饼形图",
        "categories": "=Sheet1!$A$1:$B$1",
        "values": "=Sheet1!$A$2:$B$2",
        # 定义各饼块的颜色
        "points": [
            {"fill": {"color": "green"}},
            {"fill": {"color": "red"}}
        ]
    })
    print u'饼图已生成'
    chart.set_title({"name": "test_result"})
    chart.set_style(3)
    print u"插入D5"
    sheet1.insert_chart("D5", chart)
    workbook.save("test_result.xlsx")
    workbook.close()
'''
#获取测试所需的环境和xpath等信息
def getEnvData(env):
    os.chdir(r'E:\\pythonProject\\login_rule_test\\test_data\\')
    data = xlrd.open_workbook(r'env_data.xls')
    table=data.sheet_by_index(0)
    for i in range(table.ncols):
        if (table.cell_value(0, i) == env):
            url = table.cell_value(1, i)
            login_xpath=table.cell_value(2,i)
            psw_xpath=table.cell_value(3,i)
            login_botton_xpath=table.cell_value(4,i)
            logout_botton_xpath=table.cell_value(5,i)
            error_close_xpath=table.cell_value(6,i)
            error1_xpath=table.cell_value(7,i)
            error2_xpath=table.cell_value(8,i)
            return url,login_xpath,psw_xpath,login_botton_xpath,logout_botton_xpath,error_close_xpath,error1_xpath,error2_xpath
        else:
            print ''
# 获取某个元素的值
def element_text(element_path):
    try:
        return WebDriverWait(driver, 20).until(lambda x: x.find_element_by_xpath(element_path)).text
    except :
        print u"未找到元素：" + element_path

# 向文本框输入text，获取xpath为：text_xpath
def sendkeys_text(text_xpath, text):
    try:
        element = WebDriverWait(driver, 20).until(lambda x: x.find_element_by_xpath(text_xpath))
        element.clear();
        element.send_keys(text)
    except \
            :
        print u"未找到元素：" + text_xpath
# 点击按钮/选框，获取xpath为：botton_xpath
def click_botton(botton_xpath):
    try:
        WebDriverWait(driver, 30).until(lambda x: x.find_element_by_xpath(botton_xpath)).click()
    except :
        print u"未找到元素：" + botton_xpath

# 关闭浏览器,并重启
def close_Browser():
    driver.quit()

#登录
#def login():

#登录规则校验
def login_rule(sys):
    envdata=getEnvData(sys)
    url=envdata[0]
    login_xpath=envdata[1]
    psw_xpath=envdata[2]
    login_botton_xpath=envdata[3]
    logout_botton_xpath=envdata[4]
    error_close_xpath=envdata[5]
    error1_xpath=envdata[6]
    error2_xpath=envdata[7]
    setUp(url)
    datas=dataGetExcel(r'E:\\pythonProject\\\login_rule_test\\test_data\\data.xls')
    fail = u"登录失败"
    success = u"登录成功"
    i=0
    j=0
    for k,v in datas:
        global current_time
        current_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
        driver.find_element_by_link_text(u"登录").click()
        reslutlist = ''
        error_mess = ''
        sendkeys_text(login_xpath,k)
        sendkeys_text(psw_xpath,v)
        click_botton(login_botton_xpath)
        time.sleep(3)
        try:
        #如果可以找到“退出”按钮，就点击退出
            WebDriverWait(driver,10).until(lambda x: x.find_element_by_link_text(u'退出')).click()
            i=i+1
            print k + success
            reslutlist=success
            error3='no error'
            test_data = k + ',' + v
            test_result(test_data,reslutlist,error3,i,j)   #结果输出有问题
        except:
            j=j+1
            error1=element_text(error1_xpath)
            error2=element_text(error2_xpath)
            print error1,error2
            if error1!='' or error2!='':
                print k+fail
                test_data=k+','+v
                reslutlist = fail
                error_mess=error_mess+error1+error2
            else:
                print k + fail
                test_data=k+','+v
                #reslutlist.append(k + ',')
                reslutlist.append( fail)
                error_mess.append(u"其他异常-无错误信息")
            test_result(test_data,reslutlist,error_mess,i,j)
            WebDriverWait(driver, 10).until(lambda x: x.find_element_by_xpath(error_close_xpath)).click()

