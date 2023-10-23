# -*- coding: UTF-8 -*-
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
# from selenium.webdriver.support.select import Select
from selenium.webdriver.common.action_chains import ActionChains
# from selenium.common import exceptions as seleniumException

import openpyxl

from tkinter import Tk, messagebox

import traceback
from datetime import datetime
from time import sleep
import os


def clickBtn(driver,str_xpath):
    # 等待
    WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH,str_xpath)))

    # 点击按钮
    driver.find_element(By.XPATH,str_xpath).click()

def sendKeys(driver,str_xpath,var_input):
    # 等待
    WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH,str_xpath)))

    element=driver.find_element(By.XPATH,str_xpath)

    if element.tag_name=='input':
        str_value=str(element.get_attribute('value')).strip()
        if str_value !='':
            element.clear()

    # 输入
    element.send_keys(var_input)

def selectItem(driver,str_xpath,str_item):
    # 等待
    WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH,str_xpath)))

    # 选择
    element=driver.find_element(By.XPATH,str_xpath)
    # sel=Select(element)
    # list_items=sel.options
    
    tag_name=element.tag_name

    if tag_name =='select':
        tag_name_sub='option'
    elif tag_name == 'ul':
        tag_name_sub='li'
    elif tag_name =='li':
        # 查找父元素
        element=element.find_element(By.XPATH,'..') 
        tag_name_sub='li'

    # 查找元素下面的所有满足条件的标签
    list_items=element.find_elements(By.TAG_NAME,tag_name_sub)

    str_item=str_item.strip()

    for i in range(len(list_items)):
        if list_items[i].text.strip()==str_item:
            list_items[i].click()
            # sel.select_by_index(i)
            break

def checkOrRadio(driver,str_xpath):
    # 等待
    WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH,str_xpath)))

    # 选择
    element=driver.find_element(By.XPATH,str_xpath)
    if not element.is_selected(): element.click()


def getTable(driver,str_xpath,index_col=0,str_conditon=''):
    # 等待 可见
    WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH,str_xpath)))

    # 获取<table>
    table=driver.find_element(By.XPATH,str_xpath)

    # 获取<table>下所有<tr>
    list_tr=table.find_elements(By.TAG_NAME,'tr')

    list_result=[]

    for tr in list_tr:
        # 获取<tr>下的所有<td>
        list_td=tr.find_elements(By.TAG_NAME,'td')

        list_result_td=[]

        for i in range(len(list_td)):
            td_mark=list_td[index_col]

            if str_conditon=='' or td_mark.text.strip()==str_conditon:
                list_result_td.append(list_td[i])

        list_result.append(list_result_td)

    return list_result

def mouseclick(driver,str_xpath,str_mark):
    # 等待
    WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH,str_xpath)))

    element=driver.find_element(By.XPATH,str_xpath)

    # 创建对象
    action=ActionChains(driver)

    # 鼠标移动到
    action.move_to_element(element)

    # sleep(1)

    if str_mark=='left':
        action.click(element)
    elif str_mark=='double':
        action.double_click(element)
    elif str_mark=='right':
        action.context_click(element)

    # 位移
    action.move_by_offset(10,20)

    # 执行
    action.perform()


def readExcel(filepath):
    wb = openpyxl.load_workbook(filepath,data_only=True)  # 读取文件路径
    ws = wb["Sheet1"]   # 打开指定的工作簿中的指定工作表
    # ws = wb.active  # 打开激活的工作表
    # 转为列表
    list_data = list(ws.values)[1:] # 转为列表
    wb.close()
    return list_data


INDEX_STEP_NUM=0
INDEX_STEP_DESC=1
INDEX_URLORXPTH=2
INDEX_OPERATE=3
INDEX_VALUE=4
INDEX_TABLE=5
INDEX_COLUMN=6

def getDriver():
    webBrower='edge'

    if webBrower=='chrome':
        options = webdriver.ChromeOptions()
    elif webBrower=='edge':
        options = webdriver.EdgeOptions()

    options.add_experimental_option('detach', True) #不自动关闭浏览器
    options.add_argument('--start-maximized')   #浏览器窗口最大化
    # options.add_argument("--headless")    #隐藏窗口

    if webBrower=='chrome':
        driver=webdriver.Chrome(options=options)
    elif webBrower=='edge':
        driver=webdriver.Edge(options=options)

    return driver


def getOperateAndParams(filepath):
    list_operates=readExcel(filepath)

    dict_operateAndParams={}
    dict_operate_step={}
    dict_params_step={}
    length=len(list_operates)
    tname=''

    # 获取步骤的操作及参数
    for i in range(length):
        num_step=str(list_operates[i][INDEX_STEP_NUM])

        if num_step not in dict_operate_step:
            dict_operate_step[num_step]=[list_operates[i]]
        else:
            dict_operate_step[num_step].append(list_operates[i])

        tablename=str(list_operates[i][INDEX_TABLE]).strip()

        if num_step not in dict_params_step:
            if tablename =='数据.xlsx' or tablename=='上一步结果集':
                tname=tablename
                dict_params_step[num_step]=tablename

        dict_operateAndParams[num_step]=[dict_operate_step[num_step],tname]

    # print(dict_operateAndParams)
    return dict_operateAndParams


def main():

    driver=getDriver()

    url='file:///'+os.path.abspath('使用说明.html')

    driver.get(url)

    str_xpath='//*[@id="btn_start"]'

    # 等待直到开始按钮不可点击
    WebDriverWait(driver, 60).until_not(EC.element_to_be_clickable((By.XPATH,str_xpath)))

    dict_operateAndParams=getOperateAndParams('步骤.xlsx')

    li_params=[]
    dict_tablename={}

    # 按步骤循环运行
    for li_step in dict_operateAndParams.values():
        li_operate=li_step[0]

        tablename=li_step[1]

        if tablename =='数据.xlsx':
            if tablename not in dict_tablename:
                li_params=readExcel(tablename)
                dict_tablename[tablename]=li_params

        if len(li_operate) !=1 :
            for li_p in li_params:
                for li_o in li_operate:
                    run_operate(driver,li_o,li_p)
        else:
            for li_o in li_operate:
                run_operate(driver,li_o,li_params)

    return driver


list_lastReuslt=[]

def run_operate(driver,list_operate,list_param):
    print(list_operate[INDEX_STEP_DESC])

    str_xpath=str(list_operate[INDEX_URLORXPTH]).strip()
    mark_operate=str(list_operate[INDEX_OPERATE]).strip()

    if mark_operate=='打开网址':
        url=str(list_operate[INDEX_URLORXPTH])
        driver.get(url)

    elif mark_operate=='输入':
        str_input=str(list_operate[INDEX_VALUE]).strip()
        sendKeys(driver,str_xpath,str_input)

    elif mark_operate=='点击':
        clickBtn(driver,str_xpath)

    elif mark_operate=='等待':
        # driver.implicitly_wait(time_to_wait=float(list_operate[INDEX_VALUE])) # 隐式等待，有元素就继续，没有就等待
        sleep(float(list_operate[INDEX_VALUE]))

    elif mark_operate=='等待-出现':
        str_label=str(list_operate[INDEX_VALUE]).strip()
        WebDriverWait(driver, 60).until(EC.text_to_be_present_in_element((By.XPATH,str_xpath),str_label))

    elif mark_operate=='切换':
        driver.switch_to.frame(driver.find_element(By.XPATH,str_xpath))

    elif mark_operate=='选择-变量':
        column=int(list_operate[INDEX_COLUMN])
        str_item=str(list_param[column-1]).strip()
        selectItem(driver,str_xpath,str_item)

    elif mark_operate=='输入-变量':
        column=int(list_operate[INDEX_COLUMN])
        str_input=str(list_param[column-1]).strip()
        sendKeys(driver,str_xpath,str_input)

    elif mark_operate=='切换回默认窗口':
        driver.switch_to.default_content()

    elif mark_operate=='选中':
        checkOrRadio(driver,str_xpath)

    elif mark_operate=='获取表格':
        global list_lastReuslt

        str_tablename=str(list_operate[INDEX_TABLE]).strip()

        if str_tablename =='' or str_tablename =='None':
            str_conditon=str(list_operate[INDEX_VALUE]).strip()
            if str_conditon =='' or str_conditon =='None':
                list_lastReuslt=getTable(driver,str_xpath)
            else:
                column=int(list_operate[INDEX_COLUMN])
                list_lastReuslt=getTable(driver,str_xpath,column-1,str_conditon)
        else:
            column=int(list_operate[INDEX_COLUMN])
            str_conditon=str(list_param[column-1]).strip()
            index_col=int(list_operate[INDEX_VALUE])-1
            list_lastReuslt=getTable(driver,str_xpath,index_col,str_conditon)

    elif mark_operate=='点击-单元格':
        column=int(list_operate[INDEX_COLUMN])
        td=list_lastReuslt[0][column-1]
        td.click()

    elif mark_operate=='选择':
        str_item=str(list_operate[INDEX_VALUE]).strip()
        selectItem(driver,str_xpath,str_item)

    elif mark_operate=='鼠标双击':
        mouseclick(driver,str_xpath,'double')
    
    elif mark_operate=='双击-单元格':
        column=int(list_operate[INDEX_COLUMN])
        td=list_lastReuslt[0][column-1]
        ActionChains(driver).double_click(td).perform()

    elif mark_operate=='回车':
        sendKeys(driver,str_xpath,Keys.ENTER)
        sendKeys(driver,str_xpath,Keys.DOWN)


if __name__=='__main__':
    try:
        print('--开始--')

        t1=datetime.now()

        root = Tk()
        root.wm_attributes('-topmost', 1)
        root.withdraw()

        # 运行
        driver=main()

        messagebox.showinfo("运行成功！", "运行成功！\n点确定或关闭此提示框，浏览器退出。",parent=root)

        t2=datetime.now()
        print('--结束--\n开始时间：%s\n结束时间：%s' %(t1,t2))

        print('浏览器退出')
        driver.quit()       # 退出

        root.destroy()
    except Exception as e:
        # print(type(e))
        str_error='×错误：\n%s\n\n' %traceback.format_exc()
        print(str_error)
        with open('log_error.txt','a',encoding='utf8') as f:
            f.write('%s %s' %(datetime.now().strftime('%Y-%m-%d %H:%M:%S'),str_error))

        messagebox.showerror("×运行错误：", str_error,parent=root)
        root.destroy()