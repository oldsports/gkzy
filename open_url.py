# -*- coding: UTF-8 -*- 
import urllib
import time
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.select import Select
import xlrd
import xlwt
from xlwt import Workbook
from xlwt import Worksheet


from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from urllib.parse import quote

from lxml import etree
from bs4 import BeautifulSoup






headers_yx=['院校名称','填报次序','最高分','最低分','录取人数']
headers_zy=['院校名称','专业代号','专业名称','填报次序','最高分','最低分','最低分位次','录取人数']
# headers_zy=list(headers_zy)
chrome_path=r"C:\Users\zhangziyang\AppData\Local\Google\Chrome\Application\chrome.exe"
ie_path=r"C:\Program Files\Internet Explorer\iexplore.exe"
# chrome_web_options="--user-data-dir=C:\Users\zhangziyang\AppData\Local\115Chrome\User Data"
url="https://www.baidu.com"
url_gkzy="http://www1.nm.zsks.cn/xxcx/gkcx/lqmaxmin_18.jsp"
#普通文本格式
style_word=xlwt.XFStyle()
al=xlwt.Alignment()
al.horz=0x02
al.vert=0x01
style_word.alignment=al

#超链接格式
style_url=xlwt.XFStyle()
font=xlwt.Font()
font.underline=True
font.colour_index=4#超链接字体颜色，蓝色
style_url.font=font
# option= webdriver.ChromeOptions()
# option.add_argument()
# option.add_argument("headless") #设置成用户自己的数据目录
browser = webdriver.Firefox()
#browser = webdriver.Chrome(chrome_path)
#webdriver.ChromeOptions
browser.get(url_gkzy)

#批次
# 1     本科提前A
# 2     本科提前B
# 3     本科一批
# 4     本科二批
# 6     专科提前
# 7     高职高专
# C     本科一批B
# E     本科二批B

selector_pc=browser.find_element_by_name("m_pcdm")
selector_pc.find_element_by_xpath("//option[@value='E']").click()
pc_name=browser.find_element_by_name("m_pcdm").find_element_by_xpath("//option[@value='E']").text
# pc_name=selector_pc.find_element_by_xpath("//option[@value='1']").text


#科类
#A 普通文科
#B 普通理科
selector_kl=browser.find_element_by_name("m_kldm")

kl_name=selector_kl.find_element_by_xpath("//option[@value='A']").text
selector_kl.find_element_by_xpath("//option[@value='A']").click()

#排序方式
# selector_pxfs=browser.find_element_by_name("m_pxfs")
# selector_pxfs.find_element_by_xpath("//option[@value='1']").click()
browser.find_element_by_xpath(".//*[@name='m_pxfs']/option[@value='1']").click()


#院校代号

# option_list_yxdh=browser.find_element_by_tag_name("option")


selector_m_yxdh=browser.find_element_by_name("m_yxdh")
#fine_elementssssss    aaaaaaaa!!!!
option_list_yxdh=selector_m_yxdh.find_elements_by_tag_name("option")
# btn_smt=browser.find_element_by_name("query")
# btn_smt.click()
# print(option_list_yxdh)
# i=0
# option_value=[0 for i in range(len(option_list_yxdh))]
# option_text=[0 for i in range(len(option_list_yxdh))]
# for option in option_list_yxdh:
#     #获取下拉框的value和text
#     # print ("Value is:%s  Text is:%s" %(option.get_attribute("value"),option.text))
#     option_value[i]=option.get_attribute("value")
#     option_text[i]=option.text
#     # Select(browser.find_element_by_name("m_yxdh")).select_by_index(i)
#     # btn_smt=browser.find_element_by_name("query")
#     # btn_smt.click()
#     # time.sleep(0.1)
#     i=i+1
# yx_option_max=i
# print(yx_option_max)



# selector_m_yxdh.find_element_by_xpath("//option[@value='001']").click()
btn_smt=browser.find_element_by_name("query")
btn_smt.click()
# browser.find_element_by_xpath(".//*[@name='m_yxdh']/option[@value='001']").click()

#生成xls表格
xle_wbk=xlwt.Workbook(encoding='utf-8')
# kl_name=browser.find_element_by_name("m_kldm").find_element_by_xpath("//option[@value='A']").text
sheetName1=kl_name+'专业详情'
sheet=xle_wbk.add_sheet(sheetName1,cell_overwrite_ok=True)
sheetName2=kl_name+'院校详情'
sheet2=xle_wbk.add_sheet(sheetName2,cell_overwrite_ok=True)
row=0
col=0
row2=0
col2=0
#写表头
for header in headers_zy:
    sheet.write(row,col,header,style=style_word)
    col=col+1
row=row+1
col=0
for header2 in headers_yx:
    sheet2.write(row2,col2,header2,style=style_word)
    col2=col2+1
row2=row2+1
col2=0

yx_num=len(option_list_yxdh)
number1=0
number2=0
i=2
for i in range(i,len(option_list_yxdh)+2):
    # option_value=option.get_attribute("value")
    # option_xpath="//option[@value="+option_value[i]+"]"
    # selector_m_yxdh.find_element_by_xpath(option_xpath).click()
    if pc_name=='本科二批B'and kl_name=='普通理科' :
        if i==28 or i==29 or i==30 or i==31 :
            continue 
    if pc_name=='本科二批B'and kl_name=='普通文科':
        if i==21 or i==22 or i==23 or i==24 :
            continue 
    try:
        yx_select=Select(browser.find_element_by_name("m_yxdh")).select_by_index(i)
    except StaleElementReferenceException as msg:
        print(u"查找元素异常%s"%msg+"%d"%i)     
        continue       
    # print(option_list_yxdh[i-2].text)
    xpath_temp="option["+str(i+1)+"]"
    yx_name=browser.find_element_by_name("m_yxdh").find_element_by_xpath(xpath_temp).text
    yx_value=browser.find_element_by_name("m_yxdh").find_element_by_xpath(xpath_temp).get_attribute("value")
    print(yx_name+'   '+str(i))#打印院校名称，测试用

    #点确定
    btn_smt=browser.find_element_by_name("query")
    btn_smt.click()
    # time.sleep(0.1)


    #专业名称
    #测试是否存在表格
    zymc_flag=0
    yx_maxmin_flag=0
    try:
        browser.find_element_by_xpath("//center/p[2]")
    except NoSuchElementException as msg:
        sheet.write(row,col,yx_name,style=style_word)
        row=row+1
        sheet2.write(row2,col2,yx_name,style=style_word)
        row2=row2+1
        print(u"查找元素异常%s"%msg+yx_name)
        continue
    else:
        yx_maxmin_list=browser.find_element_by_xpath("//center/p[1]").find_elements_by_tag_name("tr")
        for yx_maxmin in yx_maxmin_list:
            if yx_maxmin_flag==0:
                yx_maxmin_flag=1
                continue
            else:
                # sheet2.write(row2,col2,yx_value,style=style_word)
                sheet2.write(row2,col2,yx_name,style=style_word)
                # col2=col2+1
                col2=col2+1
                
                #td_num是为了找到第二个格写入超链接，在院校当中没有用
                td_num=1
                # print(yx_name)
                dyg_list=yx_maxmin.find_elements_by_tag_name("td")#单元格
                
                for dyg in dyg_list:
                    # col2=col2+2
                    yx_text=dyg.text
                    sheet2.write(row2,col2,yx_text,style=style_word)
                    td_num=td_num+1
                    col2=col2+1
            row2=row2+1
            col2=0  
        number2=number2+1          

        zymc_list=browser.find_element_by_xpath("//center/p[2]").find_elements_by_tag_name("tr")
        for zymc in zymc_list:
            if zymc_flag==0:
                zymc_flag=1
                continue
            else:
                # sheet.write(row,col,yx_value,style=style_word)
                sheet.write(row,col,yx_name,style=style_word)
                # col=col+1
                col=col+1
                td_num=1
                # print(yx_name)
                dyg_list=zymc.find_elements_by_tag_name("td")#单元格
                if len(dyg_list)<7:
                    # sheet.write(row,col,)
                    td_num=td_num+2
                    col=col+2
                for dyg in dyg_list:
                    zymc_text=dyg.text
                    if td_num==2:
                        LINK_str1="HYPERLINK("
                        LINK_str2=dyg.find_element_by_xpath("p/a").get_attribute('href')
                        LINK=LINK_str1+"\""+LINK_str2+"\";\""+zymc_text+"\")"
                        sheet.write(row,col,xlwt.Formula(LINK),style=style_url)
                    else:
                        sheet.write(row,col,zymc_text,style=style_word)
                    td_num=td_num+1
                    col=col+1
            row=row+1
            col=0    
        number1=number1+1                


xle_wbk.save(pc_name+kl_name+'.xlsx')
if number1!=yx_num:
    print("专业有误，%d"%yx_num+"    %d"%number1)
if number2!=yx_num:
    print("院校有误，%d"%yx_num+"    %d"%number2)

# Worksheet.write(row,col,xlwt.Formula('HYPERLINK("http://www.google.com";"Google")'),style=Worksheet.)
# Worksheet

    # i=i+1


