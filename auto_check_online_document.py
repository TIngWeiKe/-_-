#!/usr/bin/env python
# coding: utf-8

# In[157]:


from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from time import sleep
from selenium.webdriver.support.select import Select
import arrow
from selenium.common.exceptions import NoSuchElementException   
from selenium.common.exceptions import NoSuchWindowException   
import sys


# In[158]:


driver = webdriver.Chrome("C:\\Users\\TFD\\Anaconda3\\chromedriver.exe")
wait = WebDriverWait(driver,10)


# In[159]:


driver.maximize_window()
driver.get("https://ssopxy.gov.taipei/SSOPXY/Default.aspx")
id =('j0470@tfd.gov.tw')
password = ('j180065@')


# In[160]:


driver.find_element_by_name('TextBox_userID').send_keys(id)


# In[161]:


driver.find_element_by_name('TextBox_userPwd').send_keys(password)


# In[162]:


driver.find_element_by_name('Button_Login').click()


# In[163]:


driver.find_element_by_class_name('category-button1').click()


# In[164]:


driver.find_element_by_css_selector('input[value=新公文系統]').click()


# In[165]:


child = driver.window_handles[1]      
driver.switch_to.window(child)
print(child)


# In[166]:


driver.implicitly_wait(10)
try:
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "toSaveCheck"))     
    )
finally:
    driver.find_element_by_xpath('//*[@id="toSaveCheck"]').click()


# In[167]:


sleep(2)
driver.find_element_by_id('menu1484').click()


# In[168]:


driver.switch_to.frame('dTreeContent')
driver.find_element_by_name('conditionQueryAll').click()


# In[169]:


sleep(2)
driver.find_element_by_xpath('//*[@id="qryModal"]/div[2]/div/div[2]/div/div/div/table/tbody/tr[4]/td[4]/select/option[3]').click()


# In[170]:


dayOfMonth = int(arrow.now().format('DD'))
month = arrow.now().format('MM')
year = "107"
dayOfWeek=arrow.now().format('d')
if dayOfWeek == "1" :
    dayOfMonth -= 4
elif dayOfWeek == "2" :
    dayOfMonth -= 4
else:
    dayOfMonth -= 1
print('星期'+dayOfWeek)
date = year+month+str(dayOfMonth)


# In[171]:


driver.find_element_by_xpath('//*[@id="qryModal"]/div[2]/div/div[2]/div/div/div/table/tbody/tr[2]/td[2]/input[3]').clear()
driver.find_element_by_xpath('//*[@id="qryModal"]/div[2]/div/div[2]/div/div/div/table/tbody/tr[2]/td[2]/input[3]').send_keys(date)


# In[172]:


driver.find_element_by_name('querySubmit').click()


# In[173]:


try :
    pageSizeTag = driver.find_element_by_xpath('//*[@id="form1"]/div[2]/table[2]/tbody/tr/td/table/tbody/tr/td[1]/span').get_attribute('innerHTML')
    pz = pageSizeTag.find('共')+1
    pageSize = pageSizeTag[pz:pz+3]
    if pageSize.isdigit() :
        print("共有"+pageSize+"筆公文")
    else:
        pageSize = pageSizeTag[pz:pz+2]
        print("共有"+pageSize+"筆公文")
    driver.find_element_by_xpath('//*[@id="form1"]/div[2]/table[2]/tbody/tr/td/table/tbody/tr/td[1]/span/input').send_keys(pageSize)
    sleep(2)
except NoSuchElementException:
    print('沒有公文可以點收')
    driver.close()
    driver.switch_to.window(driver.window_handles[0])
    driver.close()
    sys.exit()


# In[174]:


def check_exists_by_id():
    try:
        if driver.find_element_by_css_selector('#labdDecisionDocMark > div > span').get_attribute('innerHTML')  == "決" and driver.find_element_by_css_selector('#labdDecisionDocMark').value_of_css_property('z-index') == '6' :
            return True
        else:
            return False
    except NoSuchElementException:
        return False


# In[175]:


t = 1
for x in range (1,int(pageSize)+1):
    target = '#listTBODY > tr:nth-child('+ str(x) +') > td:nth-child(12) > span'
    driver.find_element_by_css_selector(target).click()
    child2 = driver.window_handles[2]
    driver.switch_to.window(child2)
    driver.implicitly_wait(30)
    check = check_exists_by_id()
    driver.switch_to.window(driver.window_handles[1])
    driver.switch_to.frame('dTreeContent')
    if check == True:
        driver.find_element_by_xpath('//*[@id="listTBODY"]/tr['+ str(x) +']/td[1]/div/ins').click()
        print('第 '+str(t)+' 筆公文checked')
        t += 1
    else :
        print('第 '+str(t)+' 筆公文無法checked!!!!!!!!!!!!!!')
        print('')
        t += 1


# In[ ]:


driver.switch_to.window(child2)
driver.close()
driver.switch_to.window(driver.window_handles[1])
driver.switch_to.frame('dTreeContent')
driver.find_element_by_xpath('//*[@id="functionMenuContainer"]/span[1]/input').click()
sleep(5)
driver.find_element_by_xpath('//*[@id="toSaveCheck"]').click()


# In[ ]:


def auto_refresh():
    try:
        qw=driver.window_handles[2]
        driver.switch_to.window(qw)
    except IndexError:
        print("No refresh")
        return
    except NoSuchWindowException: 
            print("No refresh")
            return
    try :
        print(driver.current_url)
    except NoSuchWindowException :
        print("No refresh")          
        return
    try:       
        driver.current_url =="http://localhost:16888/doPostMsg" 
        qw=driver.window_handles[2]
        driver.switch_to.window(qw)
        if driver.current_url =="http://localhost:16888/doPostMsg": 
            try:
                driver.refresh()
                sleep(3)
                print('已重新整理')
            except NoSuchWindowException:
                return
        else:
            print("double checking didn't confirm")
    except NoSuchWindowException:
        print("No refresh")


# In[ ]:


def break_point():
    driver.switch_to.window(driver.window_handles[1])
    driver.switch_to.frame('dTreeContent')
    try:
        driver.find_element_by_id('iframe01')
        return True
    except NoSuchElementException:
        return False


# In[ ]:


while(break_point()):
    sleep(3)
    auto_refresh()


# In[154]:


sys.exit()

