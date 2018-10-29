#!/usr/bin/env python
# coding: utf-8

# In[397]:


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


# In[398]:


driver = webdriver.Chrome("C:\\Users\\TFD\\Anaconda3\\chromedriver.exe")
wait = WebDriverWait(driver,10)


# In[399]:


driver.maximize_window()
driver.get("https://ssopxy.gov.taipei/SSOPXY/Default.aspx")
id =('****')
password = ('****')


# In[400]:


driver.find_element_by_name('TextBox_userID').send_keys(id)


# In[401]:


driver.find_element_by_name('TextBox_userPwd').send_keys(password)


# In[402]:


driver.find_element_by_name('Button_Login').click()


# In[403]:


driver.find_element_by_class_name('category-button1').click()


# In[404]:


driver.find_element_by_css_selector('input[value=新公文系統]').click()


# In[405]:


child = driver.window_handles[1]      
driver.switch_to.window(child)
print(child)


# In[406]:


driver.implicitly_wait(10)
try:
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "toSaveCheck"))     
    )
finally:
    driver.find_element_by_xpath('//*[@id="toSaveCheck"]').click()


# In[407]:


sleep(2)
driver.find_element_by_id('menu1484').click()


# In[421]:


driver.switch_to.frame('dTreeContent')
driver.find_element_by_name('conditionQueryAll').click()


# In[422]:


sleep(2)
driver.find_element_by_xpath('//*[@id="qryModal"]/div[2]/div/div[2]/div/div/div/table/tbody/tr[4]/td[4]/select/option[3]').click()


# In[423]:


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


# In[424]:


driver.find_element_by_xpath('//*[@id="qryModal"]/div[2]/div/div[2]/div/div/div/table/tbody/tr[2]/td[2]/input[3]').clear()
driver.find_element_by_xpath('//*[@id="qryModal"]/div[2]/div/div[2]/div/div/div/table/tbody/tr[2]/td[2]/input[3]').send_keys(date)


# In[425]:


driver.find_element_by_name('querySubmit').click()


# In[427]:


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


# In[414]:


def check_exists_by_id():
    try:
        if driver.find_element_by_css_selector('#labdDecisionDocMark > div > span').get_attribute('innerHTML')  == "決" :
            return True
        else:
            return False
    except NoSuchElementException:
        return False


# In[ ]:


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


# In[415]:


driver.switch_to.window(child2)
driver.close()
driver.switch_to.window(driver.window_handles[1])
driver.switch_to.frame('dTreeContent')
driver.find_element_by_xpath('//*[@id="functionMenuContainer"]/span[1]/input').click()
sleep(3)
driver.find_element_by_xpath('//*[@id="toSaveCheck"]').click()


# In[416]:


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
            print('已重新整理')
            try:
                driver.refresh()
            except NoSuchWindowException:
                return
        else:
            print("double checking didn't confirm")
    except NoSuchWindowException:
        print("No refresh")


# In[432]:


def break_point():
    driver.switch_to.window(driver.window_handles[1])
    driver.switch_to.default_content()
    try:
        driver.find_element_by_id('boxMsg_s')
        return False
    except NoSuchElementException:
        return True


# In[434]:


while(break_point()):
    auto_refresh()


# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:




