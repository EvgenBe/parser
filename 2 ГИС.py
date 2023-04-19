from cgitb import text
from turtle import goto
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
import time
from selenium.webdriver.common.by import By
import pandas as pd
import re
import tkinter.filedialog as tk
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
print('выбери файл для загрузки')
openfilename = tk.askopenfilename()                                     #диалог выбора файла
print('выбран файл '+openfilename)
print('загружаю данные...') 

sheet1 = pd.read_excel(openfilename)
print(sheet1)


url='https://2gis.ru/ivanovo'
options = webdriver.ChromeOptions()
#options.headless=True
options.add_argument('--ignore-ssl-errors')
#driver = webdriver.Chrome(ChromeDriverManager().install()
##, options=options
#)
#
#print('открываю сайт')
#driver.get(url)       #адрес сайта
#print('сайт открыт')
#time.sleep(5)

table1=pd.DataFrame(columns=['Кад.номер','адрес','этажи','год постройки','Материал стен','Перекрытия'])
i=0
while i < len(sheet1):
    if len(table1) % 300==0: 
        driver = webdriver.Chrome(ChromeDriverManager().install())
        driver.get(url)
        time.sleep(2)


    stroka=sheet1.iloc[i]
    stroka=pd.DataFrame(stroka)
    print(str(i+1)+' из '+str(len(sheet1)))
    #print(stroka)
    stolb1=stroka.iat[0,0]
    stolb2=stroka.iat[1,0]
    print(stolb1)
    print(stolb2)
    print('-----------------------')
    
    
    #print(kod1)
    #print(kod2)
    table=pd.DataFrame(columns=['Кад.номер','адрес','этажи','год постройки','Материал стен','Перекрытия'])

    
    driver.find_element(By.XPATH,'//input[@placeholder="Поиск в 2ГИС"]').send_keys(stolb1+'\n')   #ввод кодномера
    time.sleep(3)
    notfound=len(driver.find_elements(By.XPATH,'//*[.="Точных совпадений нет. Посмотрите похожие места или измените запрос."]')) 
    polomalos=len(driver.find_elements(By.XPATH,'//*[.="Что-то пошло не так"]'))
    notfound2=len(driver.find_elements(By.XPATH,'//*[.="Ничего не нашлось, попробуйте уточнить запрос"]'))
    #time.sleep(2)
    if notfound>0:
        etazh='Точных совпадений нет.'
        god=''
        stena=''
        perekrytiya=''
        print('Точных совпадений нет.')
        driver.find_element(By.XPATH,'//button[@aria-label="Очистить локальное поле поиска"]').click()
        i=i+1
    elif polomalos>0:
        #driver.find_element(By.XPATH,'//button[@class="_g55hbm9"]').click()
        etazh='Что-то пошло не так'
        god=''
        stena=''
        perekrytiya=''
        print('Что-то пошло не так')
        driver.find_element(By.XPATH,'//button[@aria-label="Очистить локальное поле поиска"]').click()
        i=i+1
    elif notfound2>0:
        etazh='Ничего не нашлось, попробуйте уточнить запрос'
        god=''
        stena=''
        perekrytiya=''
        print('Ничего не нашлось, попробуйте уточнить запрос')
        driver.find_element(By.XPATH,'//button[@aria-label="Очистить локальное поле поиска"]').click()
        i=i+1
    else:
        стрнедоступна=len(driver.find_elements(By.XPATH,'//*[.="Страница недоступна"]'))
        while стрнедоступна>0:
            driver.find_element(By.XPATH,'//button[@id="reload-button"]').click()
            time.sleep(2)
            стрнедоступна=len(driver.find_elements(By.XPATH,'//*[.="Страница недоступна"]'))
        #driver.find_element(By.XPATH,'//button[@aria-label="Очистить локальное поле поиска"]').click()
        driver.find_element(By.XPATH,'//div[@class="_1h3cgic"]').click()
        time.sleep(2)
        #find=driver.find_element(By.XPATH,'//div[@class="_49kxlr"]').text
        find=len(driver.find_elements(By.XPATH,'//div[@class="_599hh"]'))
        #print(find)
        if find==0:
            print('кривой адрес')
            etazh='кривой адрес'
            god=''
            stena=''
            perekrytiya=''
            i=i+1
        else:
            find=driver.find_element(By.XPATH,'//div[@class="_599hh"]').text
            #print(find)
            etazh=re.search(r"\d*\sэтаж.*",find)
            if etazh==None:
                etazh='нет информации'
                print('нет информации об этажах')
                #i=i+1
            else:
                etazh=etazh.group(0)
                print (etazh)
            
            
            god=re.search(r"Год.*\n*\d*",find)
            if god==None:
                god='нет информации о годе постройки'
            else:
                god=god.group(0)
                god=god.replace('Год постройки\n','')
        
            stena=re.search(r"Материал.*\n*.*",find)
            if stena==None:
                stena='нет информации о материале стен'
            else:
                stena=stena.group(0)
                stena=stena.replace('Материал стен\n','')

            perekrytiya=re.search(r"Перекрытия.*\n*.*",find)
            if perekrytiya==None:
                perekrytiya='нет информации о перекрытиях'
            else:
                perekrytiya=perekrytiya.group(0)
                perekrytiya=perekrytiya.replace('Перекрытия\n','')
        
            driver.find_element(By.XPATH,'//button[@aria-label="Очистить локальное поле поиска"]').click()
            i=i+1
            print('---------------')
    table.loc['Кад.номер']=stolb2
    table['адрес']=stolb1
    table['этажи']=etazh
    table['год постройки']=god
    table['Материал стен']=stena
    table['Перекрытия']=perekrytiya
    #print(table)
    table1=pd.concat([table1,table])
    print(table1)
    table1.to_csv('сайт_два_гис.cs_', index=False,sep='#')
    print('Работа с номером '+str(stolb2)+' завершена')
    if len(table1) % 300==0:
        driver.quit()
    #print(len(table1))        
print('работа закончена')