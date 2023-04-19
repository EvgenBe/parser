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
openfilename = tk.askopenfilename()                                                             #диалог выбора файла
print('выбран файл '+openfilename)
print('загружаю данные...') 

sheet1 = pd.read_excel(openfilename)
print(sheet1)

url='https://pkk.rosreestr.ru'
options = webdriver.ChromeOptions()
#options.headless=True
options.add_argument('--ignore-ssl-errors')
driver = webdriver.Chrome(ChromeDriverManager().install()
#, options=options
)

print('открываю сайт')
driver.get(url)       #адрес сайта
print('сайт открыт')
time.sleep(10)

table1=pd.DataFrame(columns=['Кад.номер','Вид','площадь','категория','Разрешенное использование','Адрес','Наименование','Назначение','стоимость'])
i=0
while i < len(sheet1):
    stroka=pd.DataFrame(sheet1.iloc[i].values)
    kod=str(stroka.iat[0,0])  
    print('-----------------------')
    print(str(i+1)+' из '+str(len(sheet1)))
    print(kod)
    
    #kod=''
    vid=''
    pl=''
    kat=''
    ri=''
    adr=''
    naim=''
    nazn=''
    stoim=''
    #kod='37:24:040607:88'
    #time.sleep(3)

    if len(driver.find_elements(By.XPATH,'//*[.=" Пропустить обучение "]'))>0:           #пропустить обучение       
        driver.find_element(By.XPATH,'//button[@class="tutorial-button-outline"]').click()           #пропустить обучение
    table=pd.DataFrame(columns=['Кад.номер','Вид','площадь','категория','Разрешенное использование','Адрес','Наименование','Назначение','стоимость'])
    
    print('ищу номер '+str(kod))

        #смена выбора типа для поиска
    выбор=driver.find_element(By.XPATH,'//div[@class="drop-down-list-type"]')
    driver.execute_script("arguments[0].setAttribute('style','')",выбор)
    driver.find_element(By.XPATH,'//div[@class="drop-down-list-type"]/*[.="Участки "]').click()

    driver.find_element(By.XPATH,'//input[@placeholder="Найти объекты..."]').clear()
    driver.find_element(By.XPATH,'//input[@placeholder="Найти объекты..."]').send_keys(kod)   #ввод кодномера
    driver.find_element(By.XPATH,'//div[@title="Найти"]').click()


    #-----------------------------------------------------------------------------------------------------------------------------------------------
    #смена выбора типа для поиска
    #выбор=driver.find_element(By.XPATH,'//div[@class="drop-down-list-type"]')
    #driver.execute_script("arguments[0].setAttribute('style','')",выбор)
    #driver.find_element(By.XPATH,'//div[@class="drop-down-list-type"]/*[.="ОКС "]').click()
    #driver.find_element(By.XPATH,'//div[@class="drop-down-list-type"]/*[.="Участки "]').click()
    #-----------------------------------------------------------------------------------------------------------------------------------------------
    
    #qqq=driver.find_element(By.XPATH,'//ul[@aria-labelledby="dropdownMenu"]')
    #driver.execute_script("arguments[0].setAttribute('style','display: none;')",qqq)   
    time.sleep(3)
    
    if len(driver.find_elements(By.XPATH,'//*[.="Ничего не найдено"]'))>0:   #ничего не найдено
        выбор=driver.find_element(By.XPATH,'//div[@class="drop-down-list-type"]')
        driver.execute_script("arguments[0].setAttribute('style','')",выбор)
        driver.find_element(By.XPATH,'//div[@class="drop-down-list-type"]/*[.="ОКС "]').click()
        time.sleep(3)
    if len(driver.find_elements(By.XPATH,'//*[.="Ничего не найдено"]'))>0:   #ничего не найдено
        print('Номер не найден на сайте')
        vid='Номер не найден на сайте'
        #kat=''
        #ri=''
        #adr=''
        #naim=''
        #nazn=''
        #stoim=''
        #goto

    else:
        while len(driver.find_elements(By.XPATH,'//div[@class="detail-info-container"]'))==0:
            print('сайт тупит. делаю все что могу...')

            #driver.find_element(By.XPATH,"//ul[@aria-labelledby='dropdownMenu']").setAttribute('style', 'display: none;')
            qqq=driver.find_element(By.XPATH,'//ul[@aria-labelledby="dropdownMenu"]')
            #for select in qqq:
            driver.execute_script("arguments[0].setAttribute('style','display: none;')",qqq)
            #qqq.style="display: none;"
            ActionChains(driver)\
            .move_to_element(driver.find_element(By.XPATH,'//div[@class="info-item-container"]'))\
            .perform()
            time.sleep(10)
            driver.find_element(By.XPATH,'//div[@isfounditem="true"]').click()
            #driver.find_element(By.XPATH,'//div[@class="info-item-container highlighted"]').click()
            time.sleep(5)
            if len(driver.find_elements(By.XPATH,'//*[.=" Не удалось выполнить поиск. Повторите позднее "]'))>0:   #Не удалось выполнить поиск. Повторите позднее
                driver.find_element(By.XPATH,'//div[@title="Найти"]').click()
                time.sleep(3)        
        
    info=len(driver.find_elements(By.XPATH,'//div[@data-v-ea369eca]'))
    if info==0:
        i=i+1
    else:
        info=driver.find_element(By.XPATH,'//div[@data-v-ea369eca]').text
        #print('-------------')
        #info=str(info)
        #print(info)
        #print('-------------')
        vid=re.search(r"Вид.*\n.*",info)
        pl=re.search(r"Площадь.*\n.*",info)
        kat=re.search(r"Категория.*\n.*",info)
        ri=re.search(r"Разрешенное.*\n.*",info)
        adr=re.search(r"Адрес.*\n.*",info)
        naim=re.search(r"Наименование.*\n.*",info)
        nazn=re.search(r"Назначение.*\n.*",info)
        stoim=re.search(r"стоимость.*\n.*",info)
        #print(vid)
        #print(pl)
        #print(kat)
        #print(ri)
        #print(adr)
        #print(naim)
        #print(nazn)
        #print(stoim)


        if vid==None:
            print('Вид пустая')
        else:
            vid=vid.group(0)
            vid=vid.replace('Вид:\n','')

        if pl==None:
            print('Площадь пустая')
        else:
            pl=pl.group(0)
            pl=pl.replace('Площадь уточненная:\n','')
            pl=pl.replace('Площадь декларированная:\n','')
            pl=pl.replace('Площадь общая:\n','')
        if kat==None:
            print('Категория пустая')
        else:
            kat=kat.group(0)
            kat=kat.replace('Категория земель:\n','')
        if ri==None:
            print('Разрешенное пустая')
        else:
            ri=ri.group(0)
            ri=ri.replace('Разрешенное использование:\n','')
        if adr==None:
            print('Адрес пустая')
        else:
            adr=adr.group(0)
            adr=adr.replace('Адрес:\n','')
        if naim==None:
            print('Наименование пустая')
        else:
            naim=naim.group(0)
            naim=naim.replace('Наименование:\n','')
        if nazn==None:
            print('Назначение пустая')
        else:
            nazn=nazn.group(0)
            nazn=nazn.replace('Назначение:\n','')   
        if stoim==None:
            print('стоимость пустая')
        else:
            stoim=stoim.group(0)
            stoim=stoim.replace('стоимость:\n','') 
        #print(pl)
        #print(kat)
        #print(ri)
        #print(adr)
        #print(naim)
        #print(nazn)
        #print(stoim)
        i=i+1
    table.loc['Кад.номер']=kod
    table['Вид']=vid
    table['площадь']=pl
    table['категория']=kat
    table['Разрешенное использование']=ri
    table['Адрес']=adr
    table['Наименование']=naim
    table['Назначение']=nazn
    table['стоимость']=stoim

    #print(table)
    table1=pd.concat([table1,table])
    
    #print(table1)
    
    #savefile = tk.asksaveasfilename(filetypes=(("Excel files", "*.xlsx"),("All files", "*.*") ))     #диалог выбора папки для сохранения
    Sohranenie = pd.ExcelWriter('итог.xlsx')   
    print('сохраняю файл...') 
    table1.to_excel(Sohranenie,index=False)                                                         #запись новой таблицы в переменную
    Sohranenie.save()                                                                               #сохранение
    print('файл сохранен ')
    table1.to_csv('итог.cs_', index=False,sep='#')
    print('Работа с номером '+str(kod)+' завершена')
       
    #time.sleep(3)
    
driver.quit()
print('Работа завершена')