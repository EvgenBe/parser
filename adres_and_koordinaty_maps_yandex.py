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
openfilename = tk.askopenfilename()                     #диалог выбора файла с адресами
print('выбран файл '+openfilename)
print('загружаю данные...') 

sheet1 = pd.read_excel(openfilename)
print(sheet1)


url='https://yandex.ru/maps'
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

kn=0
while kn < len(sheet1):
    if kn % 50==0: 
        while 1:
            try:
                driver = webdriver.Chrome(ChromeDriverManager().install(),chrome_options=options)#,desired_capabilities=caps
                time.sleep(2)
                break
            except Exception as exc:
                print('браузер не запустился')
                time.sleep(5)


    stroka=pd.DataFrame(sheet1.iloc[kn].values)
    adress_xls=str(stroka.iat[0,0])  
    print('-----------------------')
    print(str(kn+1)+' из '+str(len(sheet1)))
    print(adress_xls)

    driver.get(url)                                                                                     #открытие карты
    found1=len(driver.find_elements(By.XPATH,'//input[@placeholder="Поиск мест и адресов"]'))           #проверка загрузки панели поиска
    #print(found1)                                           
    while found1==0:                                                                                    #ждать пока не прогрузится панель поиска
        time.sleep(1)
        found1=len(driver.find_elements(By.XPATH,'//input[@placeholder="Поиск мест и адресов"]'))
    driver.find_element(By.XPATH,'//input[@placeholder="Поиск мест и адресов"]').send_keys(adress_xls+'\n')     #ввод адреса для поиска

    load=1

    found=len(driver.find_elements(By.XPATH,'//div[@class="toponym-card-title-view__description"]'))            #проверка загрузки результата поиска
    while found==0 and load<5:
        time.sleep(1)
        load=load+1
        print(load)
        found=len(driver.find_elements(By.XPATH,'//div[@class="toponym-card-title-view__description"]'))            #проверка загрузки результата поиска
    print('найдено точных результатов '+str(found))

    found2=len(driver.find_elements(By.XPATH,'//li[@class="search-snippet-view"]'))                             #проверка если несколько результатов
    print('найдено альтернативных результатов '+str(found2))

    while found==0 and found2==0:
        time.sleep(1)
        found=len(driver.find_elements(By.XPATH,'//div[@class="toponym-card-title-view__description"]'))            #проверка загрузки результата поиска
        print('найдено точных результатов '+str(found))
        found2=len(driver.find_elements(By.XPATH,'//li[@class="search-snippet-view"]'))                             #проверка если несколько результатов
        print('найдено альтернативных результатов '+str(found2))

        if len(driver.find_elements(By.XPATH,'//div[@class="route-form-view__points"]'))>0:
            find_adress='не найдено'
            find_koord='не найдено'
            break

        if len(driver.find_elements(By.XPATH,'//div[@class="card-title-view__wrapper _view_normal"]'))>0:

            if len(driver.find_elements(By.XPATH,'//meta[@itemprop="address"]'))==0:
                find_adress='не найдено'
                find_koord='не найдено'
                break

            find_adress1_0=driver.find_element(By.XPATH,'//meta[@itemprop="address"]').get_attribute("content")
            find_adress2_0=driver.find_element(By.XPATH,'//h1[@class="card-title-view__title"]').text
            find_adress=str(find_adress1_0)+', '+str(find_adress2_0)
            print(find_adress)
            find_koord=driver.find_element(By.XPATH,'//div[@class="search-placemark-view _position_right"]').get_attribute("data-coordinates")
            print(find_koord)
            break


        if len(driver.find_elements(By.XPATH,'//div[@class="card-title-view__wrapper"]'))>0:
            find_adress1_0=driver.find_element(By.XPATH,'//meta[@itemprop="address"]').get_attribute("content")
            find_adress2_0=driver.find_element(By.XPATH,'//h1[@class="card-title-view__title"]').text
            find_adress=str(find_adress1_0)+', '+str(find_adress2_0)
            print(find_adress)
            if len(driver.find_elements(By.XPATH,'//div[@class="search-placemark-view _position_right"]'))==0:
                koordinaty=driver.find_element(By.XPATH,'//meta[@property="og:image"]').get_attribute("content")
                print(koordinaty)
                koordinaty=re.search(r"\d{2}\.\d{2,}.*\d{2}\.\d{2,}",koordinaty).group(0)
                print(koordinaty)
                koordinaty = re.sub(r"%.*?[^\d{2}]", ",", koordinaty, 0)
                print(koordinaty)
                find_koord=koordinaty
                break
            else: 
                find_koord=driver.find_element(By.XPATH,'//div[@class="search-placemark-view _position_right"]').get_attribute("data-coordinates")
                print(find_koord)
                break


        if len(driver.find_elements(By.XPATH,'//*[.="Ничего не найдено."]')):
            print('Не найдено')
            find_adress='Не найдено'
            find_koord='Не найдено'
            break

    if found==1:
        find_adress=driver.find_element(By.XPATH,'//div[@class="toponym-card-title-view__description"]').text
        print(find_adress)

        find_koord=driver.find_element(By.XPATH,'//div[@class="toponym-card-title-view__coords-badge"]').text
        print(find_koord)


    elif found2>0:
        driver.find_element(By.XPATH,'//li[@class="search-snippet-view"]').click()                              #переход по первому из нескольких найденых

        find_adress1=len(driver.find_elements(By.XPATH,'//div[@class="business-contacts-view__address"]'))
        if find_adress1>0:
            find_adress1=driver.find_element(By.XPATH,'//div[@class="business-contacts-view__address"]').text
            find_adress2=driver.find_element(By.XPATH,'//div[@class="card-title-view__wrapper _view_normal"]').text
            print(find_adress1)
            print(find_adress2)
            find_adress=str(find_adress1)+', '+str(find_adress2)
            find_koord=driver.find_element(By.XPATH,'//div[@class="search-placemark-view _position_right"]').get_attribute("data-coordinates")
            print(find_koord)

        find_adress1_1=len(driver.find_elements(By.XPATH,'//div[@class="toponym-card-title-view__description"]'))
        if find_adress1_1>0:
            find_adress=driver.find_element(By.XPATH,'//div[@class="toponym-card-title-view__description"]').text
            find_koord=driver.find_element(By.XPATH,'//div[@class="toponym-card-title-view__coords-badge"]').text
        
        
        print(find_adress)
        print(find_koord)

        
    #else:
    #    while found==0:
    #        time.sleep(1)
    #        found=len(driver.find_elements(By.XPATH,'//div[@class="toponym-card-title-view__description"]'))

    

    itog_info=adress_xls+'|'+find_adress+'|'+find_koord
    print(itog_info)
    with open ('yandex_OKS.txt','a',encoding='utf-8') as tmp:
        tmp.write(itog_info+'\n')
    kn=kn+1

    if kn % 50==0:
        driver.quit()