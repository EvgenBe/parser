#import requests
#from bs4 import BeautifulSoup
import time
import pandas as pd
import re
import traceback
from cgitb import text
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.common.by import By
import tkinter.filedialog as tk
#from selenium.common.exceptions import TimeoutException
#from selenium.webdriver.support.ui import WebDriverWait
#from selenium.webdriver.support import expected_conditions as EC
#from selenium.common.exceptions import NoSuchElementException
#from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
options = webdriver.ChromeOptions()                                                             #запуск браузера

#binary_yandex_driver_file = 'D:\\hell\\yandexdriver.exe' # path to YandexDriver

#driver = webdriver.Chrome(ChromeDriverManager().install())
options.add_argument('--ignore-ssl-errors')
options.add_argument('--ignore-certificate-errors')
print('выбери файл для загрузки')
openfilename = tk.askopenfilename()                                                             #диалог выбора файла
print('выбран файл '+openfilename)
print('загружаю данные...') 
#open=re.search(r"\d{2}",openfilename).group(0)
sheet1 = pd.read_excel(openfilename)
print(sheet1)

#Sohranenie = pd.ExcelWriter('D:\hell\проба кода\егрп\весрия с кварталами '+open+'.xlsx')
table1=pd.DataFrame(columns=['кадном','адрес'])
k='отработало без ошибок'
#k1='сервер не ругался'

caps = DesiredCapabilities().CHROME
#caps["pageLoadStrategy"] = "normal"  #  complete
#caps["pageLoadStrategy"] = "eager"  #  interactive
caps["pageLoadStrategy"] = "none"

##driver.find_element(By.XPATH,'//span[@class="select2-arrow"]').click()
##driver.find_element(By.XPATH,'//input[@aria-activedescendant="select2-result-label-41"]').send_keys('Ивановская область\n')
#del open
#with open (r'D:\\Новая папка\\url.txt','a',encoding='utf-8') as tmp:
#    tmp.write(url)

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
    kadnumber=str(stroka.iat[0,0])  
    print('-----------------------')
    print(str(kn+1)+' из '+str(len(sheet1)))
    print(kadnumber)

    url="https://publichnaya-kadastrovaya-karta.com/object?id="+kadnumber#37:24:40125:111
    print(url)
    #driver.get(url)

    while 1:  
        try:
            driver.get(url)
            break 
        except Exception as exc:
            print(exc)
            traceback.print_exc()
            time.sleep(5) 

    found=len(driver.find_elements(By.XPATH,'//div[@class="col-md-7 col-lg-7 col-sm-7 col-xs-12 pull-left"]'))
    print(found)
    notfound=1
    while notfound<10 and found==0:
            print ('не найдено. проход '+ str(notfound)+' из 10')
        #while found==0:
            while 1:  
                try:
                    driver.get(url)
                    break 
                except Exception as exc:
                    print(exc)
                    traceback.print_exc()
                    time.sleep(5) 
            time.sleep(2)
            #driver.get(url)
            found=len(driver.find_elements(By.XPATH,'//div[@class="col-md-7 col-lg-7 col-sm-7 col-xs-12 pull-left"]'))
            notfound=notfound+1
    if notfound>9 and found==0:
        print('нет такой страницы')
        info='нет такой страницы'
        all_info='не найдено'
        koordinaty='без координат'
        koordinaty_link='нет ссылки на координаты'
        with open ('D:\hell\publ-kd-karta\спарсено\не найдено ЗУ.txt','a',encoding='utf-8') as tmp:
            tmp.write(kadnumber+'\n')
        #kn=kn+1
    else:              
        info=driver.find_element(By.XPATH,'//div[@class="col-md-7 col-lg-7 col-sm-7 col-xs-12 pull-left"]').get_attribute('innerText')
        #info=info.replace('Дом на карте','').replace('Сведения на Нежилое помещение','').replace('Отчет об объекте недвижимости позволит быстро узнать достоверные данные о владельце (ФИО собственника) и проверить чистоту объекта недвижимости перед сделкой и существующие ограничения на регистрационные действия (залог банка, арест) для объекта недвижимости: Ивановская область, г. Иваново, ул. Лежневская, д. 4. Запрос выписки из Росреестра для получения информации о правах собственности, стоимости по кадастру и поиска другой информации в кратчайшие сроки.','')
        print(info)
        all_info=info.replace('\n','-')

        koordinaty_found=len(driver.find_elements(By.XPATH,'//a[@class="ymaps-logo-link ymaps-logo-link-ru"]'))
        if koordinaty_found>0:
            koordinaty_link=driver.find_element(By.XPATH,'//a[@class="ymaps-logo-link ymaps-logo-link-ru"]').get_attribute('href')
            print(koordinaty_link)
            koordinaty=re.search(r"\d{2}\.\d{2,}.*\d{2}\.\d{2,}",koordinaty_link).group(0)
            print(koordinaty)
            koordinaty = re.sub(r"%.*?[^\d{2}]", ",", koordinaty, 0)
            print(koordinaty)
        else:
            koordinaty='без координат'
    
    tip_nedvizh=re.search(r"Тип недвижимости\n.*",info)
    if tip_nedvizh==None:
        tip_nedvizh='Тип недвижимости|-'
        print(tip_nedvizh)
    else:
        tip_nedvizh=re.search(r"Тип недвижимости\n.*",info).group(0)
        info=info.replace(re.search(r"Тип недвижимости.*",info).group(0),'')
        tip_nedvizh=tip_nedvizh.replace('\n','|')
        print(tip_nedvizh)

    tip_obj=re.search(r"Тип объекта\n.*",info)
    if tip_obj==None:
        tip_obj='Тип объекта|-'
        print(tip_obj) 
    else:
        tip_obj=re.search(r"Тип объекта\n.*",info).group(0)
        info=info.replace(re.search(r"Тип объекта.*",info).group(0),'')
        tip_obj=tip_obj.replace('\n','|')
        print(tip_obj)    

    kadnum=re.search(r"Кадастровый номер\n.*",info)
    if kadnum==None:
        kadnum='Кадастровый номер|'+str(kadnumber)
        print(kadnum)
    else:
        kadnum=re.search(r"Кадастровый номер\n.*",info).group(0)
        info=info.replace(re.search(r"Кадастровый номер.*",info).group(0),'')
        kadnum=kadnum.replace('\n','|')
        print(kadnum)

    status=re.search(r"Статус объекта\n.*",info)
    if status==None:
        status='Статус объекта|-'
        print(status)
    else:
        status=re.search(r"Статус объекта\n.*",info).group(0)
        info=info.replace(re.search(r"Статус объекта.*",info).group(0),'')
        status=status.replace('\n','|')
        print(status)

    data_post=re.search(r"Дата постановки на.*учет\n.*",info)
    if data_post==None:
        data_post='Дата постановки на кадастровый учет|-'
        print(data_post)
    else:
        data_post=re.search(r"Дата постановки на.*учет\n.*",info).group(0)
        info=info.replace(re.search(r"Дата постановки на.*учет.*",info).group(0),'')
        data_post=data_post.replace('\n','|')
        print(data_post)

    kateg_zem=re.search(r"Категория.*\n.*",info)
    if kateg_zem==None:
        kateg_zem='Категория земель|-'
        print(kateg_zem)
    else:
        kateg_zem=re.search(r"Категория.*\n.*",info).group(0)
        info=info.replace(re.search(r"Категория.*",info).group(0),'')
        kateg_zem=kateg_zem.replace('\n','|')
        print(kateg_zem)
    
    ploshad=re.search(r"Площадь.*\n.*",info)
    if ploshad==None:
        ploshad='Площадь.|-'  
        print(ploshad)
    else:  
        ploshad=re.search(r"Площадь.*\n.*",info).group(0)
        info=info.replace(re.search(r"Площадь.*",info).group(0),'')
        ploshad=ploshad.replace('\n','|')
        print(ploshad)

    razr_isp=re.search(r"Разрешенное.*\n.*",info)
    if razr_isp==None:
        razr_isp='Разрешенное использование|-'
        print(razr_isp)
    else:
        razr_isp=re.search(r"Разрешенное.*\n.*",info).group(0)
        info=info.replace(re.search(r"Разрешенное.*",info).group(0),'')
        razr_isp=razr_isp.replace('\n','|')
        print(razr_isp)

    ed_izmer=re.search(r"Единица измерения.*\n.*",info)
    if ed_izmer==None:
        ed_izmer='Единица измерения.|-'
        print(ed_izmer)
    else:
        ed_izmer=re.search(r"Единица измерения.*\n.*",info).group(0)
        info=info.replace(re.search(r"Единица измерения.*",info).group(0),'')
        ed_izmer=ed_izmer.replace('\n','|')
        print(ed_izmer)

    kad_stoim=re.search(r"Кадастровая стоимость.*\n.*",info)
    if kad_stoim==None:
        kad_stoim='Кадастровая стоимость|-'
        print(kad_stoim)
    else:
        kad_stoim=re.search(r"Кадастровая стоимость.*\n.*",info).group(0)
        info=info.replace(re.search(r"Кадастровая стоимость.*",info).group(0),'')
        kad_stoim=kad_stoim.replace('\n','|')
        print(kad_stoim)

    data_vnesen_stoim=re.search(r"Дата внесения стоимости\n.*",info)
    if data_vnesen_stoim==None:
        data_vnesen_stoim='Дата внесения стоимости|-'
        print(data_vnesen_stoim)
    else:
        data_vnesen_stoim=re.search(r"Дата внесения стоимости\n.*",info).group(0)
        info=info.replace(re.search(r"Дата внесения стоимости.*",info).group(0),'')
        data_vnesen_stoim=data_vnesen_stoim.replace('\n','|')
        print(data_vnesen_stoim)

    data_utver_stoim=re.search(r"Дата утверждения стоимости\n.*",info)
    if data_utver_stoim==None:
        data_utver_stoim='Дата утверждения стоимости|-'
        print(data_utver_stoim)
    else:
        data_utver_stoim=re.search(r"Дата утверждения стоимости\n.*",info).group(0)
        info=info.replace(re.search(r"Дата утверждения стоимости.*",info).group(0),'')
        data_utver_stoim=data_utver_stoim.replace('\n','|')
        print(data_utver_stoim)

    data_opred_stoim=re.search(r"Дата определения стоимости\n.*",info)
    if data_opred_stoim==None:
        data_opred_stoim='Дата определения стоимости|-'
        print(data_opred_stoim)
    else:
        data_opred_stoim=re.search(r"Дата определения стоимости\n.*",info).group(0)
        info=info.replace(re.search(r"Дата определения стоимости.*",info).group(0),'')
        data_opred_stoim=data_opred_stoim.replace('\n','|')
        print(data_opred_stoim)
    
    adres=re.search(r"Адрес .*\n.*",info)
    if adres==None:
        adres='Адрес (местоположение)|-'
        print(adres)
    else:
        adres=re.search(r"Адрес .*\n.*",info).group(0)
        info=info.replace(re.search(r"Адрес .*",info).group(0),'')
        adres=adres.replace('\n','|')
        print(adres)

    tip=re.search(r".*Тип\n.*",info)
    if tip==None:
        tip='Тип|-'
        print(tip) 
    else:
        tip=re.search(r".*Тип\n.*",info).group(0)
        info=info.replace(re.search(r"Тип.*",info).group(0),'')
        tip=tip.replace('\n','|')
        print(tip)    

    etazh=re.search(r".*Этажность\n.*",info)
    if etazh==None:
        etazh='Этажность|-'
        print(etazh)
    else:
        etazh=re.search(r".*Этажность\n.*",info).group(0)
        info=info.replace(re.search(r".*Этажность.*",info).group(0),'')
        etazh=etazh.replace('\n','|')
        print(etazh) 

    podz_etazh=re.search(r".*Подземная этажность\n.*",info)
    if podz_etazh==None:
        podz_etazh='Подземная этажность|-'
        print(podz_etazh) 
    else:
        podz_etazh=re.search(r".*Подземная этажность\n.*",info).group(0)
        info=info.replace(re.search(r".*Подземная этажность.*",info).group(0),'')
        podz_etazh=podz_etazh.replace('\n','|')
        print(podz_etazh) 

    stena=re.search(r".*Материал стен\n.*",info)
    if stena==None:
        stena='Материал стен|-'
        print(stena)
    else:
        stena=re.search(r".*Материал стен\n.*",info).group(0)
        info=info.replace(re.search(r".*Материал стен.*",info).group(0),'')
        stena=stena.replace('\n','|')
        print(stena)    

    zaversh_stroit=re.search(r".*Завершение строительства\n.*",info)
    if zaversh_stroit==None:
        zaversh_stroit='Завершение строительства|-'
        print(zaversh_stroit)
    else:
        zaversh_stroit=re.search(r".*Завершение строительства\n.*",info).group(0)
        info=info.replace(re.search(r".*Завершение строительства.*",info).group(0),'')
        zaversh_stroit=zaversh_stroit.replace('\n','|')
        print(zaversh_stroit) 

    data_obnov_inf=re.search(r"Дата обновления информации\n.*",info)
    if data_obnov_inf==None:
        data_obnov_inf='Дата обновления информации|-'
        print(data_obnov_inf)
    else:
        data_obnov_inf=re.search(r"Дата обновления информации\n.*",info).group(0)
        info=info.replace(re.search(r"Дата обновления информации.*",info).group(0),'')
        data_obnov_inf=data_obnov_inf.replace('\n','|')
        print(data_obnov_inf) 

    predv_rasch_naloga=re.search(r"Предварительный расчет налога по кадастровой стоимости:\n.*",info)
    if predv_rasch_naloga==None:
        predv_rasch_naloga='Предварительный расчет налога по кадастровой стоимости:|-'
        print(predv_rasch_naloga) 
    else:
        predv_rasch_naloga=re.search(r"Предварительный расчет налога по кадастровой стоимости:\n.*",info).group(0)
        info=info.replace(re.search(r"Предварительный расчет налога по кадастровой стоимости:.*",info).group(0),'')
        predv_rasch_naloga=predv_rasch_naloga.replace('\n','|')
        print(predv_rasch_naloga) 

    invent_nomer_1=re.search(r":\nИнвентарный номер\n.*",info)
    if invent_nomer_1==None:
        invent_nomer_1='Инвентарный номер|-'
        print(invent_nomer_1) 
    else:
        invent_nomer_1=re.search(r":\nИнвентарный номер\n.*",info).group(0)
        info=info.replace(invent_nomer_1,'')
        invent_nomer_1=invent_nomer_1.replace('\n','|').replace(':|','')
        print(invent_nomer_1) 

    invent_nomer_2=re.search(r"Инвентарный номер\n.*",info)
    if invent_nomer_2==None:
        invent_nomer_2='Инвентарный номер|-'
        print(invent_nomer_2)
    else:
        invent_nomer_2=re.search(r"Инвентарный номер\n.*",info).group(0)
        info=info.replace(invent_nomer_2,'')
        invent_nomer_2=invent_nomer_2.replace('\n','|')
        print(invent_nomer_2) 

    uslov_nomer=re.search(r"Условный номер:\n.*",info)
    if uslov_nomer==None:
        uslov_nomer='Условный номер:|-'
        print(uslov_nomer)
    else:
        uslov_nomer=re.search(r"Условный номер:\n.*",info).group(0)
        info=info.replace(uslov_nomer,'')
        uslov_nomer=uslov_nomer.replace('\n','|')
        print(uslov_nomer)

    inoi_nomer=re.search(r"Иной номер:\n.*",info)
    if inoi_nomer==None:
        inoi_nomer='Иной номер:|-'
        print(inoi_nomer) 
    else:
        inoi_nomer=re.search(r"Иной номер:\n.*",info).group(0)
        info=info.replace(inoi_nomer,'')
        inoi_nomer=inoi_nomer.replace('\n','|')
        print(inoi_nomer)     

    itog_info=tip_nedvizh+'|'+kadnum+'|'+status+'|'+data_post+'|'+kateg_zem+'|'+razr_isp+'|'+ploshad+'|'+ed_izmer+'|'+kad_stoim+'|'+data_vnesen_stoim+'|'+data_utver_stoim+'|'+data_opred_stoim+'|'+adres+'|'+tip+'|'+etazh+'|'+podz_etazh+'|'+stena+'|'+zaversh_stroit+'|'+data_obnov_inf+'|'+predv_rasch_naloga+'|'+invent_nomer_1+'|'+invent_nomer_2+'|'+uslov_nomer+'|'+inoi_nomer+'|'+tip_obj+'|'+all_info+'|'+koordinaty_link+'|'+koordinaty

    with open ('text ЗУ.txt','a',encoding='utf-8') as tmp:
        tmp.write(itog_info+'\n')

    kn=kn+1
    #print('-+-+-+-+-+-+-+-+-+-+-')
    if kn % 50==0:
        driver.quit()
