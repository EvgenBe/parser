from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
options = Options()
import pandas as pd
import tkinter.filedialog as tk
import re
import time
import os
import shutil
import os.path
import traceback

print('выбери файл для загрузки')
openfilename = tk.askopenfilename()                                        #диалог выбора файла
print('выбран файл '+openfilename)
print('загружаю данные...') 
options = webdriver.ChromeOptions()                                           #запуск браузера
driver = webdriver.Chrome(ChromeDriverManager().install())
options.add_argument('--ignore-ssl-errors')
from http import cookies


folder=re.search(r".*/",openfilename).group(0)                 #путь к папке с файлом (для сохранения)
print (folder)
sheet1 = pd.read_excel(openfilename)                            #датафрейм из екселя
print(sheet1)
ssylki=pd.DataFrame(sheet1['Ссылка на объявление'])              #столбец с сылками
print(ssylki)
print (len(ssylki))
s=0
while s<len(ssylki):                                           #цикл перебора ссылок и сохранение пдф
    print('-------------')
    ok=driver.find_elements(By.XPATH,'//*[.="Работаем"]')
    if len(ok)==0:
        driver.get('Hello.html')                   #заглушка для обновления страницы
    stroka=pd.DataFrame(ssylki.iloc[s].values)
    print(str(s+1)+' из '+str(len(ssylki)))
    url=str(stroka.iat[0,0])
    print(url)
    url=url.replace('www','ivanovo')
    print('открываю сайт')
    url1=re.search(r".*\.ru/",url).group(0)
    url2=re.search(r"sale/\D+/\d+",url).group(0)
    url=url1+'export/pdf/'+url2
    print(url)
    driver.get(url)       #адрес сайта
    time.sleep(1)
    print('сайт открыт')
     
    filename=re.search(r"sale/\D+/\d+",url).group(0)
    filename=filename.replace('/','_')
        #filename=filename2+str(filename)+'.pdf'
    filename=filename+'.pdf'
    print(filename)
    
    neprogruz=driver.find_elements(By.XPATH,'//*[.="Страница не загрузилась"]')
    vpn=driver.find_elements(By.XPATH,'//*[.="Кажется, вам мешает VPN"]')
    while len(vpn)>0:
        print('Кажется, вам мешает VPN')
        while 1:
            try:
                driver.get(url)
                break
            except Exception as exc:
                #print(exc)
                traceback.print_exc()
                time.sleep(5)

        f=os.path.isfile(r'Downloads'+'\\'+filename)
        if f==True:
            break
        
        vpn=driver.find_elements(By.XPATH,'//*[.="Кажется, вам мешает VPN"]')
        time.sleep(3)
        f=os.path.isfile(r'Downloads'+'\\'+filename)
        if f==True:
            break
    while len(neprogruz)>0:
        print('страница не прогрузилась, пробую еще')
        f=os.path.isfile(r'Downloads'+'\\'+filename)
        if f==True:
            break
        driver.get(url)
        neprogruz=driver.find_elements(By.XPATH,'//*[.="Страница не загрузилась"]')
        time.sleep(3)
        f=os.path.isfile(r'Downloads'+'\\'+filename)
        if f==True:
            break
    f=os.path.isfile(r'Downloads'+'\\'+filename)                    #проверка наличия файла в папке
    soedinrnie=driver.find_elements(By.XPATH,'//*[.="Соединение прервано"]')
    dostup=driver.find_elements(By.XPATH,'//*[.="Не удается получить доступ к сайту"]')
    if len(soedinrnie)>0:
        driver.get(url)
        time.sleep(1)
    if len(dostup)>0:
        driver.get(url)
        time.sleep(1)
    while f==False:
        time.sleep(1)
        f=os.path.isfile(r'Downloads'+'\\'+filename)
        if len(vpn)>0:
            print('меняй IP')
    shutil.move(r'Downloads'+'\\'+filename,folder+filename)         #перенос файла из загрузок в папку с исходной екселькой
    s=s+1
driver.quit()
print('я закончил файл '+str(openfilename))