from turtle import goto
import requests
from bs4 import BeautifulSoup
import time
import pandas as pd
import re
from cgitb import text
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.common.by import By
import tkinter.filedialog as tk
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
print('выбери файл для загрузки')
openfilename = tk.askopenfilename()                                                             #диалог выбора файла
print('выбран файл '+openfilename)
print('загружаю данные...') 
open=re.search(r"\d{2}",openfilename).group(0)
sheet1 = pd.read_excel(openfilename)
print(sheet1)

table1=pd.DataFrame(columns=['кадном','адрес'])
k='отработало без ошибок'
k1='сервер не ругался'

kv=0
while kv < len(sheet1):
    stroka=pd.DataFrame(sheet1.iloc[kv].values)
    kvartal=str(stroka.iat[0,0])  
    print('-----------------------')
    print(str(kv+1)+' из '+str(len(sheet1)))
    print(kvartal)

    url="https://egrp365.org/list5.php?id="+str(kvartal)+":"+"*"
    print(url)
    print('квартал '+str(kvartal)+":"+'*')
    html=requests.get(url, headers={'user-agent': 'my-app/0.0.1'})#.status_code
    while html.status_code!=200:
        print(html.status_code)
        print('сервер не доволен')
        k1='сервер был не доволен, но я справился'
        html=requests.get(url, headers={'user-agent': 'my-app/0.0.1'})#.status_code
        time.sleep(5)
    soup = BeautifulSoup(html.text, 'html.parser')
    найдено=len(soup.findAll('div', class_="rs_element_item"))
    print('найдено в квартале '+str(найдено))
    if найдено==0:
        print('номеров в квартале нет')
    else:
        if найдено<200:
            i=0
            while i<найдено:
                kadnum=soup.findAll('div', class_="rs_kad_number")[i].get_text()
                print(kadnum)
                adress=soup.findAll('div', class_="rs_address")[i].get_text()
                print(adress)
                adress=adress.replace('\n','')
                table=pd.DataFrame(columns=['кадном','адрес'])
                table.loc['кадном']=kadnum
                table['адрес']=adress
                table1=pd.concat([table1,table])
                i=i+1
            #kv=kv+1
            #table1.to_excel(Sohranenie,index=False)
            #Sohranenie.save()
            table1.to_csv('D:\егрп\весрия с кварталами '+open+'.cs_', index=False,sep='#')
            #kv6=kv6+1
        if найдено>199:    
            nomer1=1
            while nomer1<10:
                print('перебор номеров в квартале')
                url="https://egrp365.org/list5.php?id="+str(kvartal)+":"+str(nomer1)+"*"
                print(url)
                print(str(kvartal)+":"+str(nomer1)+"*")
                html=requests.get(url, headers={'user-agent': 'my-app/0.0.1'})#.status_code
                while html.status_code!=200:
                    print('сервер не доволен')
                    k1='сервер был не доволен, но я справился'
                    html=requests.get(url, headers={'user-agent': 'my-app/0.0.1'})#.status_code
                    time.sleep(2)
                #html=requests.get(url, headers={'user-agent': 'my-app/0.0.1'}).text
                soup = BeautifulSoup(html.text, 'html.parser')
                найдено=len(soup.findAll('div', class_="rs_element_item"))
                print('найдено '+str(найдено))
                if найдено<200:
                    i=0
                    while i<найдено:
                        kadnum=soup.findAll('div', class_="rs_kad_number")[i].get_text()
                        print(kadnum)
                        adress=soup.findAll('div', class_="rs_address")[i].get_text()
                        print(adress)
                        adress=adress.replace('\n','')
                        table=pd.DataFrame(columns=['кадном','адрес'])
                        table.loc['кадном']=kadnum
                        table['адрес']=adress
                        table1=pd.concat([table1,table])
                        i=i+1
                    #table1.to_excel(Sohranenie,index=False)
                    #Sohranenie.save()
                    table1.to_csv('D:\егрп\весрия с кварталами '+open+'.cs_', index=False,sep='#')
                    nomer1=nomer1+1
                
                
                if найдено>199:
                    print('перебор первых номеров если их более 200')
                    if nomer1==1:
                        num100=0
                    else:
                        num100=int(str(nomer1-1)+'0')
                    while num100<int(str(nomer1)+'1'):
                        url="https://egrp365.org/list5.php?id="+str(kvartal)+":"+str(num100)
                        print(url)
                        print(str(kvartal)+":"+str(num100))
                        html=requests.get(url, headers={'user-agent': 'my-app/0.0.1'})#.status_code
                        while html.status_code!=200:
                            print('сервер не доволен')
                            k1='сервер был не доволен, но я справился'
                            html=requests.get(url, headers={'user-agent': 'my-app/0.0.1'})#.status_code
                            time.sleep(2)
                        #html=requests.get(url, headers={'user-agent': 'my-app/0.0.1'}).text
                        soup = BeautifulSoup(html.text, 'html.parser')
                        найдено=len(soup.findAll('div', class_="rs_element_item"))
                        print('найдено при переборе 100: ' +str(найдено))
                        if найдено>0:
                            kadnum=soup.find('div', class_="rs_kad_number").get_text()
                            print(kadnum)
                            adress=soup.find('div', class_="rs_address").get_text()
                            print(adress)
                            adress=adress.replace('\n','')
                            table=pd.DataFrame(columns=['кадном','адрес'])
                            table.loc['кадном']=kadnum
                            table['адрес']=adress
                            table1=pd.concat([table1,table])
                            #Sohranenie = pd.ExcelWriter('D:\егрп\\'+mun+'.xlsx')
                            #table1.to_excel(Sohranenie,index=False)
                            #Sohranenie.save()
                            num100=num100+1
                        else:
                            print('номера нет')
                            num100=num100+1
                    print('просмотр номеров если более 200 после перебора первых')
                    nomer2=0
                    while nomer2<10:
                        url="https://egrp365.org/list5.php?id="+str(kvartal)+":"+str(nomer1)+str(nomer2)+"*"
                        print(url)
                        print(str(kvartal)+":"+str(nomer1)+str(nomer2)+"*")
                        html=requests.get(url, headers={'user-agent': 'my-app/0.0.1'})#.status_code
                        while html.status_code!=200:
                            print('сервер не доволен')
                            k1='сервер был не доволен, но я справился'
                            html=requests.get(url, headers={'user-agent': 'my-app/0.0.1'})#.status_code
                            time.sleep(2)
                        #html=requests.get(url, headers={'user-agent': 'my-app/0.0.1'}).text
                        soup = BeautifulSoup(html.text, 'html.parser')
                        найдено=len(soup.findAll('div', class_="rs_element_item"))
                        print('найдено если более 200: ' +str(найдено))
                        if найдено<200:
                            i=0
                            while i<найдено:
                                kadnum=soup.findAll('div', class_="rs_kad_number")[i].get_text()
                                print(kadnum)
                                adress=soup.findAll('div', class_="rs_address")[i].get_text()
                                print(adress)
                                adress=adress.replace('\n','')
                                table=pd.DataFrame(columns=['кадном','адрес'])
                                table.loc['кадном']=kadnum
                                table['адрес']=adress
                                table1=pd.concat([table1,table])
                                i=i+1
                            #table1.to_excel(Sohranenie,index=False)
                            #Sohranenie.save() 
                            table1.to_csv('D:\егрп\весрия с кварталами '+open+'.cs_', index=False,sep='#')
                            nomer2=nomer2+1
                        if найдено>199:
                            k='пиши сюда еще кусок кода'
                            print(k)
                        if найдено==0:
                            print('ничего нет')
                            nomer2=nomer2+1
                    nomer1=nomer1+1
    kv=kv+1
    print(kv)
print(k)
print(k1)
table1.drop_duplicates(subset=['кадном','адрес'], inplace=True)
#table1.to_excel(Sohranenie,index=False)    
#Sohranenie.save()
table1.to_csv('D:\егрп\весрия с кварталами '+open+'.cs_', index=False,sep='#')
print('Я закончил')                    
        