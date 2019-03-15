# -*- coding: utf-8 -*-

import subprocess
import os
import requests

workdir = os.path.dirname(os.path.abspath(__file__))
config = os.path.join(workdir,"config.ini")
version = os.path.join(workdir,"version.txt")

f = open(config, 'r')
id_user = f.readline().rstrip('\n')
id_computer = f.readline().rstrip('\n')
work_path = f.readline().rstrip('\n')
f.close()

def stop_service():
    try:
        print ("Останавливаем службу")
        subprocess.check_call([work_path+"\monyze_dev.exe", "stop"])
    except:
        print("Что то пошло не так")

def start_service():
    try:
        print ("Запускаем службу")
        subprocess.check_call([work_path+"\monyze_dev.exe", "start"])
    except:
        print("Что то пошло не так")

def remove_service():
    try:
        print ("Удаляем старую версию")
        subprocess.check_call([work_path+"\monyze_dev.exe", "remove"])
    except:
        print("Что то пошло не так")

def install_service():
    try:
        print ("Устанавливаем новую версию")
        subprocess.check_call([work_path+"\monyze_dev.exe", "install"])
    except:
        print("Что то пошло не так")

def update():
    print("Запуск процесса обновления")
    try:
        f = open(version, 'r')
        this_version = f.readline().rstrip('\n')
        f.close()
    except IOError:
        print("Невозможно открыть файл version.txt")
        exit()

    version_url = 'https://dev.monyze.ru/files/windows/version.txt'
    try:        
        f = requests.get(version_url)
        new_version = float(f.text)

    except IOError:
        print("Невозможно проверить новую версию")

    if(new_version > float(this_version)):
        whats_new_url ='https://dev.monyze.ru/files/windows/whats_new.txt'
        wn = requests.get(whats_new_url)
        whats_new = wn.text       
        print("Доступно обновление агента, текущая версия: "+this_version+" новая версия: "+str(new_version))
        print ("Что нового: " +whats_new)
        print ("Скачать обновление? (y/n)")
        answer = input()
        if (answer=='y' or answer=='yes'):
            try:
                stop_service()
                remove_service()
                print("Загрузка новой версии")
                f=open(work_path+'\monyze_dev.exe',"wb")
                ufr = requests.get("https://dev.monyze.ru/files/windows/monyze_dev.exe")
                f.write(ufr.content) 
                f.close()    
                install_service()
                start_service()
            except:
                print("Что то пошло не так")
        else:
            print('Обновление отменено')
    else:
        print('Установлена последняя версия, обновление не требуется')

update()
