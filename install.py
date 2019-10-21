import os
import shutil

'''
В папке с установщиком должны находится файлы:
    monyze_config.ini
    monyze.exe
    OpenHardwareMonitorLib.dll
'''
#Пути к файлам в директории установочника
workdir = os.path.dirname(os.path.abspath(__file__))
ohwm_dll = os.path.join(workdir, 'OpenHardwareMonitorLib.dll')
monyze_exe = os.path.join(workdir, 'monyze.exe')
monyze_config = os.path.join(workdir, 'monyze_config.ini')


print('Hey! Welcome to install Monyze Agent!')
install_path = os.path.join(os.environ['PROGRAMFILES'], 'Monyze')
install_path = os.path.normpath(install_path)

#Определяем директорию для установки (по умолчанию в program files\Monyze), создаем новую, если необходимо
while True:
    is_default_path = input('Do you want to install to ' + install_path + ' (y/n)')
    if is_default_path.lower() != 'y':
        user_path = input('Input directory name:')
        user_path = os.path.normpath(user_path)
        if os.path.exists(user_path):
            install_path = os.path.abspath(user_path)
            break
        else:
            try:
                os.makedirs(user_path)
                install_path = os.path.abspath(user_path)
                break
            except:
                print('Incorrect directory name')
    else:
        if not os.path.exists(install_path):
            os.makedirs(install_path)
        break
print('Installing to: ' + install_path)

#Перезаписываем конфиг файл с новой установочной директорией (перезаписываем сразу в место установки)
lines = []
with open(monyze_config) as f:
    for line in f:
        lines.append(line)
lines[2] = install_path
with open(os.path.join(install_path, 'monyze_config.ini'), 'w') as f:
    for line in lines:
        f.write(line)

#Копируем библиотеку и exe в новую директорию
shutil.copy(ohwm_dll, install_path)
shutil.copy(monyze_exe, install_path)

#Устанавливаем и запускаем службу
new_monyze_exe = os.path.join(install_path, 'monyze.exe')
os.chdir(install_path)
os.system('monyze.exe' + ' remove')
os.system('monyze.exe' + ' --startup auto install')
os.system('monyze.exe' + ' start')

