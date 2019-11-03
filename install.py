import os
import shutil
import platform

'''
В папке с установщиком должны находится файлы:
    monyze_config.ini
    monyze.exe
    OpenHardwareMonitorLib.dll
    keys.key
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
    is_default_path = input('Do you want to install to ' + install_path + ' ? (y/n)')
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
with open(os.path.join(workdir, 'keys.key')) as f:
    for line in f:
        lines.append(line)
lines[1] += '\n' #В keys.key не добавляется новая строка в конце второго ключа, делаем это тут
lines.append(install_path)
new_config = os.path.join(install_path, 'monyze_config.ini')#Ссыла на новый конфиг
with open(new_config, 'w') as f:
    for line in lines:
        f.write(line)

#Копируем библиотеку и exe в новую директорию. Определяем версию windows и копируем в системную попку конфиг
windir = os.environ['WINDIR']
is_64bits = platform.architecture()[0].find('64') != -1
if is_64bits:
    sys_directory = windir + r'\system32'
else:
    sys_directory = windir + r'\sysWOW64'


try:
    shutil.copy(ohwm_dll, install_path)
    shutil.copy(monyze_exe, install_path)
    shutil.copy(new_config, sys_directory)
except:
    print('Service is work now. Stopping it.')
    os.chdir(install_path)
    os.system('monyze.exe' + ' stop')
    try:
        shutil.copy(new_config, sys_directory)
        shutil.copy(monyze_exe, install_path)
    except:
        print('Cant copy config or monyze.exe file')

#Устанавливаем и запускаем службу
new_monyze_exe = os.path.join(install_path, 'monyze.exe')
os.chdir(install_path)
os.system('monyze.exe' + ' remove')
os.system('monyze.exe' + ' --startup auto install')
os.system('monyze.exe' + ' start')

