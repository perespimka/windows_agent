# -*- coding: utf-8 -*-

import socket
import sys
import win32event
import win32service
import win32serviceutil
import wmi
import time
import platform
import os
import multiprocessing
import uuid
import psutil
import requests
import json
import pprint
import netifaces
import traceback
import sys
import ifaddr
import clr
import ctypes
import logging
import logging.config
from uptime import uptime
import servicemanager

c = wmi.WMI()

# ??? Ниже в закаменченном фрагменте просто какой-то кортеж создается, на который нет ссылки.
'''
for diskDrive in c.query("SELECT * FROM Win32_DiskDrive"):
    diskDrive.Size, "\nDisk model: ", diskDrive.Model,  diskDrive.Status 
'''
try:
    import _winreg as winreg
except ImportError as err:
    try:
        import winreg
    except ImportError as err:
        pass

workdir = os.path.dirname(os.path.abspath(__file__))
config = os.path.join(workdir,"config.ini")

f = open(config, 'r')
id_user = f.readline().rstrip('\n')
id_computer = f.readline().rstrip('\n')
dll_path = f.readline().rstrip('\n')
f.close()

#logging
logConfig = {
        "version":1,
        "handlers":{
            "fileHandler":{
                "class":"logging.FileHandler",
                "formatter":"myFormatter",
                "filename":os.path.join(dll_path,"monyze-agent.log")
            }
        },
        "loggers":{
            "monyze":{
                "handlers":["fileHandler"],
                "level":"INFO",
                "level":"DEBUG",
            }
        },
        "formatters":{
            "myFormatter":{
                "format":"%(asctime)s - %(name)s - %(levelname)s - %(message)s"
            }
        }
    }

logging.config.dictConfig(logConfig)
logger = logging.getLogger("monyze.main")


ohwm_dll = os.path.join(dll_path,"OpenHardwareMonitorLib.dll")

logger.info ('------------------------------------- ')
logger.info ('Запуск агента')
logger.info ("Путь к библиотеке ohwm: "+ohwm_dll)

pp = pprint.PrettyPrinter(indent=4)
nodename = platform.node()
system = platform.system()
win32_ver = platform.win32_ver()
oss = platform.platform()
bits = platform.architecture()[0]
key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,r"Hardware\Description\System\CentralProcessor\0")
processor_brand = winreg.QueryValueEx(key, "ProcessorNameString")[0]
winreg.CloseKey(key)
get_cpu_model = processor_brand
mem = psutil.virtual_memory()
swap = psutil.swap_memory()
partitions = psutil.disk_partitions()
interfaces = netifaces.interfaces()
adapters = ifaddr.get_adapters()
net_stats = psutil.net_if_stats()

logger.info ('Переменные определены')

driveinfo = []
cpu_load = {}
cpu_cores = []

openhardwaremonitor_hwtypes = ['Mainboard', 'SuperIO', 'CPU','RAM', 'GpuNvidia', 'GpuAti', 'TBalancer', 'Heatmaster', 'HDD']
openhardwaremonitor_sensortypes = ['Voltage', 'Clock', 'Temperature', 'Load','Fan', 'Flow', 'Control', 'Level', 'Factor', 'Power', 'Data', 'SmallData']
cpu_load_sensornames = ['CPU Total', 'CPU Core', 'CPU Core #1', 'CPU Core #2','CPU Core #3', 'CPU Core #4', 'CPU Core #5', 'CPU Core #6', 'CPU Core #7', 'CPU Core #8']
cpu_temperature_sensornames = ['CPU Package', 'CPU Core #1', 'CPU Core #2', 'CPU Core #3','CPU Core #4', 'CPU Core #5', 'CPU Core #6', 'CPU Core #7', 'CPU Core #8']
mb_temperature_sensornames = ['Temperature #1', 'Temperature #2', 'Temperature #3', 'Temperature #4','Temperature #5', 'Temperature #6', 'Temperature #7', 'Temperature #8']
mb_fan_sensornames = ['Fan #1', 'Fan #2', 'Fan #3', 'Fan #4','Fan #5', 'Fan #6', 'Fan #7', 'Fan #8','Fan #9', 'Fan #10', 'Fan #11', 'Fan #12']

logger.info ('Инициализация ohwm')
def initialize_openhardwaremonitor():
    clr.AddReference(ohwm_dll)
    from OpenHardwareMonitor import Hardware
    handle = Hardware.Computer()
    handle.MainboardEnabled = True
    handle.CPUEnabled = True
    handle.RAMEnabled = True
    handle.GPUEnabled = True
    handle.HDDEnabled = True
    handle.Open()
    return handle

HardwareHandle = initialize_openhardwaremonitor()
logger.info ('ohwm инициализирован')

logger.info ('Получаем данные о дисках')
hddpos = 0
for pdisk in c.query("SELECT * FROM Win32_DiskDrive"): 
    #print(pdisk.DeviceID)
    lnames = []
    if pdisk.Size:
        hddpos += 1
        Size = round((int(pdisk.Size)/1024 ** 2), 2)
        strcl = pdisk.deviceID.replace("\\", "")
        pdiskcleared = strcl.replace(".PHYSICALDRIVE", "PhysicalDrive")
        for colpartition in c.query('ASSOCIATORS OF {Win32_DiskDrive.DeviceID="' + pdisk.deviceID + '"} WHERE AssocClass = Win32_DiskDriveToDiskPartition'):
            for logical_disk in c.query('ASSOCIATORS OF {Win32_DiskPartition.DeviceID="' + colpartition.DeviceID + '"} WHERE AssocClass = Win32_LogicalDiskToPartition'):
                lname = logical_disk.DeviceID
                if lname:
                    lnames.append(lname)
        l = {'hdd_'+str(hddpos)+'': {'name': pdisk.Model,'size': Size, 'LOGICAL': lnames}}
        #print(l['hdd_'+str(hddpos)+'']['LOGICAL'])
        driveinfo.append(l)

logger.info ('Данные о дисках получены')

def get_config_data(handle):
    logger.info ('Инициализируем конфигурацию устройства')
    cpuinfo = {}
    c_count = 0
    for i in handle.Hardware:
        i.Update()
        for sensor in i.Sensors:
            if sensor.Value:
                ###CPU'S###
                if sensor.Hardware.HardwareType == openhardwaremonitor_hwtypes.index('CPU'):
                    if sensor.Index == 0 and sensor.SensorType == 3:
                        c_count += 1
                        cpuinfo['cpu_'+str(c_count)] = sensor.Hardware.Name
                        break
                    if sensor.Index >= 0:
                        c_count += 1
                        cpuinfo['cpu_'+str(c_count)] = sensor.Hardware.Name
                        break

    net = {}
    for i, adapter in enumerate(adapters):
        try:
            model = adapter.nice_name
            nm = adapter.name.decode("utf-8")
            ip = "%s/%s" % (adapter.ips[0].ip, adapter.ips[0].network_prefix)
            reg = winreg.ConnectRegistry(None, winreg.HKEY_LOCAL_MACHINE)
            reg_key = winreg.OpenKey(
                reg, r"SYSTEM\CurrentControlSet\Control\Network\{4d36e972-e325-11ce-bfc1-08002be10318}")
            try:
                reg_subkey = winreg.OpenKey(reg_key, nm + r"\Connection")
                name = winreg.QueryValueEx(reg_subkey, "Name")[0]
            except FileNotFoundError:
                pass
            else:
                speed = net_stats[name][2]
                counters = psutil.net_io_counters(pernic=True)
                btx = counters[name][0]
                if btx > 0:
                    net['net_'+str((i)+1)] = {'model': model,'name': name, 'speed': speed, 'addr': ip}
        except KeyError:
            continue
    
    logger.info("Конфигурация устройства инициализирована")
    config_data = {
        "id": {
            "user_id": id_user,
            "device_id": id_computer
        },
        "state": "config",
        "device_config": {
            "device_name": nodename,
            "system": oss,
            "icon": "f17a",
            "cpu": cpuinfo,
            "ram": {
                "TotalPh": mem.total,
                "TotalVrt": 0,
                "TotalPF": swap.total,
                "bits": bits
            },
            "hdd": driveinfo,
            "net": net
        }
    }

    url = 'https://monyze.ru/api.php'
    try:    
        requests.post(url, json.dumps(config_data))
        logger.info('Конфигурация устройства:\n' + json.dumps(config_data, indent=4) + '\n ---------------')
        logger.info('Конфигурация устройства отправлена')
    except:
        logger.warning("Не удалось отправить конфигурацию устройства")
        logger.warning(traceback.format_exc())


def get_load_data(handle):
    memory = psutil.virtual_memory()
    cpu_count = 0

    for i in handle.Hardware:
        i.Update()

        ###CPU and Fans###
        if i.HardwareType == openhardwaremonitor_hwtypes.index('CPU'):
            if i.processorIndex == cpu_count:
                load_arr = {}
                temp_arr = {}
                cpu_load_widg_arr = []
                cpu_temp_widg_arr = []
                cpu_count += 1
                cpu_load_data = {}
                for sensor in i.Sensors:
                    if sensor.Value:
                        #CPU Load Sensors
                        if sensor.SensorType == openhardwaremonitor_sensortypes.index('Load'):
                            if sensor.Name == cpu_load_sensornames[0] or sensor.Name == cpu_load_sensornames[1]:
                                load_arr['total'] = round(sensor.Value)
                                cpu_load_widg_arr.append(round(sensor.Value))
                            elif sensor.Name == cpu_load_sensornames[2]:
                                load_arr['core_1'] = round(sensor.Value)
                            elif sensor.Name == cpu_load_sensornames[3]:
                                load_arr['core_2'] = round(sensor.Value)
                            elif sensor.Name == cpu_load_sensornames[4]:
                                load_arr['core_3'] = round(sensor.Value)
                            elif sensor.Name == cpu_load_sensornames[5]:
                                load_arr['core_4'] = round(sensor.Value)
                            elif sensor.Name == cpu_load_sensornames[6]:
                                load_arr['core_5'] = round(sensor.Value)
                            elif sensor.Name == cpu_load_sensornames[7]:
                                load_arr['core_6'] = round(sensor.Value)
                            elif sensor.Name == cpu_load_sensornames[8]:
                                load_arr['core_7'] = round(sensor.Value)
                            elif sensor.Name == cpu_load_sensornames[9]:
                                load_arr['core_8'] = round(sensor.Value)
                        #CPU Temperature sensors
                        if sensor.SensorType == openhardwaremonitor_sensortypes.index('Temperature'):
                            if sensor.Name == cpu_temperature_sensornames[0]:
                                temp_arr['total'] = round(sensor.Value)
                                cpu_temp_widg_arr.append(round(sensor.Value))
                            elif sensor.Name == cpu_temperature_sensornames[1]:
                                temp_arr['core_1'] = round(sensor.Value)
                                cpu_temp_widg_arr.append(round(sensor.Value))
                            elif sensor.Name == cpu_temperature_sensornames[2]:
                                temp_arr['core_2'] = round(sensor.Value)
                                cpu_temp_widg_arr.append(round(sensor.Value))
                            elif sensor.Name == cpu_temperature_sensornames[3]:
                                temp_arr['core_3'] = round(sensor.Value)
                                cpu_temp_widg_arr.append(round(sensor.Value))
                            elif sensor.Name == cpu_temperature_sensornames[4]:
                                temp_arr['core_4'] = round(sensor.Value)
                                cpu_temp_widg_arr.append(round(sensor.Value))
                            elif sensor.Name == cpu_temperature_sensornames[5]:
                                temp_arr['core_5'] = round(sensor.Value)
                                cpu_temp_widg_arr.append(round(sensor.Value))
                            elif sensor.Name == cpu_temperature_sensornames[6]:
                                temp_arr['core_6'] = round(sensor.Value)
                                cpu_temp_widg_arr.append(round(sensor.Value))
                            elif sensor.Name == cpu_temperature_sensornames[7]:
                                temp_arr['core_7'] = round(sensor.Value)
                                cpu_temp_widg_arr.append(round(sensor.Value))
                            elif sensor.Name == cpu_temperature_sensornames[8]:
                                temp_arr['core_8'] = round(sensor.Value)
                                cpu_temp_widg_arr.append(round(sensor.Value))
                        if sensor.SensorType == openhardwaremonitor_sensortypes.index('Clock'):
                            continue
                    cpu_load_data['load'] = load_arr
                    cpu_load_data['temp'] = temp_arr
                cpu_pos = 'cpu_'+str(cpu_count)+''
                cpu_load[cpu_pos] = cpu_load_data
                cpu_total_load_widg = round(
                    sum(cpu_load_widg_arr)/len(cpu_load_widg_arr))
                if len(cpu_temp_widg_arr) == 0:
                    cpu_total_temp_widg = 'n/a'
                else:
                    cpu_total_temp_widg = round(
                        sum(cpu_temp_widg_arr)/len(cpu_temp_widg_arr))
        
    ##Fans & MB Temperature
        if i.HardwareType == openhardwaremonitor_hwtypes.index('Mainboard'):
            mb_data = {}
            mb_temp_arr = {}
            mb_fan_arr = {}
            for SHardware in i.SubHardware:
                SHardware.Update()
                for sensor in SHardware.Sensors:
                    if sensor.Value:
                        if sensor.SensorType == openhardwaremonitor_sensortypes.index('Temperature'):
                            if sensor.Name == mb_temperature_sensornames[0]:
                                    mb_temp_arr['temp_1'] = round(sensor.Value)
                            elif sensor.Name == mb_temperature_sensornames[1]:
                                    mb_temp_arr['temp_2'] = round(sensor.Value)
                            elif sensor.Name == mb_temperature_sensornames[2]:
                                    mb_temp_arr['temp_3'] = round(sensor.Value)
                            elif sensor.Name == mb_temperature_sensornames[3]:
                                    mb_temp_arr['temp_4'] = round(sensor.Value)
                            elif sensor.Name == mb_temperature_sensornames[4]:
                                    mb_temp_arr['temp_5'] = round(sensor.Value)
                            elif sensor.Name == mb_temperature_sensornames[5]:
                                    mb_temp_arr['temp_6'] = round(sensor.Value)
                            elif sensor.Name == mb_temperature_sensornames[6]:
                                    mb_temp_arr['temp_7'] = round(sensor.Value)
                            elif sensor.Name == mb_temperature_sensornames[7]:
                                    mb_temp_arr['temp_8'] = round(sensor.Value)
                        if sensor.SensorType == openhardwaremonitor_sensortypes.index('Fan'):
                            if sensor.Name == mb_fan_sensornames[0]:
                                    mb_fan_arr['fan_1'] = round(sensor.Value)
                            elif sensor.Name == mb_fan_sensornames[1]:
                                    mb_fan_arr['fan_2'] = round(sensor.Value)
                            elif sensor.Name == mb_fan_sensornames[2]:
                                    mb_fan_arr['fan_3'] = round(sensor.Value)
                            elif sensor.Name == mb_fan_sensornames[3]:
                                    mb_fan_arr['fan_4'] = round(sensor.Value)
                            elif sensor.Name == mb_fan_sensornames[4]:
                                    mb_fan_arr['fan_5'] = round(sensor.Value)
                            elif sensor.Name == mb_fan_sensornames[5]:
                                    mb_fan_arr['fan_6'] = round(sensor.Value)
                            elif sensor.Name == mb_fan_sensornames[6]:
                                    mb_fan_arr['fan_7'] = round(sensor.Value)
                            elif sensor.Name == mb_fan_sensornames[7]:
                                    mb_fan_arr['fan_8'] = round(sensor.Value)
                            elif sensor.Name == mb_fan_sensornames[8]:
                                    mb_fan_arr['fan_9'] = round(sensor.Value)
                            elif sensor.Name == mb_fan_sensornames[9]:
                                    mb_fan_arr['fan_10'] = round(sensor.Value)

                mb_data['temp'] = mb_temp_arr
                mb_data['fans'] = mb_fan_arr 




    ram = {"load": round(memory.percent), "AvailPh": memory.available}

    ##HDD###
    hdd = {}
    hdd_widgets = {}
    hddpos = 0
    #print(driveinfo)
    #print('------------------------------------------------------------------------------------------------')
    for disk in driveinfo:
        hddpos += 1
        ldisks = []
        ldisks_widgets = []
        hdd_count = 'hdd_'+str(hddpos)+''
        for ld in disk[hdd_count]['LOGICAL']:
            if ld:
                usage = psutil.disk_usage(ld)
                perc = usage.percent
                free = round((int(usage.free)/1024 ** 2), 2)
                used = round((int(usage.used)/1024 ** 2), 2)
                l = {'ldisk': ld, 'load': round(perc), 'free': free, 'used': used}
                lw = {'ldisk': ld, 'load': round(perc)}
                ldisks.append(l)
                ldisks_widgets.append(lw)
        x = {'ldisks': ldisks}
        xw = {'ldisks': ldisks_widgets}

        hdd[hdd_count] = x
        hdd_widgets[hdd_count] = xw
    ###NET###
    net = {}
    for i, adapter in enumerate(adapters):
        try:
            nm = adapter.name.decode("utf-8")
            reg = winreg.ConnectRegistry(None, winreg.HKEY_LOCAL_MACHINE)
            reg_key = winreg.OpenKey(
                reg, r"SYSTEM\CurrentControlSet\Control\Network\{4d36e972-e325-11ce-bfc1-08002be10318}")
            try:
                reg_subkey = winreg.OpenKey(reg_key, nm + r"\Connection")
                name = winreg.QueryValueEx(reg_subkey, "Name")[0]
            except FileNotFoundError:
                pass
            else:
                counters = psutil.net_io_counters(pernic=True)
                btx = counters[name][0]
                brx = counters[name][1]
                ptx = counters[name][2]
                prx = counters[name][3]
                if btx > 0 and brx > 0:
                    netpos = 0
                    netpos = netpos+1
                    net_count = 'net_'+str(netpos)
                    netname = name
                    net_io = psutil.net_io_counters(pernic=True)
                    time.sleep(1)
                    net_io_1 = psutil.net_io_counters(pernic=True)
                    rx = round((net_io_1[netname][1] -net_io[netname][1]) / 1024 ** 2., 4)
                    tx = round((net_io_1[netname][0] -net_io[netname][0]) / 1024 ** 2., 4)
                    n = {'btx': btx, 'brx': brx, 'ptx': ptx,'prx': prx, 'tx': tx, 'rx': rx}
                    net[net_count] = (n)
        except KeyError:
            continue
    ###uptime###
    seconds = round(uptime())
    minutes, seconds = divmod(seconds, 60)
    hours, minutes = divmod(minutes, 60)
    days, hours = divmod(hours, 24)
    upt = str(days) + 'd ' + str(hours) + 'h ' + str(minutes) + 'm ' + str(seconds) + 's'    
    
    load_data = {
        "id": {
            "user_id": id_user,
            "device_id": id_computer
        },
        "state": "load",
        "load": {
            "cpu": cpu_load,
            "ram": ram,
            "hdd": hdd,
            "net": net,
            "mb": mb_data
        },
        "widgets": {
            "cpu_load": cpu_total_load_widg,
            "cpu_temp": cpu_total_temp_widg,
            "ram_load": round(memory.percent),
            "hdd_load": hdd_widgets,
            "mb_temp": mb_temp_arr,
            "mb_fans": mb_fan_arr,
            "uptime":upt
        }
    }

    url = 'https://monyze.ru/api.php'
    try:
        requests.post(url, json.dumps(load_data))
        logger.info('Состояние устройства:\n' + json.dumps(load_data, indent=4) + '\n ---------------')
        logger.info('Jobs done')
    except:
        logger.warning('Ошибка отправки данных load_data')
        logger.warning(traceback.format_exc())
        pass

'''
get_config_data(HardwareHandle)

#print('------------------------------------------------------------------------------------------------')

get_load_data(HardwareHandle)
'''
class monyze_agent(win32serviceutil.ServiceFramework):
    _svc_name_ = "monyze_agent"
    _svc_display_name_ = "Monyze agent service"

    def __init__(self, args):
        logger.info('Запуск службы')
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        socket.setdefaulttimeout(60)
        try:
            get_config_data(HardwareHandle)
            logger.info('Служба запущена')
        except:
            logger.warning("Не удалось запустить службу")
            logger.warning(traceback.format_exc())
      
    def SvcStop(self):
        logger.info('Остановка службы')
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        rc = None
        #get_config_data(HardwareHandle)
        logger.info('Отправка данных')
        while rc != win32event.WAIT_OBJECT_0:
            try:
                get_load_data(HardwareHandle)
                rc = win32event.WaitForSingleObject(self.hWaitStop, 5000)
            except:
                logger.warning("Ошибка в SvcDoRun - get_load_data")
                logger.warning(traceback.format_exc())
                pass

if __name__ == '__main__':
    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(monyze_agent)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(monyze_agent)
