# -*- coding: utf-8 -*-
import servicemanager
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

c = wmi.WMI()


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


ohwm_dll = "C:\\monyze_windows\\OpenHardwareMonitorLib.dll"

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

def get_hdd_info():
    logger.info ('Получаем данные о дисках')
    hddpos = 0
    for pdisk in c.query("SELECT * FROM Win32_DiskDrive"): 
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
            driveinfo.append(l)
            
    logger.info ('Данные о дисках получены')
    logger.info(driveinfo)


class TestService(win32serviceutil.ServiceFramework):
    _svc_name_ = "HddOnly"
    _svc_display_name_ = "HddOnly"

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        socket.setdefaulttimeout(60)

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        rc = None
        while rc != win32event.WAIT_OBJECT_0:
            get_hdd_info()
            rc = win32event.WaitForSingleObject(self.hWaitStop, 5000)




if __name__ == '__main__':
    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(TestService)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(TestService)

