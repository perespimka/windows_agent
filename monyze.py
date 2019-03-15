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
import sys
import ifaddr
import clr
import ctypes
from requests.adapters import HTTPAdapter
from requests.packages.urllib3.util.retry import Retry
c = wmi.WMI()

print('sdfsd')

print('Hi')

for diskDrive in c.query("SELECT * FROM Win32_DiskDrive"):
    diskDrive.Size, "\nDisk model: ", diskDrive.Model,  diskDrive.Status

try:
    import _winreg as winreg
except ImportError as err:
    try:
        import winreg
    except ImportError as err:
        pass

#if os.path.isfile('c:\\monyze\\agent\\windows\\keys.key'):
keys= 'c:\\monyze\\agent\\windows\\keys.key'
f = open(keys, 'r')
id_user = f.readline()
id_computer = f.readline()
f.close()
#else:
#    print('file_err')


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

driveinfo = []
hddpos = 0
cpu_load = {}
cpu_cores = []

script_dir = os.getcwd()

openhardwaremonitor_hwtypes = ['Mainboard', 'SuperIO', 'CPU','RAM', 'GpuNvidia', 'GpuAti', 'TBalancer', 'Heatmaster', 'HDD']
cputhermometer_hwtypes = ['Mainboard', 'SuperIO', 'CPU','GpuNvidia', 'GpuAti', 'TBalancer', 'Heatmaster', 'HDD']
openhardwaremonitor_sensortypes = ['Voltage', 'Clock', 'Temperature', 'Load','Fan', 'Flow', 'Control', 'Level', 'Factor', 'Power', 'Data', 'SmallData']
cputhermometer_sensortypes = ['Voltage', 'Clock', 'Temperature', 'Load', 'Fan', 'Flow', 'Control', 'Level']
cpu_load_sensornames = ['CPU Total', 'CPU Core', 'CPU Core #1', 'CPU Core #2','CPU Core #3', 'CPU Core #4', 'CPU Core #5', 'CPU Core #6', 'CPU Core #7', 'CPU Core #8']
cpu_temperature_sensornames = ['CPU Package', 'CPU Core #1', 'CPU Core #2', 'CPU Core #3','CPU Core #4', 'CPU Core #5', 'CPU Core #6', 'CPU Core #7', 'CPU Core #8']


def initialize_openhardwaremonitor():
    ohwm_dll = "c:\\monyze\\agent\\windows\\OpenHardwareMonitorLib.dll"
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

for pdisk in c.query("SELECT * FROM Win32_DiskDrive"):
    hddpos = hddpos + 1
    lnames = []
    Size = str(round((int(pdisk.Size)/1024/1024/1024), 2))
    strcl = pdisk.deviceID.replace("\\", "")
    pdiskcleared = strcl.replace(".PHYSICALDRIVE", "PhysicalDrive")
    for colpartition in c.query('ASSOCIATORS OF {Win32_DiskDrive.DeviceID="' + pdisk.deviceID + '"} WHERE AssocClass = Win32_DiskDriveToDiskPartition'):
        for logical_disk in c.query('ASSOCIATORS OF {Win32_DiskPartition.DeviceID="' + colpartition.DeviceID + '"} WHERE AssocClass = Win32_LogicalDiskToPartition'):
            lname = logical_disk.DeviceID
            lnames.append(lname)
            l = {'hdd_'+str(hddpos)+'': {'name': pdisk.Model,'size': Size, 'LOGICAL': lnames}}
    driveinfo.append(l)


def get_config_data(handle):
    cpuinfo = {}
    c_count = 0
    for i in handle.Hardware:
        i.Update()
        for sensor in i.Sensors:
            if sensor.Value is not None:
                sensortypes = openhardwaremonitor_sensortypes
                hardwaretypes = openhardwaremonitor_hwtypes
                ###CPU'S###
                if sensor.Hardware.HardwareType == hardwaretypes.index('CPU'):
                    if sensor.Index == 0 and sensor.SensorType == 3:
                        c_count = c_count+1
                        cpuinfo['cpu_'+str(c_count)] = sensor.Hardware.Name
                        break
                    if sensor.Index >= 0:
                        c_count = c_count+1
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

    url = 'http://monyze.ru/api.1.php'

    with requests.Session() as s:
        retries = Retry(
            total=10,
            backoff_factor=0.2,
            status_forcelist=[500, 502, 503, 504])

    s.mount('http://', HTTPAdapter(max_retries=retries))
    s.mount('https://', HTTPAdapter(max_retries=retries))

    r = s.post(url, json.dumps(config_data))


def get_load_data(handle):
    memory = psutil.virtual_memory()
    cpu_count = 0

    for i in handle.Hardware:
        sensortypes = openhardwaremonitor_sensortypes
        hardwaretypes = openhardwaremonitor_hwtypes
        load_sensornames = cpu_load_sensornames
        temp_sensornames = cpu_temperature_sensornames
        i.Update()

        ###CPU'S###
        if i.HardwareType == hardwaretypes.index('CPU'):
            if i.processorIndex == cpu_count:
                load_arr = {}
                temp_arr = {}
                cpu_load_widg_arr = []
                cpu_temp_widg_arr = []
                cpu_count = cpu_count + 1
                cpu_load_data = {}
                for sensor in i.Sensors:
                    if sensor.Value is not None:
                        #Сенсоры загрузки CPU
                        if sensor.SensorType == sensortypes.index('Load'):
                            if sensor.Name == load_sensornames[0] or sensor.Name == load_sensornames[1]:
                                load_arr['total'] = round(sensor.Value)
                                cpu_load_widg_arr.append(round(sensor.Value))
                            if sensor.Name == load_sensornames[2]:
                                load_arr['core_1'] = round(sensor.Value)
                            if sensor.Name == load_sensornames[3]:
                                load_arr['core_2'] = round(sensor.Value)
                            if sensor.Name == load_sensornames[4]:
                                load_arr['core_3'] = round(sensor.Value)
                            if sensor.Name == load_sensornames[5]:
                                load_arr['core_4'] = round(sensor.Value)
                            if sensor.Name == load_sensornames[6]:
                                load_arr['core_5'] = round(sensor.Value)
                            if sensor.Name == load_sensornames[7]:
                                load_arr['core_6'] = round(sensor.Value)
                            if sensor.Name == load_sensornames[8]:
                                load_arr['core_7'] = round(sensor.Value)
                            if sensor.Name == load_sensornames[9]:
                                load_arr['core_8'] = round(sensor.Value)
                        #Сенсоры температуры CPU
                        if sensor.SensorType == sensortypes.index('Temperature'):
                            if sensor.Name == temp_sensornames[0]:
                                temp_arr['total'] = round(sensor.Value)
                                cpu_temp_widg_arr.append(round(sensor.Value))
                            if sensor.Name == temp_sensornames[1]:
                                temp_arr['core_1'] = round(sensor.Value)
                                cpu_temp_widg_arr.append(round(sensor.Value))
                            if sensor.Name == temp_sensornames[2]:
                                temp_arr['core_2'] = round(sensor.Value)
                                cpu_temp_widg_arr.append(round(sensor.Value))
                            if sensor.Name == temp_sensornames[3]:
                                temp_arr['core_3'] = round(sensor.Value)
                                cpu_temp_widg_arr.append(round(sensor.Value))
                            if sensor.Name == temp_sensornames[4]:
                                temp_arr['core_4'] = round(sensor.Value)
                                cpu_temp_widg_arr.append(round(sensor.Value))
                            if sensor.Name == temp_sensornames[5]:
                                temp_arr['core_5'] = round(sensor.Value)
                                cpu_temp_widg_arr.append(round(sensor.Value))
                            if sensor.Name == temp_sensornames[6]:
                                temp_arr['core_6'] = round(sensor.Value)
                                cpu_temp_widg_arr.append(round(sensor.Value))
                            if sensor.Name == temp_sensornames[7]:
                                temp_arr['core_7'] = round(sensor.Value)
                                cpu_temp_widg_arr.append(round(sensor.Value))
                            if sensor.Name == temp_sensornames[8]:
                                temp_arr['core_8'] = round(sensor.Value)
                                cpu_temp_widg_arr.append(round(sensor.Value))
                        if sensor.SensorType == sensortypes.index('Clock'):
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

    ram = {"load": round(memory.percent), "AvailPh": memory.available}

    ##HDD###
    hdd = {}
    hdd_widgets = {}
    hddpos = 0
    for disk in driveinfo:
        hddpos = hddpos + 1
        ldisks = []
        ldisks_widgets = []
        hdd_count = 'hdd_'+str(hddpos)+''
        for ld in disk[hdd_count]['LOGICAL']:
            usage = psutil.disk_usage(ld)
            perc = usage.percent
            free = round((int(usage.free)/1024/1024/1024), 2)
            used = round((int(usage.used)/1024/1024/1024), 2)
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
                    rx = round((net_io_1[netname][1] -net_io[netname][1]) / 1024 / 1024., 4)
                    tx = round((net_io_1[netname][0] -net_io[netname][0]) / 1024 / 1024., 4)
                    n = {'btx': btx, 'brx': brx, 'ptx': ptx,'prx': prx, 'tx': tx, 'rx': rx}
                    net[net_count] = (n)
        except KeyError:
            continue
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
            "net": net
        },
        "widgets": {
            "cpu_load": cpu_total_load_widg,
            "cpu_temp": cpu_total_temp_widg,
            "ram_load": round(memory.percent),
            "hdd_load": hdd_widgets
        }
    }

    url = 'http://monyze.ru/api.1.php'

    with requests.Session() as s:
        retries = Retry(
            total=10,
            backoff_factor=0.2,
            status_forcelist=[500, 502, 503, 504])

    s.mount('http://', HTTPAdapter(max_retries=retries))
    s.mount('https://', HTTPAdapter(max_retries=retries))

    r = s.post(url, json.dumps(load_data))


class monyze_agent(win32serviceutil.ServiceFramework):
    _svc_name_ = "monyze_agent"
    _svc_display_name_ = "Monyze agent service"

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        socket.setdefaulttimeout(60)
        log = open('C:\\TestService.log', 'a')
        log.write('Service started...\n')
        log.write('Trying to send config...\n')
        log.close()
        get_config_data(HardwareHandle)
           
        
    def SvcStop(self):
        log = open('C:\\TestService.log', 'a')
        log.write('Service stopped...\n')
        log.close()
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        rc = None
        while rc != win32event.WAIT_OBJECT_0:
            with open('C:\\TestService.log', 'a') as log:
                get_load_data(HardwareHandle)
                log.write('load data sent...\n')
            rc = win32event.WaitForSingleObject(self.hWaitStop, 5000)


if __name__ == '__main__':
    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(monyze_agent)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(monyze_agent)
