# -*- coding:utf-8 -*-
import os,scipy, shutil
import numpy as np
#from mdfreader import *
import time
import matplotlib.pyplot as plt

import win32com.client
import subprocess
from subprocess import Popen, PIPE

import mdfreader, openpyxl
import pptx
from pptx.util import Inches,Pt

print('program start')
test_mode = 1   # 0=> 实际使用；1=〉测试，跳过转换格式
txt = open(os.getcwd() + '\Analysis Result.txt', 'w', encoding='utf-8')
data_load_total_duration = data_analysis_total_duration = 0
file_path_root = os.getcwd() + '\\1_Data\\'
convert_folder_name = 'Event group 1'
# 选择使用那一套信号名称定义
data_type = 1   #1=>ECU signal; 2=>MQB CAN; 3=>PQ CAN; 4=>MQB PHEV
time_name = 'Zeitkanal'
if data_type == 1:
    # ECU信号
    speed_name = 'vfzg_w'
    engine_speed_name = 'nmot_w'
    #engine_speed_name = 'MO_Drehzahl_01'
    engine_speed_sol_name = 'nsol_w'
    pedal_name = 'wped_w'
    kl15_name = 'B_kl15'
    outside_temperature_name = 'tumg'
    water_temperature_name = 'tmot'
    ac_name = 'B_koe'
    gang_name = 'Tra_numGear'
    gang_prnd_name = 'GE_Ist_Fahrstufe'
    start_stop_name = 'MO_StartStopp_Fahrerwunsch'
    relative_load_name = 'rl_w'
    throttle_name = 'wdkba_w'
    mileage_name = 'Dspl_lMlg'
    stupid_error_name = 'B_atmtpa'
    month_name = 'Month'
    day_name = 'Day'
elif data_type == 2:
    # MQB的CAN信号
    speed_name = 'ESP_v_Signal'
    engine_speed_name = 'MO_Drehzahl_01'
    engine_speed_sol_name = 'MO_Drehzahl_01'
    pedal_name = 'MO_Fahrpedalrohwert_01'
    kl15_name = 'MO_Kl_50'
    outside_temperature_name = 'KBI_Aussen_Temp_gef'
    water_temperature_name = 'MO_Kuehlmittel_Temp'
    ac_name = 'KL_AC_Schalter'
    gang_name = 'GE_Zielgang'
    gang_prnd_name = 'GE_Ist_Fahrstufe'
    start_stop_name = 'MO_StartStopp_Fahrerwunsch'
    relative_load_name = 'rl_w'
elif data_type == 3:
    # PQ的CAN信号
    speed_name = 'KO1_kmh'
    engine_speed_name = 'MO1_Drehzahl'
    engine_speed_sol_name = 'MO1_Drehzahl'
    pedal_name = 'MO3_Pedalwert'
    kl15_name = r'BSL_ZAS_Kl_15'
    outside_temperature_name = 'MO3_Aussentemp'
    water_temperature_name = 'MO2_Kuehlm_T'
    ac_name = 'MO2_Sta_Klima'
    gang_name = 'GE1_Zielgang'
    gang_prnd_name = 'GE2_PRNDS'
    start_stop_name = 'GK2_Kl_StSt_Info'
    relative_load_name = 'rl_w'
elif data_type == 4:
    # MQB的CAN信号，但是针对PHEV，转速信号的名称有所不同
    speed_name = 'ESP_v_Signal'
    engine_speed_name = 'MO_Drehzahl_VM'
    engine_speed_sol_name = 'MO_Drehzahl_01'
    pedal_name = 'MO_Fahrpedalrohwert_01'
    kl15_name = 'MO_Kl_50'
    outside_temperature_name = 'KBI_Aussen_Temp_gef'
    water_temperature_name = 'MO_Kuehlmittel_Temp'
    ac_name = 'KL_AC_Schalter'
    gang_name = 'GE_Zielgang'
    gang_prnd_name = 'GE_Ist_Fahrstufe'
    start_stop_name = 'MO_StartStopp_Fahrerwunsch'
    relative_load_name = 'rl_w'

template_folder = os.getcwd() + '/3_Template/'
ppt_path = template_folder + 'EOBD Daily Report Template.pptx'
ppt = pptx.Presentation(ppt_path)
vehicle_number = []
project_mileage = []
data_record_date = []
drive_time = []
speed_max = []
speed_above_80 = []
speed_above_120 = []
speed_average = []
engine_speed_above_3800 = []
engine_speed_max = []
pedal_above_70 = []
pedal_max = []

# 读取信号列表等信息
wbk = openpyxl.load_workbook('SVW_OBD_Signal_List.xlsx')
sht = wbk['Signal_List']
x = 2
car_list = []
dr_list = []
mileage_list = []
speed_list = []
engine_speed_list = []
pedal_list = []
month_list = []
day_list = []
while sht.cell(row=x, column=1).value != None:
    car_list.append(sht.cell(row=x, column=1).value)
    dr_list.append(sht.cell(row=x, column=2).value)
    mileage_list.append(sht.cell(row=x, column=3).value)
    speed_list.append(sht.cell(row=x, column=4).value)
    engine_speed_list.append(sht.cell(row=x, column=5).value)
    pedal_list.append(sht.cell(row=x, column=6).value)
    month_list.append(sht.cell(row=x, column=7).value)
    day_list.append(sht.cell(row=x, column=8).value)
    x += 1
# 在根目录下遍历各个日期，日期文件夹内含有MDF文件
for i_folder, folder_name in enumerate(os.listdir(file_path_root)):
    # 如果不是文件夹，即不是车号文件夹，则跳过
    if not os.path.isdir(file_path_root + folder_name):
        continue
    # 判断记录仪的类型、信号名称等（iav or gin）
    i_car = car_list.index(folder_name)
    logger = dr_list[i_car]
    if logger == 'IAV':
        # 识别日期并转换格式
        # ------------------------------------------------------------------------------------------------------------------
        iav_convert_path = r'C:\Program Files (x86)\IAV GmbH\Drive Recorder NG\2.1\DrNGBatchConv.exe'

        ini_folder = file_path_root + folder_name   # 需要确保ini文件的路径中没有空格，否则无法执行转换格式的程序
        # 识别带有数据的文件夹
        for index, item in enumerate(os.listdir(ini_folder)):
            if os.path.isdir(ini_folder + '\\' + item) and item != 'SYS':
                data_folder = file_path_root + folder_name + '\\' + item
                break

        folder_day = []
        file_month = []
        file_size = []
        each_day_info_size = []
        each_day_info_day = []
        each_day_info_month = []
        for i_file, file_name in enumerate(os.listdir(data_folder)):
            statinfo = os.stat(data_folder + '/' + file_name)
            if file_name[0] == 'A' and file_name[-5:] == 'NA.EV':  # 判断是否为A开头的文件，即需要分析的文件
                folder_day.append(time.localtime(statinfo.st_mtime).tm_mday)
                file_month.append(time.localtime(statinfo.st_mtime).tm_mon)
                file_size.append(statinfo.st_size / 1000 / 1000)

        date_flag = 0
        date_count = 0
        for i, date in enumerate(folder_day):
            # 一天里的第一个文件
            if date_flag == 0:
                i_date_start = i
                date_flag = 1
                each_day_info_size.append(file_size[i])
                each_day_info_day.append(folder_day[i])
                each_day_info_month.append(file_month[i])
            # 一天里的最后一个文件
            if i == len(folder_day) - 1 or date != folder_day[i + 1]:
                i_date_end = i
                date_flag = 0
                for i_temp in range(i_date_start + 1, i_date_end + 1):
                    each_day_info_size[date_count] += file_size[i_temp]
                date_count += 1
        #print(file_size)
        #print(each_day_info_size)
        # 转化格式
        print(folder_name + '转换格式中')
        execute_string = '"' + iav_convert_path + '"' + ' /SRC=""' + ini_folder + "\iavdrvng.ini" + '"" -EDir=' + '"' + ini_folder + '"' + " -EFormat=2"
        if test_mode == 0:
            subprocess.call(execute_string)
        # ------------------------------------------------------------------------------------------------------------------
    elif logger == 'GIN':
        # 转化格式
        print(folder_name + '转换格式中')

        execute_string = file_path_root + folder_name + '/convert.cmd'
        if test_mode == 0:
            shutil.copyfile(template_folder + folder_name + '/convert.cmd', file_path_root + folder_name + '/convert.cmd')
            p = subprocess.Popen(execute_string, cwd=file_path_root + folder_name)
            p.wait()
