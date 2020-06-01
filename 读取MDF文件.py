# -*- coding:utf-8 -*-
import os,scipy, shutil
import numpy as np


import mdfreader


# 将数据加载到内存中
DataFile = mdfreader.mdf()
DataFile.read(fileName=File_Path, channelList=signal_list)
# 由于加载信号时可能会没法加载时间通道以至于无法进行resample，所以需要额外再加载一次以补上时间通道
timeline_list = []
for i in range(len(signal_list)):
    timeline_list.append(DataFile.getChannelMaster(signal_list[i]))
DataFile.read(fileName=file_path, channelList=timeline_list)
'''
# 针对过去遇到的一些读取信号数值时遇到的乱码问题，现在读取时没有这个问题
for n in range(len(special_signal_list)):
    key_list=[]
    value_list_of_key=[]
    for key,value in DataFile.masterChannelList.items():
        key_list.append(key)
        value_list_of_key.append(value)
    for i in range(len(value_list_of_key)):
        if special_signal_list[n] in value_list_of_key[i]:
            #get_value_index = value_list_of_key[i].index(Special_Signal_List[n])
            vars()[special_signal_list[n] + '_value'] = np.where(DataFile.getChannelData(special_signal_list[n]) == 'FALSE', 0, 1)
            #print(vars()[Special_Signal_List[n] + '_value'])
            DataFile.remove_channel(special_signal_list[n])
            #print(DataFile.masterChannelList)
            DataFile.add_channel(dataGroup=key_list[i], channel_name=special_signal_list[n], data=vars()[special_signal_list[n] + '_value'], master_channel=key_list[i], )
            #print(DataFile.masterChannelList)
print(DataFile.getChannelData(kl15_name).max())
'''
# resample,如果不填写数字或者信号则默认以采样频率最高者为基准
DataFile.resample(0.01)
# 获取信号数值，为numpy格式
engine_speed_name = 'nmot_w'
engine_speed_value = DataFile.getChannelData(engine_speed_name)
# 获取信号所在组的组名
engine_speed_group = DataFile.getChannelMaster(engine_speed_name)
# 移除、添加、修改信号
DataFile.remove_channel(engine_speed_name)
DataFile.add_channel(dataGroup=engine_speed_group, channel_name=engine_speed_name, data=engine_speed_value, master_channel=engine_speed_group)
DataFile.setChannelData(engine_speed_name, engine_speed_value, compression=False)



# 修复里程信号（有时数值会变为0）
if mileage_name in DataFile.keys():
    mileage_channel_master = DataFile.getChannelMaster(mileage_name)
    mileage_value_temp = DataFile.getChannelData(mileage_name)
    for i in range(len(mileage_value_temp) - 1, -2, -1):
        if mileage_value_temp[i] == 0:
            mileage_value_temp[i] = mileage_value_temp[i + 1]
    DataFile.remove_channel(mileage_name)
    DataFile.add_channel(dataGroup=mileage_channel_master, channel_name=mileage_name, data=mileage_value_temp, master_channel=mileage_channel_master)

print(DataFile.masterChannelList)   # resample前的信号列表
#print(DataFile.MDFVersionNumber)    # MDF版本
# 对信号进行resample，默认以采样频率最高的信号为基准
print('file loaded, start resampling')
#DataFile.resample(masterChannel=DataFile.getChannelMaster(engine_speed_name))
DataFile.resample(0.01)
#DataFile.resample()
print(DataFile.masterChannelList)   # resample后的信号列表
# resample后的时间通道
#t_new = 'master'

# 把resample后的数据赋值给变量(并进行拼接)
t_new = DataFile.getChannelMaster(engine_speed_name)
time_value_temp = DataFile.getChannelData(t_new)
#print(time_value_temp)
if i_file > 0:
    time_value_temp += time_value.max()
time_value = np.append(time_value, time_value_temp )
for i in range(len(signal_list)):
    if signal_list[i] in DataFile.keys():
        vars()[signal_name_list[i] + '_value'] = np.append(vars()[signal_name_list[i] + '_value'], DataFile.getChannelData(signal_list[i]))
    else:
        print(signal_list[i] + ' not found')
        #nofound_signal_list.append(signal_list[i])
        #vars()[signal_name_list[i] + '_value'] = np.append(vars()[signal_name_list[i] + '_value'], np.zeros(len(time_value), dtype=np.int))


# ----------------------------------------------------------------------------------
# 将一个物理量可能的信号名称都存放在Excel中，然后程序依次判断应该根据哪个信号名来读取数据
# ----------------------------------------------------------------------------------
# -------------------------------------------------------
# 设置要加载的信号列表
# -------------------------------------------------------
self.wbk_signal = openpyxl.load_workbook('Signal List.xlsx')
self.sht_signal = self.wbk_signal.worksheets[0]
# 表单1，用于大图
self.signal_name_list_1 = [
    'speed',
    'engine speed',
    'pedal',
    'water temp',
    'Lamb Precat',
    'Lamb Postcat'
]
self.speed_list = self.get_signal_list(2)
self.engine_speed_list = self.get_signal_list(3)
self.pedal_list = self.get_signal_list(4)
self.water_temp_list = self.get_signal_list(5)
self.lamb_precat_list = self.get_signal_list(6)
self.lamb_postcat_list = self.get_signal_list(7)


def get_signal_list(self, x):
    y = 2
    signal = []
    while self.sht_signal.cell(row=x, column=y).value != None:
        signal.append(self.sht_signal.cell(row=x, column=y).value)
        y += 1
    return signal


def get_signal_name(self, list, names):
    for name in names:
        if name in list:
            return name
        
self.data_file = mdfreader.mdf()
# 先读取文件内所有信号
self.data_file.read(fileName=self.file_path)
channel_list = self.data_file.keys()
# 根据文件内信号和之前读取的信号列表来确定各个物理量的信号名称
# 1
self.speed_name = self.get_signal_name(channel_list, self.speed_list)
self.engine_speed_name = self.get_signal_name(channel_list, self.engine_speed_list)
self.pedal_name = self.get_signal_name(channel_list, self.pedal_list)
self.water_temp_name = self.get_signal_name(channel_list, self.water_temp_list)
self.lamb_precat_name = self.get_signal_name(channel_list, self.lamb_precat_list)
self.lamb_postcat_name = self.get_signal_name(channel_list, self.lamb_postcat_list)
self.signal_list_1 = [
    self.speed_name,
    self.engine_speed_name,
    self.pedal_name,
    self.water_temp_name,
    self.lamb_precat_name,
    self.lamb_postcat_name,
]
self.data_file = mdfreader.mdf()
self.data_file.read(fileName=self.file_path, channelList=self.signal_list)
# 由于一个插件bug导致加载信号时可能会没法加载时间通道以至于无法进行resample，所以需要额外再加载一次以补上时间通道
timeline_list = []  # 需要的信号列表
for i in range(len(self.signal_list)):
    timeline_list.append(self.data_file.getChannelMaster(self.signal_list[
                                                       i]))  # Extract channel master name from mdf structure; return name string 如转速信号和对应的时间组成channel master
self.data_file.read(fileName=self.file_path,
              channelList=timeline_list)  # masterChannelList, a dict containing a list of channel names per datagroup 重新加载一次以补上时间通道，之前只通过信号名称加载可能没有时间通道
print('file loaded, start resampling')