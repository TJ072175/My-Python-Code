"""
目的：尽量自动化地将ini文件从压缩包中导入到目标文件夹
输出：
1）将ini文件从压缩包解压并存放在目标文件夹
2）显示未能自动导入的ini文件并将记录存放在Unfinished_List.Txt文件中。
现阶段问题：由于未知原因，部分压缩文件无法正常解压，导致ini文件无法提取。
注意事项：
1）需要手动修改路径
2）最好先运行Classify_File_Copy模块
"""
import os,shutil
import rarfile

# 数据源，需要手动修改
Source_Path = r'F:\2017_CPT\2017_CPT_2_CITY\RAW'
# 车辆列表，需要手动修改
Car_List = ['A5S-007', 'A9H-008', 'AC6-016', 'AS7-447', 'AS7-448', 'BV5-933', 'LNF-012', 'MC6-305', 'MT6-029', 'SR8-038', 'SV8-024']
# 目标文件夹，即ini将被导入到此文件夹内
Target_Path = os.getcwd() + "/1_Test Data"
# ini文件的名字
Logger_Config_Identify_File_Name = 'ml_rt.ini'
# 将未能自动导入的ini文件记录在txt文件里。未能导入的原因可能有：1，程序未能解压压缩包；2，路径错误导致文件不存在
Unfinished_List_Txt = open(os.getcwd() + "/" + "Unfinished_List.txt", 'w')
# 建立临时文件夹，用于存放解压出来的ini文件
if not os.path.exists(Target_Path + '/' + 'Temp'):   os.makedirs(Target_Path + '/' + 'Temp')
# 如果目标文件夹内没有车辆文件夹，则新建一个
for Car_Name in Car_List:
    if not os.path.exists(Target_Path + '/' + Car_Name):   os.makedirs(Target_Path + '/' + Car_Name )
# 遍历数据源里各个日期的压缩包
for iDate, Classify_Folder_Date in enumerate(os.listdir(Source_Path)):
    password_file = open(Source_Path + "/" + Classify_Folder_Date + "/password.txt", 'r')   # 存放密码的文件
    for iLine, line in enumerate(password_file.readlines()):
        if iLine == 2:password = str(line)  # 读取密码
    print(Classify_Folder_Date, password)
    rf = rarfile.RarFile(Source_Path + '/' + Classify_Folder_Date + '/' + Classify_Folder_Date + '.rar')    # 打开压缩包
    # 依次提取各个车辆在该天的ini文件
    for Car_Name in Car_List:
        ini_path = Classify_Folder_Date + '/' + Car_Name + '/' + Logger_Config_Identify_File_Name   # 定位ini文件
        # 尝试提取ini文件，并将其解压到临时文件夹
        try:
            rf.extract(ini_path, path=Target_Path + '/' + 'Temp', pwd=password)
        except: # 如果由于压缩包问题提取失败，则将情况记录下来
            print(ini_path + '----bad')
            Unfinished_List_Txt.write(ini_path)
            Unfinished_List_Txt.write('\n')
        else:
            try:    # 提取成功后，将其从临时文件夹移动到目标文件夹
                shutil.move(Target_Path + '/' + 'Temp/' + ini_path,
                        Target_Path + '/' + Car_Name + '/' + Classify_Folder_Date + '/' + Logger_Config_Identify_File_Name)
            except: # 如果出现问题，则说明压缩包里的路径有问题，所以虽然可以解压，但是并没有提取到文件
                print(ini_path + '----no classify file')
            else:
                print(ini_path + '----good')


Unfinished_List_Txt.close()
# 删除临时文件夹
shutil.rmtree(Target_Path + '/' + 'Temp')






'''
password = '123'
rf = rarfile.RarFile('2017-09-13_08-05-48.rar')
rf.setpassword(password)
for f in rf.infolist():
    if f.filename.find(Logger_Config_Identify_File_Name) != -1:
        print(f.filename)
        rf.extract(f.filename, path=Target_Path + '/' + Car_Name + '/' + Classify_Folder_Date, pwd=password)
'''

