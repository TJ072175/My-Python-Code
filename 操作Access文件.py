import os, pyodbc
# --------------------------------------------------------------------------------------------------------------
# 连接数据库
# --------------------------------------------------------------------------------------------------------------
DBfile = os.getcwd() + '/' + 'MDF_Data' + '.accdb'
'''
DBfile = os.getcwd() + '/' + file_name + '.accdb'

if os.path.exists(DBfile) == False:
    pyodbc.
'''
conn_str = (r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'r'DBQ=' + DBfile + ';')
cnxn = pyodbc.connect(conn_str)
crsr = cnxn.cursor()
# --------------------------------------------------------------------------------------------------------------
# 删除旧的记录并为新的记录新建表格和字段
# --------------------------------------------------------------------------------------------------------------
print('Delete Old Record')
self.status_label.SetLabel('Delete Old Record')
# 删除旧的记录
crsr.execute('drop table Data')
# 为新的记录新建表格和字段
sql = ''
for signal_name in signal_list:
    sql = sql + signal_name + ' numeric, '
sql = 'Create TABLE Data (' + sql[:-2] + ')'
crsr.execute(sql)
cnxn.commit()
# --------------------------------------------------------------------------------------------------------------
# 插入新的记录
# --------------------------------------------------------------------------------------------------------------
print('Insert New Record')
self.status_label.SetLabel('Insert New Record')
x_max = max(signal_length)
self.progress_bar.Show(True)
p = 0
for x in range(x_max):
    # 进度条
    if int(x / x_max * 100) > p:
        print(str(p) + '%')
        p = int(x / x_max * 100)
        self.progress_label.SetLabel(str(p) + '%')
        self.progress_bar.SetValue(p)
    # 插入记录
    sql = 'Insert Into Data ('
    for n in range(len(signal_list)):
        sql = sql + '[' + signal_list[n] + '], '    # 字段名
    sql = sql[:-2] + ') values ('
    for n in range(len(signal_list)):
        if x < len(vars()['signal_value_' + str(n)]):
            sql = sql + str(vars()['signal_value_' + str(n)][x]) + ', ' # 数值
        else:
            sql = sql + '0' + ', '  # 由于采样频率不同，低频的信号长度比高频的要短，所以后续的数值用0来填补
    sql = sql[:-2] + ')'
    # 执行SQL语句
    crsr.execute(sql)


# 保存结果
cnxn.commit()
cnxn.close()