import os, shutil
# 复制txt文件（整个文件夹复制）
shutil.copytree(Source_Path + '\\' + Classify_Folder_Date + '\\' + Car_Name + '\\Classify',
                Target_Path + '\\' + Car_Name + '\\' + Classify_Folder_Date + '\\Classify')

# 重命名
os.rename(txt_path, ascii_path)