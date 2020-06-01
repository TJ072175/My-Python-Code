import os, os.path


Result_Path = os.getcwd() + "\\2_Analysis Result\\" + Car_Number
if not os.path.exists(Result_Path):
    os.makedirs(Result_Path)

