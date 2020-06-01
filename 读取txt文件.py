

ClassifyFileTemp = open(ClassifyFilePath + "/" + ClassifyFileNameTemp)
for iLineTemp, lineTemp in enumerate(ClassifyFileTemp.readlines()):
    if iLineTemp == 7:  # 数据在txt的第8行
        lineTemp = lineTemp.split()
        # 如果有更新配置文件，则需要将当天数值存放在临时变量中
        if Logger_Config_Changed == 1:
            AnalysisResult_Condition_Acc_Offset[iCase] = AnalysisResult_Condition_Acc[iCase][iDate - 1]
        if CaseType[iFile + 1] == 2:
            lineTemp[0] = lineTemp[0].replace(",", ".") # 针对类型2，即工况为计时，将原txt中的逗号转换为点，以便将文字转换为数字
            AnalysisResult_Condition_Acc[iCase][iDate] = float(lineTemp[0]) // 60 + AnalysisResult_Condition_Acc_Offset[iCase]
        else:
            AnalysisResult_Condition_Acc[iCase][iDate] = float(lineTemp[0]) + AnalysisResult_Condition_Acc_Offset[iCase]
        # if iCase == 26:
            # print(AnalysisResult_Condition_Acc_Offset[iCase])

