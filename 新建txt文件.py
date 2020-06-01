# ------------------------------------------------------------------------------------------------------------------
# 将计算后各课题的总次数和工况总次数写入TXT，在填写汇总分析时会用到
# ------------------------------------------------------------------------------------------------------------------
ResultTxt = open(Result_Path + "\\" + Car_Number + ".txt",'w')
for iCase in range(1, iCaseMax + 1):
    ResultTxt.write(str(max(AnalysisResult_Acc[iCase])))
    ResultTxt.write(' ')
    ResultTxt.write(str(max(AnalysisResult_Condition_Acc[iCase])))
    ResultTxt.write('\n')
ResultTxt.close()

