import xlrd
from openpyxl import load_workbook
import numpy as np
import sys
import os
import collections

#excelPath = str(sys.argv[1])
txtPath="D:\minwoo\Working_Directory/1711072542_2018M3C9A6017315.txt"
excelPath = "D:\minwoo\Working_Directory/03_박동현_진행성 대장암 예후 예측 바이오마커 발굴 및 임상유용성 검증_20190720_박지환 수정.xlsx"

'''
첫번째인자값 : excel파일 경로, 두번째인자값 : text파일 경로
'''



def getEveryColumnValues(workingSheetName,listName,pathListName,column_Number):
    '''
    사용할시트,저장할배열,파싱할 칼럼값 받아와서 해당 시트의 해당 칼럼에
    해당하는 모든값을 인자로 들어온배열에저장하는 함수 정의
    '''
    item = "0"
    i = 5
    while True:
        item = workingSheetName[str(column_Number)+str(i)].value
        if item==None :
            break
        elif item=='NA' or item=='NULL' or item=='null':
            i += 1
        else:
        #    print('temp:')
            listName.append(item)
            (temp,trash) = str(workingSheetName[str('AA')+str(i)].value).rsplit('/',1)
            temp = temp + '/' + str(object=item)
            pathListName[str(item)] = str(temp)
            #print(temp)
            i += 1


xmlFileList = []
xmlFilePathDic = collections.OrderedDict()
txtFileList = []
txtFilePathDic = collections.OrderedDict()
commonFileList = []

try:
    targetExcel = load_workbook(excelPath,data_only=True) # 엑셀 연다.

    workingSheet1 = targetExcel["3) Experiment_Human (1)"]
    workingSheet2 = targetExcel["3) Experiment_Human (2)"]
    workingSheet3 = targetExcel["3) Experiment_Human (3)"]

    '''
    파일목록 4개 가져와서 모두 하나의 배열에 저장
    '''
    getEveryColumnValues(workingSheet1,xmlFileList,xmlFilePathDic,'V')
    getEveryColumnValues(workingSheet1,xmlFileList,xmlFilePathDic,'X')
    getEveryColumnValues(workingSheet2,xmlFileList,xmlFilePathDic,'V')
    getEveryColumnValues(workingSheet2,xmlFileList,xmlFilePathDic,'X')
    getEveryColumnValues(workingSheet3,xmlFileList,xmlFilePathDic,'V')
    getEveryColumnValues(workingSheet3,xmlFileList,xmlFilePathDic,'X')
except IOError as err:
    print("IO Error : " + str(err))



try:
    targetTxt = open(txtPath,'rt',encoding='UTF8')
    lines = targetTxt.readlines()
    '''
    txt파일 가져와서 파일목록 배열에 저장
    '''
    for line in lines:
        (trash, value) = line.rsplit('/',1)
        temp = value.rstrip('\n')
        txtFileList.append(temp)
        txtFilePathDic[str(temp)] = '/home/qu/KOBIC/2018'+ str(line).split('.',1)[1]
        #txt 경로 셋 저장
except IOError as err:
    print("Txt File Error: " + str(err))



#for key, value in txtFilePathDic.items():
    #print(key,":",value)


onlyXml = []
onlyXmlPath = []
onlyTxt = []
onlyTxtPath = []
common = []
commonPath = []


for x in xmlFileList: #xml파일 순회
    if x not in txtFileList: #txt에 없으면
        onlyXml.append(x)
        onlyXmlPath.append(xmlFilePathDic[x].rstrip('\n'))
            #xml만있고, txt파일에 없는 목록들 저장



for x in txtFileList:
    if x not in xmlFileList:
        onlyTxt.append(x)
        onlyTxtPath.append(txtFilePathDic[x].rstrip('\n'))
        #txt파일에만있고, xml에 없는 목록들 저장
    else:
        common.append(x)
        commonPath.append(txtFilePathDic[x].rstrip('\n'))


print(str(object=len(xmlFileList)))
print(str(object=len(txtFileList)))



outputfile = open('outputfile.txt', mode='w', encoding='utf-8')
outputfile.write('only_excel' + '\t' + 'only_excel_filePath' + '\t' + 'common' + '\t'+'common_filepath'+'\t' + 'only_server'+'\t' + 'only_server_filePath'+'\n')

lengths = [len(onlyXml),len(onlyTxt),len(common)]
rotation = max(lengths)
print(rotation)
r = 0

while r < rotation:
    if r < len(onlyXml):
        outputfile.write(str(object=onlyXml[r])+'\t')
        outputfile.write(str(object=onlyXmlPath[r])+'\t')
    else:
        outputfile.write('\t')
        outputfile.write('\t')

    if r < len(common):
        outputfile.write(str(object=common[r])+'\t')
        outputfile.write(str(object=commonPath[r])+'\t')
    else:
        outputfile.write('\t')
        outputfile.write('\t')

    if r < len(onlyTxt):
        outputfile.write(str(object=onlyTxt[r])+'\t')
        outputfile.write(str(object=onlyTxtPath[r])+'\n')
    else:
        outputfile.write('\t')
        outputfile.write('\n')

    r += 1

outputfile.close()
