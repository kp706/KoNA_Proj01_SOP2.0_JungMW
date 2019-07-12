import xlrd
from openpyxl import load_workbook
import numpy as np
import sys
import os

excelPath = str(sys.argv[1])
txtPath = str(sys.argv[2])
excelPath = os.path.abspath(excelPath)
txtPath = os.path.abspath(txtPath)
'''
첫번째인자값 : excel파일 경로, 두번째인자값 : text파일 경로
'''


def getEveryColumnValues(workingSheetName,listName,column_Number):
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
            listName.append(item)
            i += 1


xmlFileList = []
txtFileList = []

try:
    targetExcel = load_workbook(excelPath,data_only=True) # 엑셀 연다.

    workingSheet1 = targetExcel["3) Experiment_Human (1)"]
    workingSheet2 = targetExcel["3) Experiment_Human(2)"]


    '''
    파일목록 4개 가져와서 모두 하나의 배열에 저장
    '''
    getEveryColumnValues(workingSheet1,xmlFileList,'V')
    getEveryColumnValues(workingSheet1,xmlFileList,'X')
    getEveryColumnValues(workingSheet2,xmlFileList,'V')
    getEveryColumnValues(workingSheet2,xmlFileList,'X')

    print(xmlFileList)
    print(len(xmlFileList))



except IOError as err:
    print("IO Error : " + str(err))
