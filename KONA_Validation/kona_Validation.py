
import xlrd
from openpyxl import load_workbook
import sys
import os
from datetime import datetime
import io
import numpy as np
sys.stdout = io.TextIOWrapper(sys.stdout.detach(), encoding = 'utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.detach(), encoding = 'utf-8')

excelPath = str(sys.argv[1])
excelPath = os.path.abspath(excelPath)
#excelPath = 'D:/minwoo/Working_Directory/03_이대엽_크로마틴 구조기반 간암 유방암 예후예측 3D-nucleome 바이오마커 발굴_20190710.xlsx'
'''
input value = excel File
'''

bioSample_SampleName = []
experiment_SampleName = []

def checkingReleaseDate(release_date):
    '''
    날짜를 받아와서 현재 날짜를 기준으로 1년 이내이면 true, 이후이면 false 반환
    '''
    (year,month,day) = release_date.split('-',2)
    currentTime = datetime.now()
    day = str(day).split(' ',1)[0]
    releaseTime = datetime(int(str(year)),int(str(month)),int(str(day)))

    if int(((releaseTime-currentTime)).days) > 365:
        return False
    else:
        return True



def bioProject_Validation(targetSheet):
    '''
    bioProject 시트에서 조건들을 검사한다.
    '''
    flag = 0

    #Submission date check
    if not str(object=targetSheet[str('E')+str(17)].value).count('-') == 2:
        print('[ERROR] [BIOPROJECT] "Submission date" 입력값 형식이 적절하지 않습니다.(17 row, 입력형식:YYYY-MM-DD)')
        flag = 1
    #Release date Check
    if str(targetSheet[str('E')+str(18)].value) == "Release on specified date":
        if not str(targetSheet[str('E')+str(19)].value).count('-') == 2:
            print('[ERROR] [BIOPROJECT] "Release on specified date" 를 선택한 경우 반드시 공개날짜를 입력해야합니다.(19 row,입력형식:YYYY-MM-DD')
            flag = 1
        elif not checkingReleaseDate(str(targetSheet[str('E')+str(19)].value)):
            print("[ERROR] [BIOPROJECT] 'Release Date'가 현재로부터 1년 이후로 설정되어있습니다.(19 row)")
            flag = 1
    elif not str(targetSheet[str('E')+str(18)].value) == "Release immediately following curation (recommended)":
        print("[ERROR] [BIOPROJECT] 'Release date section' 선택 입력값이 적절하지 않습니다.(18 row, 설명에있는 예시중 선택해야함)")
        flag = 1

    #M/O Filed Check
    i = 3
    while i < 52:
        if (str(targetSheet[str('B')+str(i)].value) == 'M') or (str(targetSheet[str('B')+str(i)].value) == 'O'):
            i += 1
        else:
            if i==16 or i==20 or i==35 or i==37 or i==40 or i==42 or i==49:
                i+=1
            else:
                print("[ERROR] [BIOPROJECT] " + str(i) + "번째 row의 M/O 필드 값이 적절하지 않습니다.")
                flag = 1
                i += 1

    #Project type Check
    if str(object=targetSheet[str('E')+str(26)].value)=='총괄':
        print('[ERROR] [BIOPROJECT] "Project type"이 총괄인 경우 따로 결정해서 정리해야합니다.(26 row)')
        flag = 1

    #Government department Check
    if str(object=targetSheet[str('E')+str(21)].value) not in ['과기정통부','해양수산부','보건복지부','농림축산부','산업부','농진청','산림청',]:
        print("[ERROR] [BIOPROJECT] 'Government department' 선택 입력 값이 잘못되었습니다. (21 row)")
        flag = 1

    if(flag==0):
        print("<<< bioProject : NO PROBLEM >>>")






def bioSample_Validation(targetSheet):
    flag = 0


    #Submission date check
    if not str(object=targetSheet[str('E')+str(19)].value).count('-') == 2:
        print('[ERROR] [BIOSAMPLE] Submission date 입력값 형식이 적절하지 않습니다.(17 row, 입력형식:YYYY-MM-DD)')
        flag = 1

    #Release date Check
    if str(targetSheet[str('E')+str(20)].value) == "Release on specified date":
        if not str(targetSheet[str('E')+str(21)].value).count('-') == 2:
            print('[ERROR] [BIOSAMPLE] "Release on specified date" 를 선택한 경우 반드시 공개날짜를 입력해야합니다.(19 row,입력형식:YYYY-MM-DD')
            flag = 1
        elif not checkingReleaseDate(str(targetSheet[str('E')+str(21)].value)):
            print("[ERROR] [BIOSAMPLE] Release Date가 현재로부터 1년 이후로 설정되어있습니다.(19 row)")
            flag = 1
    elif not str(targetSheet[str('E')+str(20)].value) == "Release immediately following curation (recommended)":
        print("[ERROR] [BIOSAMPLE] Release date section 선택 입력값이 적절하지 않습니다.(18 row, 설명에있는 예시중 선택해야함)")
        flag = 1

    #Project accession check
    if targetSheet[str('E')+str(17)].value == None or str(object=targetSheet[str('E')+str(17)].value) == 'NA':
        print('[ERROR] [BIOSAMPLE] Project accession 값을 입력해야합니다.(17 row)')
        flag = 1


    if flag==0:
        print("<<< bioSample : NO PROBLEM >>>")

def sampleType_Validation(targetSheet):

    flag = 0
    i = 5

    #Save Sample names
    while True:
        temp = object=targetSheet[str('A')+str(i)].value
        if temp == None:
            break
        else:
            bioSample_SampleName.append(str(temp))
            i += 1
#def Experiment_Validation(targetSheet):
#    flag = 0

def Experiment_Validation(targetSheet):
    flag= 0

    i = 5
    #Save Sample names
    while True:
        temp = object=targetSheet[str('A')+str(i)].value
        if temp == None:
            break
        else:
            experiment_SampleName.append(str(temp))
            i += 1

    # Compare sample_name of experiment with sample_name of bio_sample
    if not np.array_equal(experiment_SampleName,bioSample_SampleName):
        print("[ERROR] [EXPERIMENT] Sample name 항목이 BioSample에 기입된 Sample name 항목과 동일하지 않습니다.")


    #Release date Check
    if str(targetSheet[str(object='C'+str(5))].value) == "Release on specified date":
        i = 5
        while True:
            temp = targetSheet[str('D')+str(i)].value
            if temp == None:
                break
            else:
                if not (checkingReleaseDate(temp)):
                    print("[ERROR] [EXPERIMENT] Release Date가 현재로부터 1년 이후로 설정되어있습니다.(" + str(i) + " row)")
                i += 1


    #Size value check
    i = 5
    while True:
        temp = targetSheet[str('Q')+str(object=i)].value
        if temp == None:
            break
        if str(temp) == "Paired-end":

            if (str(targetSheet[str('R')+str(i)].value)=='None' or str(targetSheet[str('R')+str(i)].value)=='NA') and (str(targetSheet[str('S')+str(i)].value)=='None' or str(targetSheet[str('S')+str(i)].value)=='NA'):
                print("[ERROR] [EXPERIMENT] Paired-end인 경우 Insert size 또는 Normal size중 하나는 값을 입력해야합니다.("+str(object=i)+" row)")
            i += 1
        else:
            i += 1




try:

    targetExcel = load_workbook(excelPath,data_only=True) # 엑셀 연다.

    bioProject = targetExcel["1) BioProject"]
    bioSample = targetExcel["2) BioSample_Human (1)"]
    sampleType = targetExcel["2-5) Sample type_5_Human (1)"]
    experiment = targetExcel["3) Experiment_Human (1)"]
    bioProject_Validation(bioProject)
    bioSample_Validation(bioSample)
    sampleType_Validation(sampleType)
    Experiment_Validation(experiment)


except IOError as err:
    print("IO Error : " + str(err))
