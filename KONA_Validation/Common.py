#import xlrd
#from openpyxl import load_workbook
import sys
import os
from datetime import datetime
import io
#import numpy as np


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

def notMatchedFieldName(sheet,index,fieldName):
    '''
    필드이름이 올바른지 확인한다. (true,false)
    '''
    if str(sheet[str(object=index)].value) != fieldName:
        return True
    else:
        return False;
