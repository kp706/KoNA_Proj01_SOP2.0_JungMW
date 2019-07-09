# -*- coding: utf-8 -*-
import gzip
import sys
import os

'''
Validation of fastq_file program
입력값 : fastq.gz 혹은 fq.gz 파일들이 들어있는 디렉토리
(상대경로든,절대경로든 상관 없지만, 폴더가 아닌 파일을 입력으로 넣으면 안된다.)
출력값 : input으로 들어온 디렉토리 내부에 있는 fastq.gz 혹은 fq.gz 파일의
validation 여부, Read개수, Base 개수

* 실행 방법 : python newtest.py "디렉토리경로"
'''


def search(dirname, outputfile_s):
    '''
    디렉토리를 받아서,내부 목록들을 스캔하고, fastq.gz파일인경우에는 validation함수를
    호출해서 작업을 시작하고, 또 다른 폴더인경우에는 재귀호출을 해주면서 모든 파일을 순회한다.
    '''
    filenames = os.listdir(dirname)  # 하위 항목들(파일 혹은 디렉토리) 저장
    for filename in filenames:  # 하위 항목들로 반복문
        # 파일이름에 경로 추가해서 절대경로 만들어주고
        full_filename = os.path.join(dirname, filename)
        if os.path.isdir(full_filename):  # 해당 항목이 디렉터리인경우 재귀호출
            search(full_filename)
        elif os.path.isfile(full_filename):  # 해당 항목이 파일인 경우
            # isgz = os.path.splitext(full_filename)[-1]  # 확장자명 얻어와서 테스트해보고
            ext = str(full_filename).rsplit('.', 2)[1]

        #    isgz = os.path.split(full_filename)[-1] #gz압축파일인경우
            if (ext == 'fastq' or ext == 'fq'):
                # 확장자명,gz파일까지 일치하면 validation 호출
                validation(full_filename, outputfile_s)


def validation(filepath, outputfile_v):  # 유효성 확인 함수
    '''
    validation 함수의 inputValue는 fastq.gz 또는 fq.gz이다. 압축은 풀지않고 모듈을 이용해서
    파일을 Read한다. 파일의 내부는 건들필요도없고 건들이면 안되기때문에 'rt'라는 인자를주어 읽기전용으로 open한다.
    그리고 헤더가 몇줄인지 계산하고, 헤더를 제외한 부분으로 line수, read수, base 수를 업데이트 해나가고,
    마지막에 line의 수가 4의 배수이면 Valid, 아니면 Not Valid로 처리해준다.
    '''
    try:
        target = gzip.open(filepath, 'rt')  # 읽기 전용으로 파일 연다.
        temptarget = gzip.open(filepath, 'rt')
        #lines = target.readlines()
        # lines에 text값들 복사된다.

        print(str(filepath) + " validation_processing...")
        numberOfRead = 0  # Read 개수 초기화
        numberOfBase = 0  # Base 개수 초기화
        numberOfLine = 0  # 헤더를 제외한 Line 개수 초기화
        length = 0  # 라인 하나하나의 길이를 임시로 보관할 변수
        headerCount = 0  # 헤더 줄 수를 저장할 변수

        for line in temptarget:  # +기호가 나올때까지 카운팅한다.
            headerCount += 1
            if line.strip() == "+":
                break
        headerCount -= 3  # 헤더를 계산한다.
        temptarget.close()  # 용도를 다했으니 파일 닫는다.

        k = 0
        while(k < headerCount):  # 헤더를 다 읽어서 버린다.
            line = target.getline()
            print(line)
            k += 1

        for line in target:  # 파일을 한줄한줄 순서대로 읽으며 반복문을 돈다.
            numberOfLine += 1  # 라인수 증가지시고
            if numberOfLine % 4 == 2:  # 한 Read 내에서 두번째줄인경우
                length = len(line.strip())  # strip함수로 마지막 개행문자 제거한뒤 길이계산
                numberOfRead += 1  # 4줄중에서 한번진입하므로, Read수 1 증가
                numberOfBase += length  # Base수는 sequence 길이만큼 증가

        print("File Name : " + str(filepath))  # 프로세싱한 전체파일경로 출력

        flag = (numberOfLine % 4 == 0)  # 라인의 개수가 4의 배수가 맞는지 검증해서 flag에 저장
        if flag:  # flag값에 따라 유효성 여부 출력
            print("This file is VALID")  # 4의 배수이면 VALID
        else:
            print("This file is NOT VALID")  # 4의배수가 아니면 NOT VALID
            # 왜 NOT VALID인지 Line의 수를 보여준다.
            print("number of line is " + str(numberOfLine))

        print(str(filepath) + " numberOfRead : " +
              str(numberOfRead))  # Read 수 출력
        print(str(filepath) + " numberOfBase : " +
              str(numberOfBase))  # Base 수 출력
        print()  # 결과값 출력해주고 개행
        target.close()  # 파일 닫는다.


        '''
        결과값들을 순서대로 outputfile에 적어준다.
        '''
        outputfile_v.write(str(filepath) + "\t") #파일명 적고 tab
        outputfile_v.write((str(filepath)).rsplit('\\',1)[1] + "\t") #파일이름 파싱해서 적고 tab
        if flag:
            outputfile_v.write("VALID" + "\t") #유효여부 적고 tab
        else:
            outputfile_v.write("NOT VALID" + "\t")
        outputfile_v.write(str(numberOfRead) + "\t") # Read개수 적고 tab
        outputfile_v.write(str(numberOfBase) + "\n")  # 마지막데이터쓰고 개행

    except IOError as err:
        print("Input file Error: " + str(err))


print()  # 개행
'''
main에 해당하는 부분으로, 경로를 받아와서 절대경로로 변환한 뒤 search 함수에 넣어준다.
'''
p = str(sys.argv[1])  # 첫번째 인자값에서 경로 받아와서
p = os.path.abspath(p)  # 절대경로로 만들어주고
outputfile = open("validation_result.txt", 'w')  # output Text파일 생성
outputfile.write("filePath" + "\t" + "fileName" + "\t" + "isValidation" + "\t" +
                 "numberOfRead" + "\t" + "numberOfBase" + "\n")
# 목차 적어준다.
search(p, outputfile)  # 함수 호출
outputfile.close() #output Text파일 닫고 저장
