import xml.etree.ElementTree as elemTree
import numpy as np
import os
import sys

xmlPath = str(sys.argv[1])  # 첫번째 인자값에서 경로 받아와서
xmlPath = os.path.abspath(xmlPath)  # 절대경로로 만들어주고

proj_id = str(sys.argv[2]) # 프로젝트 아이디 받아오고
contents = str(sys.argv[3]) #채울 내용 받아온다.


try:
    tree = elemTree.parse(xmlPath)
    root = tree.getroot()
    #xml파일 파싱

    for ele in root.findall("./project_set/first_data"):
        if ele.get('project_id') == proj_id:
            ele.set('sample_desc',contents)
            #prject_id 일치하는것 찾아서 sample_desc 수정

    targetXML = open(xmlPath,'w',encoding="UTF-8")
    tree.write(xmlPath, encoding="UTF-8")
    targetXML.close() #수정한내용 파일에 덮어쓰고 저장
except IOError as err:
    print("Xml File Error : " + str(err))
