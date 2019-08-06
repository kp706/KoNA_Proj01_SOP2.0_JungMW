KONA_Validation
===============
사용방법 : Shell창 입력: python kona_Validation.py "excel파일경로" 
----------------------------------------------------------------
<br>

![캡처](https://user-images.githubusercontent.com/46260961/62503784-aa808080-b830-11e9-87f6-75ff714d338d.PNG)

## BioProject Validation
* 제출된 과제명이 총괄이 아닌지 검사
* 날짜 입력 양식 검사 (YYYY-MM-DD)
* Release date selection 선택타입으로 바르게 입력했는지 검사
* Release on specified date를 선택한 경우 공개날짜를 입력했는지 검사
* Release date이 현재로부터 1년 이내인지 검사
* M/O 에 올바른 입력이 되어있는지 검사
* 7개 부처중 선택할때 예시에있는 부처들중에서 입력했는지 검사 <br>

## BioSample Validation
* 날짜 입력 양식 검사(YYYY-MM-DD)
* Release date selection 선택타입으로 바르게 입력했는지 검사
* Release on specified date를 선택한 경우 공개날짜를 입력했는지 검사
* Release date이 현재로부터 1년 이내인지 검사
* project accession 항목 기입 검사 <br>

## Experiment Validation
* Sample name 항목이 BioSample에 기입된 sample name 항목과 동일한지 검사
* Release date이 현재로부터 1년 이내인지 검사
* 'Fragment/Paired read' 필드가 Pairde-end인 경우, "insert size" 혹은 "Nominal size"필드 중 적어도 하나 이상의 필드에 값이 기입되었는지 검사<br>

