# PyExcelToSQLite

## 용도

개인적인 프로젝트에 쓰기 위한 스크립트이다. 어떤 환경에서도 돌아가게 하기 위해 만들었다.

```Excel```에 미리 정의된 Data Model을 ```SQLite DB```와 ```C#```으로 생성해준다.

## 사용한 도구

- Python 3.6
  - openpyxl
  - pystache (mustache)

## 설치

Python 3.6 이상 버전을 설치 후, 아래 구문을 실행한다.

> ```$ pip install -r requirements.txt```

## Excel Sheet 규칙

1. Sheet 이름의 맨 앞 글자에 ```_```가 들어가는 경우 DB에 삽입되지 않음.

- 예) ```_info```

2. Sheet에서 1~5번행은 항상 지켜져야 함.

- 1번행: 해당 열에 대한 설명
- 2번행: 해당 열의 용도 (사용되지 않을 예정)
- 3번행: 해당 열의 Attribute
- 4번행: 해당 열의 Data Type
- 5번행: 해당 열의 이름

3. 내부 데이터 값은 비어있지 않아야 함.

## 사용법

```$ PyExcelToSQLite.py -o ../master.db -p ../sources -e sample.xlsx```

1. ```sample.xlsx```와 같은 형식의 Excel(xlsx) 파일을 준비한다. (Repository에 있음)

2. 적절한 인자 값을 넣고 ```PyExcelToSQLite.py```을 실행한다. 인자를 넣는 순서는 상관없다.
- ```-o```, ```--output```: 생성된 DB의 경로 및 파일 이름
- ```-p```, ```--cspath```: 템플릿을 통해 출력된 C# 파일들의 저장 경로
- ```-e```, ```--excel```: Excel(xlsx) 경로와 파일 이름

3. 나온 결과물을 확인한다.

## 해야할 일

- SQLCipher 적용
- 규칙 개선 (너무 빡빡함)
- 코드 리펙토링
- Google Sheet 연동