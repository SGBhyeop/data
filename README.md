# data


각 파일 설명
gui.py : openpyxl 사용하는 gui 코드
analyze.py : openpyxl 사용하는 gui 없는 버전 코드 
ings.py : xlwings 사용하는 코드
실행파일(exe 파일)은 xlwings 기반 제작함

실행방법 (택1)
1. .exe 파일 실행 후 파일 선택 후 실행
2. exe 답답하면 파이썬 파일에서 경로 수정 후 직접 실행

실행 시 주의사항
1. 필요한 라이브러리
	pandas, openpyxl, os 
	gui의 경우 tkinter, sys 까지 필요
2. csv, txt 등 읽을 수 있음, 구분자 상관 없음
3. 파일과 같은 경로에 form.xlsx 파일 유지 필수
4. 파일과 같은 경로에 새로운 파일 생성
5. 창 하단부 삭제된 데이터 출력 

코드 구조
1. reading 함수
- 파일 읽기
- 0~ 데이터부터 읽기

2. preprocessing1 
- 중복 데이터 제거하는 함수
- 앞뒤 행 차이가 1이하면 뒷 행 제거
- 제거된 행 출력

3. preprocessing2
- 누락된 데이터 있으면 해당 그룹 제거
- 각 행이 몇 번 point 데이터인지 포인트 번호 지정
- 포인트 번호 기반으로 1~5 데이터면 df_index_valid에 저장
- 그 외 저장되지 않은 데이터는 제거된 데이터(누락되어서)로 간주
- 제거된 데이터 출력

4. point_extract
- 전처리 한 데이터를 각 포인트별로 리그룹

5. insert_data
- 준비된 엑셀파일에 리그룹한 데이터 삽입

6. 나머지 함수
- gui를 위한 함수 
