# Production Scheduler
MILP를 활용한 생산스케줄러

진행과정

interface 파이썬 파일 실행 -> 스케줄링 하기 선택 -> 제품정보 엑셀 선택 -> 생산계획정보 엑셀 선택
-> 데이터 불러오기 -> 데이터 구조화 -> 제약식 설정 -> 스케줄시작(생산 시작시간 설정 -> 생산계획 계산 시간)
-> 엑셀불러오기(결과엑셀파일 저장) -> 간트차트 불러오기(간트차트 파일 저장)

![image](https://github.com/user-attachments/assets/3f7dc571-1398-4b7a-bf94-d941f301d17d)





- 생산계획 엑셀 시트별 역할-
목록 : 필요한 정보를 읽기위한 시트명이 적혀 있음(라인명을 기준으로 행에 해당되는 각각의 시트정보를 읽음)
라인 사용여부 : 라인 비가동시 추가
우선작업 : 각 라인별 우선작업
생산정보 : 디스패칭을 수행할 작업 입력(연배, 수량, 양면,단면(단면의 경우 우선작업 없음))
초기셋업 : 각 라인별 초기 셋업(미입력시 초기셋업은 0)
작업시간 : 각 라인변 시작시간(조달기간 고려)
라인효율 : 각 라인별 효율 및 각 라인별 가동시간입력(가동시간은 최대한 지켜지게 만듬)
묶음생산 : 제품별 묶음생산 정보

- 제품정보 엑셀 시트별 역할-
목록 : 필요한 정보를 읽기위한 시트명이 적혀 있음(라인명을 기준으로 행에 해당되는 각각의 시트정보를 읽음)
제품정보 : 제품명, 전용라인, 선행작업 등의 정보 입력
작업시간 : 제품별 작업시간입력
셋업시간 : A->B 제품 변경간에 셋업시간


라인별 결과 엑셀 시트
![image](https://github.com/user-attachments/assets/81be9cc2-ff8d-428e-90c9-56fe3836e10d)

간트차트 결과 시각화

![image](https://github.com/user-attachments/assets/fe754609-9fe4-400f-b17a-d024cfaa7f4e)
