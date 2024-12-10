
import pandas as pd
import ortools.sat.python.cp_model as cp
from datetime import datetime, timedelta, time
import plotly.figure_factory as ff
import plotly.io as pio
import copy
import numpy as np
import openpyxl
import openpyxl.cell._writer


class productionScheduler:
    def __init__(self):
        #초기 변수 설정
        self.solver = cp.CpModel()
        self.infinity = 100000000
        self.senario = ['04-18']
        self.senario = self.senario[0]
        self.running_time = 1800
        self.work_time = {}
        self.Working_order = {}
        self.Working_order['결과없음'] = pd.DataFrame()
        self.rest_Working_order = {}
        
        
        self.all_protime = pd.DataFrame()
        self.all_setup_time = pd.DataFrame()
        self.dedicated_line = pd.DataFrame()
        self.job_type = pd.DataFrame()
        
        self.d_information = pd.DataFrame()
        
        self.all_production = pd.DataFrame()
        self.production = pd.DataFrame()
        self.start_set = pd.DataFrame()
        self.product_start_time = pd.DataFrame()
        self.efficiency = pd.DataFrame()
        self.job_sequence = {}
        
        self.order_processing = {}
        self.rest_order_processing = {}
        
        self.product_production_information = pd.DataFrame()
    
    def load_data(self,p_filepath,d_filepath):
        #제품정보 엑셀 데이터 불러오기
        self.p_information = pd.read_excel(p_filepath[0], p_filepath[1])
        self.all_protime = [pd.read_excel(p_filepath[0], i) for i in self.p_information['작업시간']]
        self.all_setup_time = [pd.read_excel(p_filepath[0], i) for i in self.p_information['셋업시간']]
        self.dedicated_line = pd.read_excel(p_filepath[0], self.p_information['전용라인'][0])
        self.job_type = pd.read_excel(p_filepath[0], self.p_information['선행작업'][0])
        self.product_df = pd.read_excel(p_filepath[0], sheet_name='제품정보')
        print(self.product_df)
        
        #주문 엑셀 데이터 불러오기
        self.d_information = pd.read_excel(d_filepath[0], d_filepath[1])
        self.all_production = pd.read_excel(d_filepath[0], sheet_name = self.d_information['생산정보'][0], usecols=[0, 1, 2, 3, 4, 5])
        self.all_production = self.all_production.dropna(subset=['제품명'])
        self.start_set = pd.read_excel(d_filepath[0], self.d_information['초기셋업'][0], index_col='라인명')
        self.product_start_time = pd.read_excel(d_filepath[0], self.d_information['작업시간'][0], index_col='라인명')
        self.efficiency = pd.read_excel(d_filepath[0], self.d_information['라인효율'][0], index_col='라인명')
        self.job_sequence = {self.d_information.loc[i,'라인명'] : pd.read_excel(d_filepath[0], self.d_information['우선작업'][i]) 
                        for i in range(len(self.d_information['우선작업']))}
        self.assemble_production = pd.read_excel(d_filepath[0], self.d_information['묶음생산'][0])
    
    def data_structures(self):
        #전용라인 리스트로 변환
        def convert_value_to_list(x, line_use):
            arr = x.split(',')
            arr = [i for i in arr if line_use[i]=='사용']
            return arr
        
        #더미작업 추가로 인한 셋업시간 행열에 0추가
        def add_0(s, start_setup, setup_dict, product, output_dict):
            num_cols = len(s[0]) if s else 0
            zero_row = [0] * (num_cols + 2)
            if start_setup in setup_dict:
                s.insert(0, [0] + [setup_dict[start_setup][j] if j in setup_dict[start_setup] else 1000000 for j in product if output_dict[j]['생산여부'] == 1] + [0])
            else:
                s.insert(0, zero_row)
            s.append(zero_row)
            
            for idx, row in enumerate(s):
                if idx != 0 and idx != len(s) - 1:
                    row.insert(0, 0)
                    row.append(0)
            return s
        
        
        self.line_name = self.p_information['라인명']
        self.all_line = self.line_name
        self.protime = {line : self.all_protime[idx] for idx,line in enumerate(self.line_name)}
        self.setup_time = {line : self.all_setup_time[idx] for idx,line in enumerate(self.line_name)}
        self.line_use = {line : self.d_information.loc[self.d_information['라인명']==line, '라인사용여부'].iloc[0] for line in self.d_information['라인명']}
        self.line_name = [i for i in self.line_name if self.line_use[i] == '사용']
        self.line_num = len(self.line_name)
        
        self.all_production['총생산량'] = self.all_production['양면생산량'] + self.all_production['단면생산량']
        self.production = self.all_production.copy()
        
        self.production['생산여부'] = 0
        self.production.loc[self.production['단면생산량'] > 0,'생산여부'] = 1
        self.production.loc[self.production['양면생산량'] > 0,'생산여부'] = 1
        
        
        test_product = list(self.production['제품명'])
        #중복 및 빈값 오류 탐색
        for i in test_product:
            if not i in list(self.product_df['제품명']):
                raise TypeError(f"{i} 제품입력 엑셀파일의 제품정보시트에 존재하지 않음")
            product_line = self.product_df.loc[self.product_df['제품명'] == i, '라인명']
            product_line = [item for sublist in [i.split(',') for i in product_line] for item in sublist]
            
            for line in product_line:
                if not i in list(self.protime[line]['제품명']):
                    raise TypeError(f"{i} 제품입력 엑셀파일의 {line}작업시간에 존재하지 않음")
            for line in product_line:
                if not i in list(self.setup_time[line]['시작모델']):
                    raise TypeError(f"{i} 제품입력 엑셀파일의 {line}셋업시간에 존재하지 않음")
        
                    
        start = self.start_set['작업'].dropna().to_dict()
        for line, j in start.items():
            if not j in list(self.setup_time[line]['시작모델']):
                raise TypeError(f"초기셋업 {i} 제품입력 엑셀파일의 {line}셋업시간에 존재하지 않음")
            
        
        
        
        try:
            self.output_dict = self.production.set_index("제품명").to_dict(orient="index")
        except ValueError as e:
            raise TypeError("생산량 중복값 존재 오류") from e
        
        
        self.protime_dict = {}
        for line in self.line_name:
            try:
                arr=[]
                arr= self.protime[line].set_index("제품명").to_dict(orient="index")
                self.protime_dict[line] = arr
            except ValueError as e:
                raise TypeError(f"{line}작업시간 중복값 존재 오류") from e
        
        
        self.setup_dict = {}
        for line in self.line_name:
            try:
                arr=[]
                arr= self.setup_time[line].set_index("시작모델").to_dict(orient="index")
                self.setup_dict[line] = arr
            except ValueError as e:
                raise TypeError(f"{line}셋업시간 중복값 존재 오류") from e
        
        
        try:
            self.line_dict = self.dedicated_line.set_index("제품명").to_dict(orient="index")
        except ValueError as e:
            raise TypeError("전용라인에 중복값 존재") from e
        
        # 분리 비분리 여부를 통해 선행작업 적용여부 결정
        self.production['작업유형'] = '비분리'
        original_indices = self.production[self.production['분리여부'] == '분리'].index
        rows_to_duplicate = self.production.loc[original_indices]
        
        if len(rows_to_duplicate) !=0 or len(original_indices) !=0:
            self.production = pd.concat([self.production, rows_to_duplicate]).reset_index(drop=True)
            
            copied_indices = self.production.index[-len(original_indices):]
            
            self.production.loc[original_indices, '단면생산량'] = 0
            self.production.loc[original_indices, '작업유형'] = '양면'
            
            self.production.loc[copied_indices, '양면생산량'] = 0
            self.production.loc[copied_indices, '작업유형'] = '단면'
            
            self.production['생산여부'] = 0
            self.production.loc[self.production['단면생산량'] > 0,'생산여부'] = 1
            self.production.loc[self.production['양면생산량'] > 0,'생산여부'] = 1
        
        self.job_num = sum(self.production['생산여부'])
        
        self.product = self.production['제품명'].tolist()
        self.product_list =[[0]]+ [[i,k] for i, j, k in zip(self.product, self.production['생산여부'].tolist(), self.production['작업유형'].tolist())  for j in range(j)] +[[0]]
        
        #생산시간 설정
        self.p={}
        for line in self.line_name:
            arr = [0]
            for i in self.product_list[1:-1]:
                for _ in range(self.output_dict[i[0]]['생산여부']):
                    try:
                        if i[0] in self.protime_dict[line] and  i[1] == '양면':
                            arr.append(round(self.protime_dict[line][i[0]]['작업시간']*self.output_dict[i[0]]['양면생산량']))
                        elif i[0] in self.protime_dict[line] and  i[1] == '단면':
                            arr.append(round(self.protime_dict[line][i[0]]['작업시간']*self.output_dict[i[0]]['단면생산량']))
                        elif i[0] in self.protime_dict[line]:
                            arr.append(round(self.protime_dict[line][i[0]]['작업시간']*(self.output_dict[i[0]]['단면생산량']+self.output_dict[i[0]]['양면생산량'])))
                        else:
                            arr.append(10000000)
                    except:
                        raise TypeError(f"{line}작업시간 시트 오류")
            
            arr.append(0)
            self.p[line] = arr
        
        #라인 효율 적용
        for line in self.line_name:
            for i in range(1, self.job_num + 1):
                self.p[line][i] = round(self.p[line][i] / self.efficiency.loc[line,'라인효율'])
        
        #각 라인별 가동시간 설정
        for line in self.line_name:
            for i in range(1, self.job_num + 1):
                self.work_time[line] = round(self.efficiency.loc[line,'각 라인별 가동시간']) * 60 * 60
        
        #셋업시간 설정
        self.s = {}
        for line in self.line_name:
            try:
                arr=[]
                arr= add_0([[round(self.setup_dict[line][i][j]) if i in self.setup_dict[line] and j in self.setup_dict[line] else 1000000 for j in self.product for _ in range(self.output_dict[j]['생산여부'])] for i in self.product for _ in range(self.output_dict[i]['생산여부'])],
           str(self.start_set.loc[line, '작업']), self.setup_dict[line], self.product, self.output_dict)
                self.s[line] = arr
            except:
                raise TypeError(f"{line}셋업시간 처리중 오류 발생")
        
        
        #선행작업 설정
        first_column_values = [item[0] for item in self.product_list]
        try:
            self.r=[[-1]]
            for i in range(1, self.job_num+1):
                n_p = self.product_list[i][0]
                r_job=self.job_type.loc[self.job_type['제품명']== n_p, '선행작업'].iloc[0]
                if r_job in self.output_dict and self.output_dict[r_job]['양면생산량']>0:
                    self.r.append([first_column_values.index(r_job),
                              {line : round(self.protime_dict[line][n_p]['작업시간'] * self.output_dict[n_p]['단면생산량']) if n_p in self.protime_dict[line] else -10000 for line in self.line_name}])
                else:
                    self.r.append([-1])
            self.r = self.r+[[-1]]
        except:
            raise TypeError("선행작업 시트 오류")
        
        #작업별 시작시간 설정
        try:
            self.start_time = []
            
            for line, row in self.product_start_time.iterrows():
                if row['작업명'] in first_column_values:
                    if row['분리여부'] == '분리' and row['작업유형'] == '단면':
                        reversed_index = first_column_values[::-1].index(row['작업명'])
                        original_index = len(first_column_values) - 1 - reversed_index
                        self.start_time.append([original_index, row['시작시간(단위 : 분)'] , line.split(',')])
                    else:
                        self.start_time.append([first_column_values.index(row['작업명']), row['시작시간(단위 : 분)'], line.split(',')])
        except:
            raise TypeError("작업시간 시트 오류")
        
        
        #우선작업설정
        self.sequence = {}
        first_column_values = [item[0] for item in self.product_list]
        for line, job in self.job_sequence.items():
            try:
                arr=[]
                current_job = 0
                for idx, row in enumerate(job['작업명']):
                    if row in first_column_values:
                        if job.loc[idx,'분리여부'] == '분리' and job.loc[idx,'작업유형'] == '단면':
                            reversed_index = first_column_values[::-1].index(row)
                            original_index = len(first_column_values) - 1 - reversed_index
                            arr.append([current_job, original_index])
                            current_job = original_index
                        else:
                            arr.append([current_job, first_column_values.index(row)])
                            current_job =  first_column_values.index(row)
                self.sequence[line] = arr
            except:
                raise TypeError(f"{line} 우선작업 시트 오류")
        
        
        #묶음생산 설정
        try:
            self.ass_list = []
            assemble_production = self.assemble_production
            assemble_list = assemble_production.groupby('묶음생산')['제품명'].apply(list).tolist()
            for i in assemble_list:
                arr=[]
                for j in first_column_values:
                    if j in i:
                        arr.append(first_column_values.index(j))
                self.ass_list.append(arr)
            print(self.ass_list)
        except:
            raise TypeError("묶음생산 시트 오류")
        
        try:
            self.line_list = [-1] + [convert_value_to_list(self.line_dict[i[0]]['라인명'], self.line_use) for i in self.product_list if i[0] != 0]  + [-1]
        except:
            raise TypeError("전용라인 시트 오류")
        
        
        
        for idx,i in enumerate(self.line_list):
            if i ==[]:
                raise TypeError(f"{self.product_list[idx][0]} 할당가능 라인의 부재로 작업배치불가")
        
        
        
        p_list = [i[0] for i in self.product_list[1:-1]]
        line_list = self.line_list[1:-1]
        re_list = [self.product_list[i[0]][0] for i in self.r[1:-1]]
        
        
        self.product_production_information = pd.DataFrame([p_list, line_list ,re_list]).T
        self.product_production_information.columns = ['제품명', '전용라인', '선행작업']
        self.product_production_information =  self.product_production_information.sort_values(by='제품명')
        
        
        familiy_product_map = dict(zip(self.assemble_production['제품명'], self.assemble_production['묶음생산']))
        
        self.product_production_information['묶음생산'] = self.product_production_information['제품명'].map(familiy_product_map)
        
    
    #제약조건 설정
    def setup_constraints(self):
        job_num = self.job_num
        line_name = self.line_name
        line_num = self.line_num
        line_list = self.line_list
        sequence = self.sequence
        work_time = self.work_time
        p = self.p
        s = self.s
        r = self.r
        start_time = self.start_time
        ass_list = self.ass_list
        
        #x[i][j] i작업후에 j작업을 수행하면 1 아니면 0
        self.x = [[self.solver.NewIntVar(0, 1, f'self.x{i}_{j}') for j in range(job_num + 2)] for i in range(job_num + 2)]
        #y[l][i] l라인에서 i작업을 수행하면 1 아니면 0
        self.y = {line : [self.solver.NewIntVar(0, 1, f'self.y{line}_{i}') for i in range(job_num + 2)] for line in line_name}
        #w[l][i] 라인l의 시작작업이 i면 1 아니면 0
        self.w =  {line : [self.solver.NewIntVar(0, 1, f'self.w{line}_{i}') for i in range(job_num + 2)] for line in line_name}
        #st[i] i작업의 시작시간
        self.st = [self.solver.NewIntVar(0, self.infinity, f'self.st{i}') for i in range(job_num + 2)]
        #e[l][i] 라인 l의 종료작업이면 1 아니면 0
        self.e = {line : [self.solver.NewIntVar(0, 1, f'self.e{line}_{i}') for i in range(job_num + 2)] for line in line_name}
        #l_endtime[l][t] 라인 l의 종료시간
        self.l_endtime = {line : self.solver.NewIntVar(0, self.infinity, f'self.l_endtime{line}') for line in line_name}
        #l_p 라인 l의 초과 근무 수행시간
        self.l_p = {line : self.solver.NewIntVar(0, self.infinity, f'self.l_p{line}') for line in line_name}
        #x_y[l][i][j] y[l][i]와 x[i][j]가 모두 1일 때만 1
        self.x_y = {line : [[self.solver.NewIntVar(0, 1, f'self.x_y{line}_{i}_{j}') for j in range(job_num + 2)] for i in range(job_num + 2)] for line in line_name}
        #초기작업은 라인 수 이하만큼 배정된다.
        self.solver.Add(sum(self.x[0][j] for j in range(1, job_num + 2)) <= line_num)
        
        # 각 작업은 한 번만 수행되어야 함
        for i in range(1, job_num + 1):
            self.solver.Add(sum(self.x[i][j] for j in range(job_num + 2)) == 1)
        for j in range(1, job_num + 1):
            self.solver.Add(sum(self.x[i][j] for i in range(job_num + 1)) == 1)
        # 작업의 '선행'과 '후행'이 존재
        for j in range(1, job_num + 1):
            self.solver.Add(sum(self.x[i][j] for i in range(0, job_num + 1)) - sum(self.x[j][k] for k in range(1, job_num + 2)) == 0)
        # 같은 작업은 배정될 수 없음
        for i in range(job_num + 2):
            self.solver.Add(self.x[i][i] == 0)
        
        # 리엔트런트 작업은 선행작업 이후 진행 될 수 있음(현재는 선행작업 시작후 선행작업을 마친 재고를 고려하여 1시간 후에 시작되게 되어있음)
        for i in range(1, job_num+2):
            for line in line_name:
                if r[i][0] > 0 and r[i][1][line] != -10000:
                    self.solver.Add(self.st[i] >= self.st[r[i][0]] + 3600 - r[i][1][line]).OnlyEnforceIf(self.y[line][i])
        
        # 작업 간의 시간 관계 설정
        for i in range(0,job_num + 1):
            for j in range(1, job_num + 2):
                for line in line_name:
                    if i != j:
                        if j == job_num+1:
                            self.solver.Add(self.st[j] >= self.st[i] + s[line][i][j] + p[line][i]).OnlyEnforceIf(self.x[i][j],self.y[line][i])
                        else:
                            if r[j] == [-1]:
                                self.solver.Add(self.st[j] == self.st[i] + s[line][i][j] + p[line][i]).OnlyEnforceIf(self.x[i][j],self.y[line][j])
                            else:
                                self.solver.Add(self.st[j] >= self.st[i] + s[line][i][j] + p[line][i]).OnlyEnforceIf(self.x[i][j],self.y[line][j])
                                self.solver.Add(self.st[j] <= self.st[i] + s[line][i][j] + p[line][i]+3600).OnlyEnforceIf(self.x[i][j],self.y[line][i],self.y[line][j])
        
        # 각 작업은 한 라인에 1번만 할당됨
        for line in line_name:
            for i in sequence[line]:
                if len(i)>1:
                    self.solver.Add(self.x[i[0]][i[1]] == 1)
                    self.solver.Add(self.y[line][i[1]] == 1)
        
        for i in range(1, job_num+1):
            self.solver.Add(sum(self.y[line][i] for line in line_name) == sum(self.x[i][j]for j in range(job_num + 2)))
        
        # 초기 작업 할당 제약 조건(초기 작업은 각 라인에 하나씩만 할당됨)
        for line in line_name:
            for i in range(1, job_num + 1):
                self.solver.Add(self.w[line][i] == 1).OnlyEnforceIf(self.x[0][i], self.y[line][i])
            self.solver.Add(sum(self.w[line][i] for i in range(1, job_num + 1)) == 1)
        
        # 라인별 작업 제약 조건
        for line in line_name:
            for i in range(job_num+1):
                for j in range(job_num+1):
                    if i != j:
                        self.solver.Add(self.y[line][j] == 1).OnlyEnforceIf(self.x[i][j], self.y[line][i])
        
        # 전용 라인 제약 조건
        for i in range(1, job_num+1):
            if len(line_list[i]) < line_num:
                self.solver.Add(sum(self.y[line][i] for line in line_list[i]) == 1)
        
        #종료작업은 더미작업으로 배정됨
        for line in line_name:
            for i in range(1, job_num+1):
                self.solver.Add(self.e[line][i] == 1).OnlyEnforceIf(self.x[i][job_num+1], self.y[line][i])
        
        #라인별 종료시간은 마지막 작업의 종료시간임
        for line in line_name:
            for i in range(1, job_num+1):
                self.solver.Add(self.l_endtime[line] == self.st[i]+p[line][i]).OnlyEnforceIf(self.e[line][i])
        
        #각 라인의 초과 근무시간
        for line in line_name:
            self.solver.Add(self.l_p[line] >= self.l_endtime[line] - work_time[line])
        
        #각 작업은 시작시간 이후에 수행가능
        for i in start_time:
            self.solver.Add(self.st[i[0]] >= i[1]*60)
            self.solver.Add(sum(self.y[line][i[0]] for line in i[2]) == 1)
        
        #x_y제약 조건
        for line in line_name:
            for i in range(job_num+1):
                for j in range(job_num+2):
                    self.solver.Add(self.x_y[line][i][j] == 1).OnlyEnforceIf(self.x[i][j], self.y[line][i])
                    
        
        #묶음생산 제약조건
        for j_f in ass_list:
            if j_f != []:
                self.solver.Add(sum(self.x[i][j] for i in j_f for j in j_f) == len(j_f)-1)


        #우선적으로 각 라인별 가동시간을 넘기지 않게 셋업시간을 최소로 하게 스케줄링
        self.solver.Minimize(sum(self.x_y[line][i][j] * s[line][i][j] for line in line_name for i in range(job_num+1) for j in range(1,job_num+1))
                        + 100 * sum(self.l_p[line] for line in line_name)
                        + 0.001 * sum(self.l_endtime[line] for line in line_name)
                        + 0.001*self.st[job_num+1]
                        + sum(self.st[i[0]] for i in start_time))
        
        
        
        print('문제를 설정했습니다.')
        print(datetime.now())
    
    def solve(self):
        #문제 수행
        self.solver_cp = cp.CpSolver()
        self.solver_cp.parameters.max_time_in_seconds = self.running_time
        status = self.solver_cp.Solve(self.solver)
        print(datetime.now())
        if status == cp.OPTIMAL:
            print('최적해를 구했습니다.')
        else:
            print("최적해를 못구했습니다.")
    
    def output_results(self):
        #결과 정리
        makespan = max([self.solver_cp.Value(self.l_endtime[line]) for line in self.line_name])
        
        print(f"전체 완료 시간 (makespan): {makespan}")
        for i in range(self.job_num + 2):
            for j in range(self.job_num + 2):
                if self.solver_cp.Value(self.x[i][j]) > 0:
                    print(f'x{i}_{j} = {self.solver_cp.Value(self.x[i][j])}')
        for line in self.line_name:
            for i in range(self.job_num + 2):
                if self.solver_cp.Value(self.y[line][i]) > 0:
                    print(f'{line}_{i} = {self.solver_cp.Value(self.y[line][i])}')
        
        s_s = 0
        for i in range(0,self.job_num + 1):
            for j in range(1, self.job_num + 2):
                for line in self.line_name:
                    s_s += self.solver_cp.Value(self.x_y[line][i][j]) * self.s[line][i][j]
        print(s_s)
        
        
        for i in range(1, self.job_num + 1):
            print(f'st{i} = {self.solver_cp.Value(self.st[i])}')
        result_dict = {}
        
        for i in range(self.job_num + 2):
            for j in range(self.job_num + 2):
                if self.solver_cp.Value(self.x[i][j]) > 0:
                    result_dict.setdefault(i, []).append(j)
        
        #순서파악 함수 각 라인별 초기작업을 통해 라인에 작업 배부
        def trace_job_sequence(start_job, result_dict):
            line = [start_job]
            current_job = result_dict[start_job][0]
            while current_job in result_dict:
                line.append(current_job)
                next_job = result_dict[current_job][0]
                current_job = next_job
            return line
        
        #라인별로 작업 배정 딕셔너리
        self.lines = {}
        for start_job in result_dict[0]:
            for line in self.line_name:
                if self.solver_cp.Value(self.y[line][start_job]) > 0:
                    self.lines[line] = trace_job_sequence(start_job, result_dict)
        
        for line, sl in self.lines.items():
            print(line, sl) 
        for line in self.line_name:
            print(f'{line} = {self.solver_cp.Value(self.l_endtime[line])}')
        for line in self.line_name:
            for i in range(self.job_num + 2):
                if self.solver_cp.Value(self.e[line][i]) > 0:
                    print(f'{line}_{i} = {self.solver_cp.Value(self.e[line][i])}')
        
        self.Working_order = {line : pd.DataFrame(columns=['모델명', '시작시간', '종료시간']) for line in self.all_line}
        for line in self.all_line:
            try:
                for idx, i in enumerate(self.lines[line]):
                    self.Working_order[line].loc[idx] = [self.product_list[i][0], self.solver_cp.Value(self.st[i]), self.solver_cp.Value(self.st[i]) + self.p[line][i]]
            except:
                self.Working_order[line].loc[0] = ['제품 없음', 0, 0]
        
        
        self.rest_Working_order= copy.deepcopy(self.Working_order)
        
        output_excel_Working_Order = {}
        product_assemble_num = self.all_production.set_index('제품명')['연배'].to_dict()
        protime_dict = {line : self.protime[line].set_index('제품명')['작업시간'] for line in self.all_line}
        production_dict = self.all_production.set_index('제품명')['계획수량'].to_dict()
        customer_dict = self.product_df.set_index('제품명')['고객사'].to_dict()
        part_number_dict = self.product_df.set_index('제품명')['파트넘버'].to_dict()
        
        purpose_production = self.all_production.set_index('제품명')['총생산량'].to_dict()
        
        
        for line in self.all_line:
            temp_df = self.Working_order[line][['모델명', '시작시간', '종료시간']].copy()
            
            additional_columns = ['생산예상시간', 'M/C(분)', '생산시간(분)', '연배', 'T/T', 'UPH', '매수', '목표수량']
            for col in additional_columns:
                temp_df[col] = 0  
            
            output_excel_Working_Order[line] = temp_df
        for line in output_excel_Working_Order.keys():
            
            time_diff = output_excel_Working_Order[line]['시작시간'] - np.roll(output_excel_Working_Order[line]['종료시간'], 1)
            time_diff_in_minutes = time_diff // 60  
            
            time_diff_in_minutes[0] = 0  
            
            output_excel_Working_Order[line]['P/N'] = output_excel_Working_Order[line]['모델명'].map(part_number_dict)
            output_excel_Working_Order[line]['고객사'] = output_excel_Working_Order[line]['모델명'].map(customer_dict)
            output_excel_Working_Order[line]['M/C(분)'] = time_diff_in_minutes
            output_excel_Working_Order[line]['생산시간(분)'] = (output_excel_Working_Order[line]['종료시간']-output_excel_Working_Order[line]['시작시간'])//60
            output_excel_Working_Order[line]['연배'] = output_excel_Working_Order[line]['모델명'].map(product_assemble_num)
            output_excel_Working_Order[line]['T/T'] = output_excel_Working_Order[line]['모델명'].map(protime_dict[line])
            output_excel_Working_Order[line]['UPH'] = round((3600 / output_excel_Working_Order[line]['T/T'] * output_excel_Working_Order[line]['연배'])* self.efficiency.loc[line,'라인효율'])
            output_excel_Working_Order[line]['매수'] = output_excel_Working_Order[line]['모델명'].map(purpose_production)
            output_excel_Working_Order[line]['목표수량'] = output_excel_Working_Order[line]['모델명'].map(production_dict)
            
            
            output_excel_Working_Order[line]['시작시간'] = pd.to_datetime(output_excel_Working_Order[line]['시작시간'], unit='s', origin=f'{self.senario[0]}-{self.senario[1]} {self.senario[2]}:00')
            output_excel_Working_Order[line]['종료시간'] = pd.to_datetime(output_excel_Working_Order[line]['종료시간'], unit='s', origin=f'{self.senario[0]}-{self.senario[1]} {self.senario[2]}:00')
            output_excel_Working_Order[line]['생산예상시간'] = output_excel_Working_Order[line]['시작시간'].dt.strftime('%H:%M:%S') + '-' + output_excel_Working_Order[line]['종료시간'].dt.strftime('%H:%M:%S')
        
        self.production_setup_result_dict = {line : output_excel_Working_Order[line].set_index('모델명')['M/C(분)'] for line in self.all_line}
        for line in output_excel_Working_Order.keys():
            output_excel_Working_Order[line] = output_excel_Working_Order[line].rename(columns={'시작시간': '계획 시작시간', '종료시간': '계획 종료시간'})
        
        columns = ['모델명', '고객사', 'P/N', '계획 시작시간', '실제 시작시간', '계획 종료시간', 
           '실제 종료시간', '생산예상시간', 'T/T', 'UPH', '생산시간(분)', '실제 생산시간(분)', 
           'M/C(분)', '연배', '매수', '목표수량', '실적', '비고']
        
        new_working_order = {line : i.reindex(columns=columns) for line,i in output_excel_Working_Order.items()}
        
        for line in new_working_order.keys():
            new_working_order[line].insert(0, line, range(1, len(new_working_order[line]) + 1))
        
        #self.output_excel_Working_Order = new_working_order

        print(self.line_use)
        over_working_order = copy.deepcopy(new_working_order)
        for line in over_working_order.keys():
            if self.line_use[line] == '사용':
                target_time = pd.to_datetime(self.work_time[line], unit='s', origin=f'{self.senario[0]}-{self.senario[1]} {self.senario[2]}:00')
                over_working_order[line] = over_working_order[line][over_working_order[line]['계획 종료시간'] > target_time]
                # 타겟 시간 전에 시작하고, 타겟 시간 이후에 종료되는 작업 찾기
                df_target = over_working_order[line][over_working_order[line]['계획 시작시간'] <= target_time]
                if len(df_target)==1:
                    selected_task = df_target.iloc[0]  # 첫 번째(유일한) 행 선택
                    total_duration = selected_task['계획 종료시간'] - selected_task['계획 시작시간']
                    #완료 작업시간
                    elapsed_time = target_time - selected_task['계획 시작시간']
                    #미완료 작업시간
                    remaining_time = selected_task['계획 종료시간'] - target_time
                    #작업 완료비율
                    elapsed_ratio = elapsed_time / total_duration
                    #작업 미완료비율
                    remaining_ratio = remaining_time / total_duration
                    
                    df_target['계획 시작시간'] = target_time
                    df_target['생산예상시간'] = df_target['계획 시작시간'].dt.strftime('%H:%M:%S') + '-' + df_target['계획 종료시간'].dt.strftime('%H:%M:%S')
                    df_target['생산시간(분)'] = ((df_target['계획 종료시간'] - df_target['계획 시작시간']).dt.total_seconds() / 60).astype(int)
                    
                df_target.reindex(columns=columns)
                
                if len(df_target)==1:
                    over_working_order[line].iloc[0] = df_target.iloc[0]
                
                over_working_order[line]['매수'] = round(
                                                            (pd.to_datetime(over_working_order[line]['계획 종료시간']) - 
                                                             pd.to_datetime(over_working_order[line]['계획 시작시간'])).dt.total_seconds() /
                                                        over_working_order[line]['T/T'] * self.efficiency.loc[line,'라인효율']) 
                over_working_order[line]['목표수량'] = round(over_working_order[line]['매수'] * over_working_order[line]['연배'])

        # 추가할 문자열
        add_str = '_초과작업'
        
        # over_working_order의 키-값 쌍을 new_working_order에 추가
        
        for key, value in over_working_order.items():
            new_key = str(key)+add_str  # 새로운 키 이름 생성
            new_working_order[new_key] = value  # 새 키와 기존 값을 new_working_order에 추가
        
        self.output_excel_Working_Order = new_working_order

        
    
    def output_excel(self,Working_order):
        #for line, seq in Working_order.items():
         #   seq.to_csv(f'{line}_{self.senario[1]}_풀타임.csv', encoding='utf-8-sig', index=False)
        # ExcelWriter를 초기화하고 xlsxwriter 엔진을 사용합니다.
        with pd.ExcelWriter(f'{self.senario[1]} 생산계획.xlsx', engine='xlsxwriter') as writer:
            for line, seq in Working_order.items():
                # 각 데이터 프레임을 엑셀 시트로 저장
                seq.to_excel(writer, sheet_name=line, index=False)
        
                # 현재 작업 중인 워크시트 객체를 가져옵니다.
                worksheet = writer.sheets[line]
        
                # 각 열에 대해 최대 너비를 계산하여 설정
                for idx, col in enumerate(seq):  # idx는 열의 인덱스, col은 열의 이름
                    # 열 데이터 중 가장 긴 것의 길이를 측정합니다.
                    column_len = seq[col].astype(str).map(len).max()
                    # 열 이름의 길이도 고려합니다.
                    column_len = max(column_len, len(col)) + 1  # 여유분을 더합니다.
                    # 계산된 너비로 열 너비를 설정합니다.
                    worksheet.set_column(idx, idx, column_len)

    
    def rest_combine(self):
        for line in self.rest_Working_order.keys():
            df = self.rest_Working_order[line]
            
            base_time_str = '08:30'
            
            break_times = [(7200, 7200+600), (13200, 13200+3600), (25200, 25200+600),
                           (32400, 32400+1800), (48600, 48600+600), (55800, 55800+3600),
                           (66600, 66600+600), (73800, 73800+600), (81000, 81000+600),
                           (84600, 84600+1800),
                           (93600, 93600+600), (99600, 99600+3600), (111600, 111600+600),
                           (118800, 118800+1800), (135000, 135000+600), (142200, 142200+3600),
                           (153000, 153000+600), (160200, 160200+600), (167400, 167400+600),
                           (171000, 171000+1800),
                           (180000, 180000+600), (186000, 186000+3600), (198000, 198000+600), 
                           (205200, 205200+1800), (221400, 221400+600), (228600, 228600+3600), 
                           (239400, 239400+600), (246600, 246600+600), (253800, 253800+600), 
                           (257400, 257400+1800),
                           (266400, 266400+600), (272400, 272400+3600), (284400, 284400+600), 
                           (291600, 291600+1800), (307800, 307800+600), (315000, 315000+3600), 
                           (325800, 325800+600), (333000, 333000+600), (340200, 340200+600), 
                           (343800, 343800+1800),
                           (352800, 352800+600), (358800, 358800+3600), (370800, 370800+600), 
                           (378000, 378000+1800), (394200, 394200+600), (401400, 401400+3600), 
                           (412200, 412200+600), (419400, 419400+600), (426600, 426600+600), 
                           (430200, 430200+1800)]
            
            current_work_start = datetime.strptime(self.senario[2], "%H:%M")
            base_time = datetime.strptime(base_time_str, "%H:%M")
            
            time_diff = (current_work_start - base_time).total_seconds()
            
            adjusted_break_times = [(start - time_diff, end - time_diff) for start, end in break_times]
            
            break_times = adjusted_break_times
            
            new_rows = []
            total_time = df.iloc[0]['시작시간']  
            for i in range(len(df)):
                row = df.iloc[i]
                
                if i == 0:  
                    ProcTime = row['종료시간'] - row['시작시간']
                    SUTime = 0
                else:
                    previous_row = df.iloc[i-1]
                    ProcTime = row['종료시간'] - row['시작시간']
                    SUTime = row['시작시간'] - previous_row['종료시간']
                
                for break_start, break_end in break_times:
                    if total_time < break_start < total_time + SUTime:  
                        new_rows.append({'모델명': '셋업', '시작시간': total_time, '종료시간': break_start})
                        SUTime -= break_start - total_time  
                        total_time = break_end
                        new_rows.append({'모델명': '휴식', '시작시간': break_start, '종료시간': break_end})
                    elif break_start <= total_time < break_end:  
                        total_time = break_end
                        new_rows.append({'모델명': '휴식', '시작시간': break_start, '종료시간': break_end})
                
                total_time += SUTime
                new_rows.append({'모델명': '셋업', '시작시간': total_time - SUTime, '종료시간': total_time})
                
                for break_start, break_end in break_times:
                    if total_time < break_start < total_time + ProcTime:  
                        new_rows.append({'모델명': row['모델명'], '시작시간': total_time, '종료시간': break_start})
                        ProcTime -= break_start - total_time  
                        total_time = break_end
                        new_rows.append({'모델명': '휴식', '시작시간': break_start, '종료시간': break_end})
                    elif break_start <= total_time < break_end:  
                        total_time = break_end
                        new_rows.append({'모델명': '휴식', '시작시간': break_start, '종료시간': break_end})
                
                total_time += ProcTime
                new_rows.append({'모델명': row['모델명'], '시작시간': total_time - ProcTime, '종료시간': total_time})
            
            new_df = pd.DataFrame(new_rows)
            new_df = new_df.drop(new_df.index[0]) 
            self.rest_Working_order[line] = new_df
        
        rest_order_processing = copy.deepcopy(self.rest_Working_order)
        
        
        
        output_excel_Working_Order = {}
        product_assemble_num = self.all_production.set_index('제품명')['연배'].to_dict()
        protime_dict = {line : self.protime[line].set_index('제품명')['작업시간'] for line in self.all_line}
        production_dict = self.all_production.set_index('제품명')['계획수량'].to_dict()
        part_number_dict = self.product_df.set_index('제품명')['파트넘버'].to_dict()
        customer_dict = self.product_df.set_index('제품명')['고객사'].to_dict()
        
        purpose_production = self.all_production.set_index('제품명')['총생산량'].to_dict()
        
        for line in self.all_line:
            temp_df = rest_order_processing[line][['모델명', '시작시간', '종료시간']].copy()
            
            additional_columns = ['P/N', '생산예상시간', 'M/C(분)', '생산시간(분)', '연배', 'T/T', 'UPH', '매수', '목표수량']
            for col in additional_columns:
                temp_df[col] = 0  
            
            output_excel_Working_Order[line] = temp_df
        
        for line in output_excel_Working_Order.keys():
            output_excel_Working_Order[line]['P/N'] = output_excel_Working_Order[line]['모델명'].map(part_number_dict)
            output_excel_Working_Order[line]['고객사'] = output_excel_Working_Order[line]['모델명'].map(customer_dict)
            output_excel_Working_Order[line]['M/C(분)'] = output_excel_Working_Order[line]['모델명'].map(self.production_setup_result_dict[line])
            output_excel_Working_Order[line]['생산시간(분)'] = (output_excel_Working_Order[line]['종료시간']-output_excel_Working_Order[line]['시작시간'])//60
            output_excel_Working_Order[line]['연배'] = output_excel_Working_Order[line]['모델명'].map(product_assemble_num)
            output_excel_Working_Order[line]['T/T'] = output_excel_Working_Order[line]['모델명'].map(protime_dict[line])
            output_excel_Working_Order[line]['UPH'] = round((3600 / output_excel_Working_Order[line]['T/T'] * output_excel_Working_Order[line]['연배']) * self.efficiency.loc[line,'라인효율'])
            output_excel_Working_Order[line]['매수'] = output_excel_Working_Order[line]['모델명'].map(purpose_production)
            output_excel_Working_Order[line]['목표수량'] = output_excel_Working_Order[line]['모델명'].map(production_dict)
            
            output_excel_Working_Order[line]['시작시간'] = pd.to_datetime(output_excel_Working_Order[line]['시작시간'], unit='s', origin=f'{self.senario[0]}-{self.senario[1]} {self.senario[2]}:00')
            output_excel_Working_Order[line]['종료시간'] = pd.to_datetime(output_excel_Working_Order[line]['종료시간'], unit='s', origin=f'{self.senario[0]}-{self.senario[1]} {self.senario[2]}:00')
            output_excel_Working_Order[line]['생산예상시간'] = output_excel_Working_Order[line]['시작시간'].dt.strftime('%H:%M:%S') + '-' + output_excel_Working_Order[line]['종료시간'].dt.strftime('%H:%M:%S')
        
        
        for line in output_excel_Working_Order.keys():
            output_excel_Working_Order[line] = output_excel_Working_Order[line].rename(columns={'시작시간': '계획 시작시간', '종료시간': '계획 종료시간'})
        
        columns = ['모델명', '고객사', 'P/N', '계획 시작시간', '실제 시작시간', '계획 종료시간', 
           '실제 종료시간', '생산예상시간', 'T/T', 'UPH', '생산시간(분)', '실제 생산시간(분)', 
           'M/C(분)', '연배', '매수', '목표수량', '실적', '비고']
        
        new_working_order = {line : i.reindex(columns=columns) for line,i in output_excel_Working_Order.items()}
        
        for line in new_working_order.keys():
            new_working_order[line].insert(0, line, range(1, len(new_working_order[line]) + 1))
        
        #self.output_excel_rest_Working_Order = new_working_order
        
        over_working_order = copy.deepcopy(new_working_order)
        for line in over_working_order.keys():
            if self.line_use[line] == '사용':
                target_time = pd.to_datetime(self.work_time[line]+4*60*60, unit='s', origin=f'{self.senario[0]}-{self.senario[1]} {self.senario[2]}:00')
                over_working_order[line] = over_working_order[line][over_working_order[line]['계획 종료시간'] > target_time]
                # 타겟 시간 전에 시작하고, 타겟 시간 이후에 종료되는 작업 찾기
                df_target = over_working_order[line][over_working_order[line]['계획 시작시간'] <= target_time]
                if len(df_target)==1:
                    selected_task = df_target.iloc[0]  # 첫 번째(유일한) 행 선택
                    total_duration = selected_task['계획 종료시간'] - selected_task['계획 시작시간']
                    #완료 작업시간
                    elapsed_time = target_time - selected_task['계획 시작시간']
                    #미완료 작업시간
                    remaining_time = selected_task['계획 종료시간'] - target_time
                    #작업 완료비율
                    elapsed_ratio = elapsed_time / total_duration
                    #작업 미완료비율
                    remaining_ratio = remaining_time / total_duration
                    
                    df_target['계획 시작시간'] = target_time
                    df_target['생산예상시간'] = df_target['계획 시작시간'].dt.strftime('%H:%M:%S') + '-' + df_target['계획 종료시간'].dt.strftime('%H:%M:%S')
                    df_target['생산시간(분)'] = ((df_target['계획 종료시간'] - df_target['계획 시작시간']).dt.total_seconds() / 60).astype(int)
                    
                df_target.reindex(columns=columns)
                
                if len(df_target)==1:
                    over_working_order[line].iloc[0] = df_target.iloc[0]
                
                
                over_working_order[line]['매수'] = round(
                                                            (pd.to_datetime(over_working_order[line]['계획 종료시간']) - 
                                                             pd.to_datetime(over_working_order[line]['계획 시작시간'])).dt.total_seconds() /
                                                        over_working_order[line]['T/T'] * self.efficiency.loc[line,'라인효율']) 
                over_working_order[line]['목표수량'] = round(over_working_order[line]['매수'] * over_working_order[line]['연배'])
        
        # 추가할 문자열
        add_str = '_초과작업'
        
        # over_working_order의 키-값 쌍을 new_working_order에 추가
        
        for key, value in over_working_order.items():
            new_key = str(key)+add_str  # 새로운 키 이름 생성
            if self.line_use[key] == '사용':
                new_working_order[new_key] = value  # 새 키와 기존 값을 new_working_order에 추가
            else:
                new_working_order[new_key] = pd.DataFrame(columns=value.columns)
        
        self.output_excel_rest_Working_Order = new_working_order
        
        
    
    
    def gantt_chart(self,result_processing):
        Working_order = copy.deepcopy(result_processing)
        product_df = self.product_df
        
        for line in Working_order.keys():
            Working_order[line].columns = ['Task', 'Start', 'Finish']
            
            
            Working_order[line]['Start'] = pd.to_datetime(Working_order[line]['Start'], unit='s', origin=f'{self.senario[0]}-{self.senario[1]} {self.senario[2]}:00')
            Working_order[line]['Finish'] = pd.to_datetime(Working_order[line]['Finish'], unit='s', origin=f'{self.senario[0]}-{self.senario[1]} {self.senario[2]}:00')
            
            Working_order[line]['Resource'] = f'수식 모델 {line}'
        
        df_combined = pd.concat(Working_order)
        new_column_order = ['Resource', 'Task', 'Start', 'Finish']
        
        
        
        CC = {'0': '#4C5C66','1': '#A77DBE', '2': '#8446A1', '3': '#A568B7', '4': '#C87EDB', '5': '#281184', '6': '#EAE255', '7': '#21C8C7', '8': '#FF7DD8', '9': '#0128B0', '10': '#E75CE1', '11': '#68D57E', '12': '#48A232', '13': '#A064A9', '14': '#866B46', '15': '#553441', '16': '#71880A', '17': '#83EA14', '18': '#F71AF8', '19': '#7DC03B', '20': '#6A10BE', '21': '#B8A796', '22': '#DD37E7', '23': '#87BD81', '24': '#DD443C', '25': '#4250F5', '26': '#7F0DC5', '27': '#0C4C95', '28': '#898D07', '29': '#16E13E', '30': '#B3C4F3', '31': '#5548AC', '32': '#E0E183', '33': '#6EA927', '34': '#059638', '35': '#A44277', '36': '#93D0BC', '37': '#DF89A8', '38': '#911BB1', '39': '#40D8B3', '40': '#CEEA69', '41': '#53DD3F', '42': '#0D2ABF', '43': '#895979', '44': '#4A410B', '45': '#3F43EA', '46': '#1B0DEE', '47': '#981372', '48': '#723159', '49': '#A4A229', '50': '#EC28A2', '51': '#8EC732', '52': '#663A2F', '53': '#52A365', '54': '#28AF4E', '55': '#3220BD', '56': '#8D03F1', '57': '#C5B02B', '58': '#7C9E2C', '59': '#FB25E2', '60': '#DBFC03', '61': '#C36B0C', '62': '#188E3A', '63': '#526F30', '64': '#4F2D68', '65': '#1A26BB', '66': '#2B3E42', '67': '#D91825', '68': '#6F84A7', '69': '#B46A25', '70': '#8CF13E', '71': '#8C5B77', '72': '#249232', '73': '#951E8C', '74': '#177A9E', '75': '#5F59BA', '76': '#F207BB', '77': '#45450E', '78': '#5397F5', '79': '#42BDE8', '80': '#E2DAD8', '81': '#6D862A', '82': '#33FB15', '83': '#E24178', '84': '#17A35F', '85': '#FC2730', '86': '#9FDDF3', '87': '#E2B811', '88': '#DEF08D', '89': '#11E161', '90': '#A4DA66', '91': '#80EE04', '92': '#CA84F9', '93': '#F6B586', '94': '#338560', '95': '#F9097C', '96': '#83D705', '97': '#8BBC87', '98': '#B9CD8D', '99': '#E0B0EF', '100': '#7A5B52', '101': '#62D9A2', '102': '#9B1D96', '103': '#AAF2DB', '104': '#4F2A02', '105': '#2BDE4B', '106': '#5BA81B', '107': '#002C0D', '108': '#5E4DEF', '109': '#02F0F1', '110': '#4B4BB2', '111': '#252DA9', '112': '#8690A4', '113': '#13C0D4', '114': '#BF06A9', '115': '#C8A91B', '116': '#53A402', '117': '#827044', '118': '#64D9A6', '119': '#8A5128', '120': '#19A88A', '121': '#D1EF83', '122': '#CF1492', '123': '#489598', '124': '#972657', '125': '#12EB73', '126': '#AA359E', '127': '#205AA6', '128': '#308D7B', '129': '#781BF9', '130': '#440BA7', '131': '#6BA04F', '132': '#A71AA8', '133': '#76F49A', '134': '#DD43A3', '135': '#8156D4', '136': '#524F2C', '137': '#440109', '138': '#BFA286', '139': '#C7CD31', '140': '#1EA9DC', '141': '#8E89B5', '142': '#61411A', '143': '#F6E579', '144': '#6202CA', '145': '#C3ADB1', '146': '#D1795D', '147': '#EEDDB5', '148': '#78AF1A', '149': '#EF629F', '150': '#F7D951', '151': '#F58982', '152': '#16EB74', '153': '#BF608A', '154': '#4F01DE', '155': '#DAB7B7', '156': '#C4DD75', '157': '#644195', '158': '#37D91D', '159': '#6CFA25', '160': '#DFDC12', '161': '#A03090', '162': '#F8BE6B', '163': '#8F9537', '164': '#A6B3C3', '165': '#3F8BFD', '166': '#DCC742', '167': '#10DC9B', '168': '#5BE8E0', '169': '#78A9E7', '170': '#B08DA0', '171': '#8A2DDA', '172': '#C05CE0', '173': '#38F6ED', '174': '#FA70BF', '175': '#1E8AA3', '176': '#A9201D', '177': '#7CF0FF', '178': '#27600C', '179': '#0F351A', '180': '#501153', '181': '#D3D011', '182': '#C2A993', '183': '#C57104', '184': '#A8B1F5', '185': '#CE43FF', '186': '#F0E47F', '187': '#38EABD', '188': '#209852', '189': '#5E2DAA', '190': '#173357', '191': '#6A5BEE', '192': '#EFFAE0', '193': '#4BA162', '194': '#51F0CD', '195': '#4674EF', '196': '#52F7E0', '197': '#80397B', '198': '#D5F240', '199': '#94095A', '200': '#0525FF', '201': '#D42A63', '202': '#22A7BB', '203': '#7CE731', '204': '#C18F41', '205': '#1D0A53', '206': '#226F17', '207': '#90A0E8', '208': '#73ECB0', '209': '#D2FC7C', '210': '#087CC6', '211': '#457DB0', '212': '#65CE5C', '213': '#C6A3C6', '214': '#3E712A', '215': '#A68E4D', '216': '#026A1E', '217': '#D37A27', '218': '#CC20BB', '219': '#8A0A60', '220': '#0DF00B', '221': '#DAFC7D', '222': '#3CAE9E', '223': '#1E523F', '224': '#6790A0', '225': '#FA6E5C', '226': '#037A05', '227': '#C8F5CA', '228': '#817DBA', '229': '#0A8374', '230': '#FAC0F7', '231': '#37C1B7', '232': '#5C7A98', '233': '#F4EE2A', '234': '#172B4E', '235': '#E7C2EF', '236': '#0438E0', '237': '#C5BD47', '238': '#1FA7ED', '239': '#6FA7D8', '240': '#70C091', '241': '#E8A7C5', '242': '#1F40BD', '243': '#328319', '244': '#0B1AB9', '245': '#7AC591', '246': '#52D96C', '247': '#679DF2', '248': '#2E929C', '249': '#BC1AF3', '250': '#005749', '251': '#95F795', '252': '#29046F', '253': '#80C9C4', '254': '#D101A9', '255': '#82FEDA', '256': '#B0CEFC', '257': '#DEB599', '258': '#F177D0', '259': '#A080BC', '260': '#DD848D', '261': '#F70AD7', '262': '#3B002B', '263': '#445CCE', '264': '#6B3C1E', '265': '#4FC0CC', '266': '#B76ACE', '267': '#88A89B', '268': '#952DFD', '269': '#2F7EE1', '270': '#FBE850', '271': '#0EC02E', '272': '#D2C9C7', '273': '#E8DC9A', '274': '#B0AFAD', '275': '#DE1C5F', '276': '#44F3B7', '277': '#B0107B', '278': '#17EEED', '279': '#3DEA6D', '280': '#A8B338', '281': '#6F9233', '282': '#73AADE', '283': '#F85A5E', '284': '#8B2BE2', '285': '#E3B909', '286': '#2A5D4C', '287': '#78F3E4', '288': '#3A4755', '289': '#90EED3', '290': '#CC82D7', '291': '#CE5BBD', '292': '#44895B', '293': '#DC7AEE', '294': '#5181F5', '295': '#E7E715', '296': '#657AA0', '297': '#BDD839', '298': '#73330F', '299': '#E15577', '300': '#D7E2C4', '301': '#70F1D9', '302': '#6A4C8A', '303': '#93E4BD', '304': '#A1AE08', '305': '#13021D', '306': '#9FD42F', '307': '#887BCE', '308': '#2A55E7', '309': '#57592E', '310': '#2B31DD', '311': '#9F5A15', '312': '#81A6F1', '313': '#28974C', '314': '#9CD714', '315': '#F27BFB', '316': '#65956E', '317': '#6C56D2', '318': '#BBF9F7', '319': '#FFFAE1', '320': '#DF51C7', '321': '#AC3ACB', '322': '#3AFE14', '323': '#2ABDCE', '324': '#11020C', '325': '#34850B', '326': '#B054A8', '327': '#BA0140', '328': '#3A2BA4', '329': '#70EE65', '330': '#09C1E1', '331': '#5C593C', '332': '#E76330', '333': '#48847A', '334': '#72B303', '335': '#DB809E', '336': '#2ACDED', '337': '#D50C57', '338': '#67AE03', '339': '#FADBA0', '340': '#F7A4D3', '341': '#8CBABA', '342': '#6277F4', '343': '#6FC2B0', '344': '#0B0F33', '345': '#B3165B', '346': '#1AF76F', '347': '#25E358', '348': '#DE84D7', '349': '#DA0ADC', '350': '#44011D', '351': '#BC5E37', '352': '#0B748F', '353': '#98209F', '354': '#7CFDB2', '355': '#358263', '356': '#8F74CC', '357': '#25ED08', '358': '#68CD81', '359': '#2BA318', '360': '#67D677', '361': '#2270C5', '362': '#56B150', '363': '#642F6E', '364': '#837FFC', '365': '#405F0F', '366': '#546490', '367': '#17519F', '368': '#7EDC04', '369': '#1F99B1', '370': '#8ED986', '371': '#B5D609', '372': '#16A06C', '373': '#433A8C', '374': '#0A1AF6', '375': '#72CF27', '376': '#C5A93D', '377': '#9774EB', '378': '#EF6796', '379': '#258B18', '380': '#D98577', '381': '#6C9188', '382': '#D6F5C8', '383': '#D3C5D0', '384': '#EA4003', '385': '#995E58', '386': '#D59B82', '387': '#74DE83', '388': '#C68922', '389': '#DAEEA9', '390': '#0CB538', '391': '#83256D', '392': '#C65F1B', '393': '#AF07C0', '394': '#2211EF', '395': '#A97ABB', '396': '#9EC1FA', '397': '#44C209', '398': '#F9E45A', '399': '#35C409', '400': '#9101BC', '401': '#C63DAB', '402': '#D1CA0F', '403': '#708102', '404': '#5F9130', '405': '#60F8F8', '406': '#F94B9E', '407': '#F41316', '408': '#681168', '409': '#F21A08', '410': '#1645F9', '411': '#F884F8', '412': '#DD7AED', '413': '#8BDDA3', '414': '#3B8CCE', '415': '#784690', '416': '#C4F71D', '417': '#EC8203', '418': '#C33C2D', '419': '#287E64', '420': '#320C2B', '421': '#446C55', '422': '#389C2F', '423': '#AE9D64', '424': '#974CE5', '425': '#71EE86', '426': '#ED7F06', '427': '#17E5E0', '428': '#332E5F', '429': '#D18925', '430': '#6A4150', '431': '#94FEAC', '432': '#948A08', '433': '#CA2972', '434': '#B6B08F', '435': '#1B3336', '436': '#B229C4', '437': '#4195F9', '438': '#B2783F', '439': '#25C2B8', '440': '#430986', '441': '#1B1F03', '442': '#475493', '443': '#EC8F53', '444': '#079F45', '445': '#99648D', '446': '#585F0B', '447': '#B2CA35', '448': '#82A392', '449': '#CEFF9F', '450': '#A356FE', '451': '#CC3DC1', '452': '#872CF8', '453': '#230220', '454': '#01212B', '455': '#97F22F', '456': '#14ED87', '457': '#FD0263', '458': '#5DA544', '459': '#9F22FF', '460': '#7787BE', '461': '#E4D5C8', '462': '#8131B0', '463': '#3A5FC3', '464': '#69D978', '465': '#BFD81E', '466': '#59A56E', '467': '#453F57', '468': '#60D229', '469': '#D9AB85', '470': '#142C45', '471': '#E10AEE', '472': '#5BDC6E', '473': '#401FBD', '474': '#5E8195', '475': '#2B4E85', '476': '#22B27F', '477': '#334DCD', '478': '#66554D', '479': '#72515F', '480': '#5AC067', '481': '#4D0E8C', '482': '#692432', '483': '#4F37F3', '484': '#ADAF5C', '485': '#812BE0', '486': '#4E9600', '487': '#81CF33', '488': '#5F1B98', '489': '#159196', '490': '#F1C09D', '491': '#CBAC19', '492': '#758455', '493': '#123214', '494': '#8F2422', '495': '#ADE162', '496': '#08C096', '497': '#AB8973', '498': '#A9340F', '499': '#2D4C25', '500': '#2CB157', '501': '#D2A421', '502': '#DCAE1C', '503': '#32A34C', '504': '#CF29C9', '505': '#3AA2A5', '506': '#9AB805', '507': '#48DC38', '508': '#D21DA8', '509': '#DAD748', '510': '#0D3E73', '511': '#E809D7', '512': '#BDB2EB', '513': '#733BC3', '514': '#D5DA49', '515': '#20CCBF', '516': '#FD4E13', '517': '#C151B9', '518': '#88C8AB', '519': '#12EAC9', '520': '#31C496', '521': '#4AF3C1', '522': '#EAE2A5', '523': '#DDC5C6', '524': '#3C8DF9', '525': '#394195', '526': '#5317D1', '527': '#BD4089', '528': '#78C354', '529': '#9B134A', '530': '#56CE22', '531': '#56ADBF', '532': '#C85C28', '533': '#BDCC97', '534': '#783B91', '535': '#168F9C', '536': '#5D134C', '537': '#4F968F', '538': '#907337', '539': '#DF9EF8', '540': '#F5C3E7', '541': '#58B91B', '542': '#DF5AAE', '543': '#1507EA', '544': '#1549D7', '545': '#950CF7', '546': '#E3738E', '547': '#1F330A', '548': '#453A53', '549': '#66AB61', '550': '#13B1B4', '551': '#C72E2B', '552': '#B7CF11', '553': '#E9F15F', '554': '#4864E5', '555': '#26DEF5', '556': '#04B609', '557': '#80BA3B', '558': '#23967F', '559': '#441986', '560': '#9A33FA', '561': '#79A7E3', '562': '#01D8D7', '563': '#105E82', '564': '#C8AD99', '565': '#E4846D', '566': '#20EADB', '567': '#31E753', '568': '#15801D', '569': '#27C5B3', '570': '#608D30', '571': '#C11481', '572': '#47DA78', '573': '#5E8853', '574': '#B04FF7', '575': '#FD60AE', '576': '#15D859', '577': '#98C088', '578': '#3D09BF', '579': '#5A72B7', '580': '#3F035F', '581': '#7EC968', '582': '#A0BA98', '583': '#3D8E5A', '584': '#1C0814', '585': '#21F121', '586': '#9FF3F8', '587': '#B4AB41', '588': '#4B2550', '589': '#B606D3', '590': '#AB5264', '591': '#3E22CB', '592': '#AF04D3', '593': '#A062A2', '594': '#74C699', '595': '#392619', '596': '#BCD51F', '597': '#B1E10A', '598': '#F6ABAE', '599': '#9E8D3D', '600': '#E8D934', '601': '#F02E26', '602': '#2A982A', '603': '#40CC55', '604': '#E79D9C', '605': '#827753', '606': '#38CF4B', '607': '#BC30D2', '608': '#585337', '609': '#1E474B', '610': '#72B624', '611': '#6961AF', '612': '#EA24AA', '613': '#6B90F4', '614': '#32E78F', '615': '#E980F8', '616': '#2B2309', '617': '#2AE8C2', '618': '#23715A', '619': '#32132A', '620': '#F326AB', '621': '#B4D025', '622': '#9C5229', '623': '#382489', '624': '#936285', '625': '#D60BAA', '626': '#E387C3', '627': '#3A2905', '628': '#62FA51', '629': '#6389DF', '630': '#4E1A10', '631': '#036CE0', '632': '#82CE5F', '633': '#CD11AB', '634': '#DE4A8C', '635': '#8CF797', '636': '#E2DFC7', '637': '#875202', '638': '#27C388', '639': '#12B0C6', '640': '#2D6348', '641': '#C950D1', '642': '#65ADB2', '643': '#957718', '644': '#5CF7B4', '645': '#692739', '646': '#3BDC59', '647': '#C08C62', '648': '#F3202E', '649': '#B82714', '650': '#9EA6EF', '651': '#38DC2A', '652': '#122957', '653': '#8BD834', '654': '#63A4E4', '655': '#800099', '656': '#1634F6', '657': '#88AD82', '658': '#10FF89', '659': '#A1663D', '660': '#22AADE', '661': '#9DDD93', '662': '#EF307C', '663': '#5E2963', '664': '#4C34D6', '665': '#0F3252', '666': '#3BC859', '667': '#1B855E', '668': '#DE13AB', '669': '#B390C6', '670': '#19544D', '671': '#25AD0D', '672': '#C9EAF2', '673': '#7BE6BC', '674': '#18A98B', '675': '#410BAF', '676': '#3C9117', '677': '#112367', '678': '#FAAE66', '679': '#FB165D', '680': '#DC5F30', '681': '#A5F276', '682': '#513984', '683': '#7D02DD', '684': '#F7F811', '685': '#AA08FE', '686': '#6E3907', '687': '#CAC529', '688': '#771296', '689': '#12C2C2', '690': '#9EE24E', '691': '#36D515', '692': '#7666B3', '693': '#9A2411', '694': '#7AACE1', '695': '#F6497C', '696': '#1FAF14', '697': '#16CBB9', '698': '#D9292C', '699': '#4E8692', '700': '#319B91', '701': '#CD7EBA', '702': '#1B4B76', '703': '#B84CAA', '704': '#88F01A', '705': '#1C3A0F', '706': '#37E53E', '707': '#BBA143', '708': '#588849', '709': '#0F6A43', '710': '#E21227', '711': '#3465D3', '712': '#518191', '713': '#640A6A', '714': '#276B4E', '715': '#EC9ADF', '716': '#20CD76', '717': '#9A3FDC', '718': '#90D2F6', '719': '#0C7B29', '720': '#BCF1D2', '721': '#CC7118', '722': '#4FBDE7', '723': '#CD66B2', '724': '#83A33A', '725': '#3DF6E4', '726': '#790BD6', '727': '#4FAE5E', '728': '#F1DE31', '729': '#EB8D2C', '730': '#BF9DD2', '731': '#6C3318', '732': '#777DE4', '733': '#FF46A3', '734': '#686CA3', '735': '#C54DA5', '736': '#CC165B', '737': '#511CA9', '738': '#1073B8', '739': '#055D95', '740': '#298B12', '741': '#3ABCD1', '742': '#BC44E7', '743': '#A71FA8', '744': '#DCDC22', '745': '#AE0F29', '746': '#D65BEC', '747': '#419D67', '748': '#5D731F', '749': '#0055CB', '750': '#BAFC1D', '751': '#477603', '752': '#645967', '753': '#E8EBE4', '754': '#CB55FE', '755': '#A179E8', '756': '#660372', '757': '#894560', '758': '#71B41B', '759': '#9DB7F1', '760': '#7A3130', '761': '#F7C8CA', '762': '#8028FE', '763': '#B5F20F', '764': '#8C4881', '765': '#9E2E15', '766': '#18CBFB', '767': '#B6AAA7', '768': '#905D4B', '769': '#96B2DD', '770': '#0B5315', '771': '#AF2BE6', '772': '#094194', '773': '#5B9018', '774': '#F6CB57', '775': '#1C5C46', '776': '#A459A5', '777': '#8671CC', '778': '#A610A0', '779': '#C551B9', '780': '#D262A5', '781': '#8AA8A2', '782': '#A4302E', '783': '#481FBC', '784': '#194BC3', '785': '#207653', '786': '#B8238B', '787': '#6A6065', '788': '#6B50BE', '789': '#9CD368', '790': '#F26BAE', '791': '#CAE736', '792': '#9FA519', '793': '#EB48B1', '794': '#847579', '795': '#7E4E11', '796': '#9A38C7', '797': '#2C6569', '798': '#62B077', '799': '#13A97D', '800': '#7A62D1', '801': '#77105E', '802': '#1CAEC3', '803': '#BB6888', '804': '#D1FC2A', '805': '#56F25D', '806': '#2265FE', '807': '#CA9DA8', '808': '#C68024', '809': '#024615', '810': '#C665C5', '811': '#4B9027', '812': '#D503B4', '813': '#A073B6', '814': '#8AAFA5', '815': '#5D545D', '816': '#596FDF', '817': '#792D8F', '818': '#9BA552', '819': '#3ECF22', '820': '#84698F', '821': '#2744DC', '822': '#B0D138', '823': '#D5279A', '824': '#19914A', '825': '#0601D5', '826': '#694BC8', '827': '#96023E', '828': '#1686F2', '829': '#979F82', '830': '#28F6A0', '831': '#E3228B', '832': '#BFA02B', '833': '#81B70D', '834': '#87AFF9', '835': '#4D8C92', '836': '#CBCCA4', '837': '#9CB50D', '838': '#D17317', '839': '#4A77B9', '840': '#39C00B', '841': '#846F54', '842': '#260D7E', '843': '#68F3E6', '844': '#1C868D', '845': '#404699', '846': '#572AC7', '847': '#335D5A', '848': '#DF1EC2', '849': '#F38C4F', '850': '#655D62', '851': '#38A682', '852': '#C36E97', '853': '#5F30AF', '854': '#0D83F1', '855': '#29E270', '856': '#AC96D6', '857': '#0BCC05', '858': '#BF365E', '859': '#DE67F5', '860': '#BC4833', '861': '#66C8B0', '862': '#606B1F', '863': '#AD30E6', '864': '#41F7E9', '865': '#678D45', '866': '#DA8338', '867': '#C7FB90', '868': '#5D8296', '869': '#35285D', '870': '#1AB1BD', '871': '#D19134', '872': '#0945F6', '873': '#200C1D', '874': '#E15A19', '875': '#6CB431', '876': '#E56F6B', '877': '#DFB1E4', '878': '#2A4129', '879': '#C164EF', '880': '#787735', '881': '#A475AE', '882': '#B8CC13', '883': '#1712BE', '884': '#728611', '885': '#1405D2', '886': '#4B4139', '887': '#7FC443', '888': '#A1FA03', '889': '#3D2A2C', '890': '#5FB368', '891': '#92E4B0', '892': '#7FB664', '893': '#B6302D', '894': '#E1FCD6', '895': '#4A7F24', '896': '#0721FF', '897': '#F9FADD', '898': '#57F2EB', '899': '#77213D', '900': '#F0D037', '901': '#E2DF96', '902': '#EA1AC6', '903': '#11E37F', '904': '#747D40', '905': '#F14A43', '906': '#6308B8', '907': '#7F85A0', '908': '#FBA414', '909': '#1E8C59', '910': '#77223F', '911': '#74BEC8', '912': '#521DB0', '913': '#DD5A9A', '914': '#7D98F1', '915': '#B6B21E', '916': '#4AE984', '917': '#1E7731', '918': '#8C6DC7', '919': '#985752', '920': '#6031BF', '921': '#1FEB3A', '922': '#D115EB', '923': '#FAA2B9', '924': '#E4211E', '925': '#B8E346', '926': '#87BB0D', '927': '#7006A9', '928': '#AB3F72', '929': '#7BD745', '930': '#FC084F', '931': '#B48331', '932': '#746709', '933': '#E53E39', '934': '#B94EBC', '935': '#C6B0C1', '936': '#ACF173', '937': '#FDD019', '938': '#EE283E', '939': '#5EA307', '940': '#F68EFE', '941': '#D4150B', '942': '#A2BAC0', '943': '#8B0010', '944': '#D9602A', '945': '#33AA2D', '946': '#7AD4EE', '947': '#DD9C34', '948': '#8AC20A', '949': '#E1C566', '950': '#D9A646', '951': '#266A79', '952': '#57B174', '953': '#ACAF45', '954': '#1A0200', '955': '#13DC15', '956': '#A5E4F1', '957': '#A966EF', '958': '#EBBB4A', '959': '#EB381D', '960': '#367CF7', '961': '#D83780', '962': '#E7DE53', '963': '#8EB42B', '964': '#E17727', '965': '#A80EE2', '966': '#99097E', '967': '#87136F', '968': '#478F15', '969': '#A5B18D', '970': '#3E9020', '971': '#C965A5', '972': '#FA4B59', '973': '#A21D64', '974': '#1A607D', '975': '#D684C7', '976': '#98B99B', '977': '#CAC8E3', '978': '#C44742', '979': '#01369B', '980': '#20A427', '981': '#B00DC1', '982': '#32A8B0', '983': '#893C96', '984': '#CA0228', '985': '#00319D', '986': '#678E15', '987': '#F79D9B', '988': '#851B99', '989': '#ED2C8E', '990': '#940AC4', '991': '#65F2A3', '992': '#DC737D', '993': '#AFEAFB', '994': '#90A0CF', '995': '#32B94F', '996': '#DB96AE', '997': '#129596', '998': '#2CA404', '999': '#C765B1', '1000': '#65B55A', '1001': '#6FAB11', '1002': '#F68A51', '1003': '#39680B', '1004': '#963FE0', '1005': '#21717A', '1006': '#2C92D1', '1007': '#66324C', '1008': '#C3AC55', '1009': '#5B9C3F', '1010': '#4B0CC4', '1011': '#28E7BD', '1012': '#F341E4', '1013': '#B437B8', '1014': '#0D1109', '1015': '#BE026F', '1016': '#D9221D', '1017': '#C73C5C', '1018': '#FB6C3A', '1019': '#9B1F43', '1020': '#522FB8', '1021': '#418F76', '1022': '#610322', '1023': '#67D059', '1024': '#AEAEAA', '1025': '#182658', '1026': '#066C7F', '1027': '#ED8A58', '1028': '#1F7F94', '1029': '#58B32F', '1030': '#9994D6', '1031': '#FDF3A8', '1032': '#49F1D8', '1033': '#366081', '1034': '#2A4851', '1035': '#D664CD', '1036': '#2494F8', '1037': '#BEC30B', '1038': '#54C0FF', '1039': '#931B66', '1040': '#1F943E', '1041': '#B82513', '1042': '#2ECE08', '1043': '#977176', '1044': '#04FAB5', '1045': '#958F14', '1046': '#D23DE0', '1047': '#F450D3', '1048': '#847903', '1049': '#B8CEB5', '1050': '#25ECF3', '1051': '#470E7D', '1052': '#98FB92', '1053': '#2E13E3', '1054': '#1B4E55', '1055': '#EFF9AB', '1056': '#270850', '1057': '#4CC277', '1058': '#EDA772', '1059': '#2AA1BB', '1060': '#4A5C86', '1061': '#C31FEF', '1062': '#7F69CF', '1063': '#8A7675', '1064': '#A02DD2', '1065': '#CFB028', '1066': '#BAA286', '1067': '#622F64', '1068': '#D8B5CF', '1069': '#49834B', '1070': '#0BD390', '1071': '#590FDA', '1072': '#7D9081', '1073': '#926557', '1074': '#D3ABB3', '1075': '#CFFB24', '1076': '#8A5B0E', '1077': '#1118AC', '1078': '#37E935', '1079': '#52F489', '1080': '#4EE534', '1081': '#CF5481', '1082': '#4EB84A', '1083': '#04C613', '1084': '#D486E8', '1085': '#5D8EED', '1086': '#164FE2', '1087': '#3EDE2E', '1088': '#965E63', '1089': '#7B2BC7', '1090': '#54FE31', '1091': '#F16ADF', '1092': '#A724E5', '1093': '#BA84F0', '1094': '#36530C', '1095': '#DC909D', '1096': '#338C2F', '1097': '#3D116B', '1098': '#A3225D', '1099': '#5C24FD', '1100': '#29C75D', '1101': '#88F24A', '1102': '#A4FE35', '1103': '#C954F6', '1104': '#545376', '1105': '#679309', '1106': '#EC4C92', '1107': '#8F93DA', '1108': '#4C1C43', '1109': '#3A78E4', '1110': '#A36FE3', '1111': '#FA25F9', '1112': '#CA2B49', '1113': '#B074E8', '1114': '#B79A5A', '1115': '#130D71', '1116': '#F44014', '1117': '#32FACE', '1118': '#58F8BC', '1119': '#AFA039', '1120': '#048B6D', '1121': '#CDB939', '1122': '#3016E4', '1123': '#BE671B', '1124': '#086DA2', '1125': '#CCE16A', '1126': '#578521', '1127': '#45E74D', '1128': '#9CA86C', '1129': '#1818E4', '1130': '#9CCAEC', '1131': '#1D4AB0', '1132': '#B639B7', '1133': '#8CCC40', '1134': '#E884A8', '1135': '#6FE33E', '1136': '#491229', '1137': '#17ED32', '1138': '#3962C6', '1139': '#419A03', '1140': '#027610', '1141': '#BD235D', '1142': '#47FE89', '1143': '#F0ACD2', '1144': '#313891', '1145': '#AC01DA', '1146': '#10AEC7', '1147': '#5C3192', '1148': '#58E193', '1149': '#8EFC10', '1150': '#CA3893', '1151': '#35BEC4', '1152': '#23D192', '1153': '#836859', '1154': '#2E288C', '1155': '#7AFD76', '1156': '#2518F3', '1157': '#DF5113', '1158': '#791F92', '1159': '#D5D9D2', '1160': '#5A837E', '1161': '#C159DA', '1162': '#E9321C', '1163': '#0E4B47', '1164': '#727518', '1165': '#016680', '1166': '#D72366', '1167': '#6E0609', '1168': '#9AE605', '1169': '#9E371D', '1170': '#16564D', '1171': '#773D33', '1172': '#2E8531', '1173': '#4EB0B8', '1174': '#BEE23C', '1175': '#4C5B2E', '1176': '#CE40FA', '1177': '#1EB1E5', '1178': '#6503C0', '1179': '#C07279', '1180': '#E17171', '1181': '#874401', '1182': '#5BF6AE', '1183': '#C01901', '1184': '#9C9648', '1185': '#C798F2', '1186': '#28A248', '1187': '#1A07B0', '1188': '#1207D7', '1189': '#3FE821', '1190': '#04B2FE', '1191': '#2AE177', '1192': '#1ADB64', '1193': '#2F4610', '1194': '#103A19', '1195': '#6F2E5A', '1196': '#A11C58', '1197': '#E9ED0A', '1198': '#E61D2D', '1199': '#222A78', '1200': '#BDE8A4', '1201': '#D55957', '1202': '#03B0F6', '1203': '#FD24BB', '1204': '#D38AF7', '1205': '#66AC56', '1206': '#1D1B89', '1207': '#703FC5', '1208': '#5979F0', '1209': '#93A610', '1210': '#9EF69B', '1211': '#5CEA83', '1212': '#0238DC', '1213': '#2849E8', '1214': '#7176DC', '1215': '#AB80FD', '1216': '#B95250', '1217': '#3A57C8', '1218': '#1662C5', '1219': '#1762CE', '1220': '#07E376', '1221': '#1A274F', '1222': '#5AA741', '1223': '#CFD965', '1224': '#F366E8', '1225': '#289D8F', '1226': '#81300A', '1227': '#2DF16A', '1228': '#8E0FED', '1229': '#DEC578', '1230': '#D7C64F', '1231': '#3EAC4F', '1232': '#94318E', '1233': '#ACF602', '1234': '#6BB49F', '1235': '#60D532', '1236': '#635100', '1237': '#856153', '1238': '#028DFF', '1239': '#769D97', '1240': '#ED9669', '1241': '#55A19D', '1242': '#9A9ADF', '1243': '#894EBD', '1244': '#2DA9BF', '1245': '#60926F', '1246': '#27F25B', '1247': '#6DE48F', '1248': '#238998', '1249': '#0BC084', '1250': '#94AD75', '1251': '#12C6BC', '1252': '#C29FE0', '1253': '#CC4C9C', '1254': '#94ED59', '1255': '#3144F3', '1256': '#83DFB4', '1257': '#7E2EA6', '1258': '#7A6A3D', '1259': '#A63107', '1260': '#0BC39C', '1261': '#6A6FEE', '1262': '#1ED365', '1263': '#AA9507', '1264': '#A50913', '1265': '#DB80D2', '1266': '#172112', '1267': '#EA8BB8', '1268': '#DBC685', '1269': '#0BD62E', '1270': '#31481E', '1271': '#5C6F7E', '1272': '#CFD841', '1273': '#F41496', '1274': '#4412EF', '1275': '#BEB6E7', '1276': '#278E21', '1277': '#AA0DDD', '1278': '#9543BB', '1279': '#213B29', '1280': '#9161A5', '1281': '#AB73C4', '1282': '#0255A5', '1283': '#A3AE79', '1284': '#E4967B', '1285': '#57BEB2', '1286': '#9F0150', '1287': '#18BFA7', '1288': '#AD5153', '1289': '#DDE4EF', '1290': '#EAD8D6', '1291': '#F7D806', '1292': '#13AD19', '1293': '#508CD0', '1294': '#DCCB72', '1295': '#E2FA04', '1296': '#457FF8', '1297': '#FBF90A', '1298': '#D441C5', '1299': '#83AE6E', '1300': '#BE216E', '1301': '#D41977', '1302': '#BBAEA7', '1303': '#3AD2D3', '1304': '#159344', '1305': '#6F0579', '1306': '#0D8614', '1307': '#7AC12D', '1308': '#D364EA', '1309': '#8D8F6C', '1310': '#18E95D', '1311': '#450180', '1312': '#B2688F', '1313': '#ABBE7A', '1314': '#16ACCF', '1315': '#92CC8D', '1316': '#F12085', '1317': '#AE5B91', '1318': '#EDC7C4', '1319': '#82C42E', '1320': '#C3C364', '1321': '#E2E499', '1322': '#2D1D4D', '1323': '#EB0209', '1324': '#F33461', '1325': '#AEF41D', '1326': '#F1C0CC', '1327': '#6F92A0', '1328': '#C7D79D', '1329': '#3C2566', '1330': '#F2D943', '1331': '#C59A83', '1332': '#B4E418', '1333': '#3C6CD7', '1334': '#168CA9', '1335': '#1628D3', '1336': '#30B67A', '1337': '#BC6985', '1338': '#14635C', '1339': '#C7D4F5', '1340': '#1F2286', '1341': '#97C8D0', '1342': '#DBA49C', '1343': '#AC7EF4', '1344': '#3BBEB4', '1345': '#135DAE', '1346': '#C3D266', '1347': '#F1C91D', '1348': '#6314D0', '1349': '#6CAB37', '1350': '#36042E', '1351': '#71E0A3', '1352': '#3C5DD2', '1353': '#6EEB57', '1354': '#01EFE3', '1355': '#92BCD0', '1356': '#6CE8C5', '1357': '#43B9FC', '1358': '#6D3D10', '1359': '#634592', '1360': '#9D7B46', '1361': '#CDEF21', '1362': '#C7D367', '1363': '#41CEF6', '1364': '#AB002A', '1365': '#160972', '1366': '#A77175', '1367': '#BC1912', '1368': '#441707', '1369': '#E77B7E', '1370': '#8D8ECF', '1371': '#FC3231', '1372': '#D5EC08', '1373': '#F4C608', '1374': '#85CA98', '1375': '#DFA9DF', '1376': '#8D856B', '1377': '#FFEC3B', '1378': '#90F68C', '1379': '#4DB02A', '1380': '#23BFCF', '1381': '#00339D', '1382': '#814C0B', '1383': '#D777F9', '1384': '#12709C', '1385': '#9F631D', '1386': '#5F9C5E', '1387': '#E4489B', '1388': '#12C685', '1389': '#73408C', '1390': '#33A38D', '1391': '#82866E', '1392': '#D10378', '1393': '#7E2C29', '1394': '#89A056', '1395': '#2BD619', '1396': '#222C73', '1397': '#5F35D7', '1398': '#6AB85C', '1399': '#8B6A53', '1400': '#63446F', '1401': '#1BBD76', '1402': '#2EE7AD', '1403': '#005D91', '1404': '#E9A527', '1405': '#9553D5', '1406': '#EACE62', '1407': '#7F4973', '1408': '#4EE0A1', '1409': '#4131A8', '1410': '#237C2D', '1411': '#2A6BCC', '1412': '#ADC694', '1413': '#576786', '1414': '#151BDD', '1415': '#AA4C1F', '1416': '#C17F90', '1417': '#6DF68E', '1418': '#8E4B01', '1419': '#B8D414', '1420': '#A18A99', '1421': '#65DC27', '1422': '#C857A9', '1423': '#EC2169', '1424': '#9763E7', '1425': '#076F60', '1426': '#10C593', '1427': '#E514D6', '1428': '#372F06', '1429': '#4BE085', '1430': '#78E93A', '1431': '#BA76BE', '1432': '#2621ED', '1433': '#11C21C', '1434': '#3EC9EF', '1435': '#AAF98A', '1436': '#85429C', '1437': '#ADB0FC', '1438': '#CEBB8B', '1439': '#6D87A6', '1440': '#3F1D05', '1441': '#8DCDEE', '1442': '#D8B11F', '1443': '#B40EDE', '1444': '#5DC079', '1445': '#B1F5C4', '1446': '#D4384D', '1447': '#D3AC37', '1448': '#C014A6', '1449': '#0EF9AE', '1450': '#14993A', '1451': '#AB7650', '1452': '#316476', '1453': '#765E7A', '1454': '#EE2EF4', '1455': '#52876A', '1456': '#CA0740', '1457': '#E89CE4', '1458': '#E20B60', '1459': '#155694', '1460': '#D251B2', '1461': '#6A8C07', '1462': '#FCA0C6', '1463': '#2BBDC7', '1464': '#CDDD85', '1465': '#A5EDBE', '1466': '#47BB7D', '1467': '#C01397', '1468': '#0EA8A8', '1469': '#29525E', '1470': '#FCA6CC', '1471': '#C0C10B', '1472': '#85915A', '1473': '#8268B7', '1474': '#E10A48', '1475': '#5B78C6', '1476': '#AFB9E1', '1477': '#30CB8C', '1478': '#FDD736', '1479': '#CFA624', '1480': '#C2A5C8', '1481': '#C0936E', '1482': '#D8D192', '1483': '#5A69B3', '1484': '#83AB1F', '1485': '#39C5DF', '1486': '#007531', '1487': '#3B69DD', '1488': '#785205', '1489': '#F7C4D0', '1490': '#A1BF69', '1491': '#9A1A44', '1492': '#BF90ED', '1493': '#23747F', '1494': '#A250BF', '1495': '#776559', '1496': '#15D143', '1497': '#CC48A3', '1498': '#7901A9', '1499': '#5771D0', '1500': '#354CEC', '1501': '#095EE0', '1502': '#2B1E39', '1503': '#E92AE7', '1504': '#19C2F1', '1505': '#07A544', '1506': '#8CC970', '1507': '#B90713', '1508': '#C3F053', '1509': '#AD3798', '1510': '#8A6B65', '1511': '#BCCE70', '1512': '#F7E373', '1513': '#CD9901', '1514': '#805A4E', '1515': '#9B1DA4', '1516': '#34FCBC', '1517': '#5B079F', '1518': '#CCD13F', '1519': '#31D4DB', '1520': '#196664', '1521': '#DB1A07', '1522': '#90DD88', '1523': '#900BD9', '1524': '#0FB1A8', '1525': '#35E951', '1526': '#1BC97E', '1527': '#7E80BF', '1528': '#A3F4BD', '1529': '#B0CEC9', '1530': '#4BE903', '1531': '#FE3AF9', '1532': '#1AC5BD', '1533': '#A681DD', '1534': '#9D99F4', '1535': '#13527E', '1536': '#100C41', '1537': '#BE66A3', '1538': '#A815BB', '1539': '#F6FB91', '1540': '#0A5838', '1541': '#897E3E', '1542': '#DBFC9E', '1543': '#56E7AF', '1544': '#BA6FE5', '1545': '#14F20A', '1546': '#C73BE9', '1547': '#678251', '1548': '#729967', '1549': '#FB4862', '1550': '#DD534C', '1551': '#7DF988', '1552': '#57B4AC', '1553': '#9A70B3', '1554': '#FD2FAE', '1555': '#B814A6', '1556': '#DD2440', '1557': '#49153F', '1558': '#2CA87A', '1559': '#4BFEF0', '1560': '#59B68C', '1561': '#4DECA7', '1562': '#4D5F7C', '1563': '#99AB74', '1564': '#F21E8A', '1565': '#B6425D', '1566': '#2A76D7', '1567': '#234DCB', '1568': '#386DE5', '1569': '#F1810E', '1570': '#29B182', '1571': '#6E50DC', '1572': '#B1D55D', '1573': '#61B187', '1574': '#70AE46', '1575': '#600ED8', '1576': '#6F0697', '1577': '#7BBD25', '1578': '#5147E5', '1579': '#977146', '1580': '#A3A1EC', '1581': '#CDE746', '1582': '#0B9D6B', '1583': '#D8520D', '1584': '#697571', '1585': '#A60918', '1586': '#986809', '1587': '#C711D9', '1588': '#68DBF4', '1589': '#6CB222', '1590': '#115EB8', '1591': '#3BDB97', '1592': '#497310', '1593': '#80FD4A', '1594': '#AA5826', '1595': '#F10BA2', '1596': '#5AB4A8', '1597': '#9D94A6', '1598': '#81C807', '1599': '#721E3E', '1600': '#0C57D4', '1601': '#D01ACC', '1602': '#23B622', '1603': '#8ED79B', '1604': '#DC8495', '1605': '#F1EAEA', '1606': '#11F87B', '1607': '#935CC3', '1608': '#E70660', '1609': '#297344', '1610': '#78D61C', '1611': '#3BB2FC', '1612': '#3F02F8', '1613': '#C244BE', '1614': '#1322BC', '1615': '#F0F5FB', '1616': '#A8F805', '1617': '#303E08', '1618': '#135EA3', '1619': '#4A6C8E', '1620': '#FF4BEB', '1621': '#E9111A', '1622': '#FEC903', '1623': '#4E7956', '1624': '#2A7735', '1625': '#71AED2', '1626': '#297E35', '1627': '#7B8CB7', '1628': '#8BEC40', '1629': '#FA2880', '1630': '#6BF273', '1631': '#214889', '1632': '#8225B4', '1633': '#FCB354', '1634': '#AF130B', '1635': '#2A20BF', '1636': '#3C84C0', '1637': '#DD5071', '1638': '#BFB77B', '1639': '#AB796E', '1640': '#B20D1D', '1641': '#12344A', '1642': '#FBEE93', '1643': '#7B3584', '1644': '#43FB37', '1645': '#623F84', '1646': '#B4B758', '1647': '#383724', '1648': '#4527A5', '1649': '#5268CF', '1650': '#FE8360', '1651': '#645DA8', '1652': '#4AD761', '1653': '#3E6566', '1654': '#E030A0', '1655': '#D22C19', '1656': '#753312', '1657': '#CF426E', '1658': '#AFCD90', '1659': '#2EB0FA', '1660': '#9C7EA6', '1661': '#4E9621', '1662': '#476157', '1663': '#401979', '1664': '#0EC044', '1665': '#A9427A', '1666': '#6987F5', '1667': '#5E2DD1', '1668': '#8869F0', '1669': '#053E61', '1670': '#74B9EC', '1671': '#5617AC', '1672': '#CC2277', '1673': '#E1F090', '1674': '#7A5AF5', '1675': '#88BACA', '1676': '#F1771C', '1677': '#B0ADD2', '1678': '#F15743', '1679': '#CB5D32', '1680': '#CFFBA8', '1681': '#D9D998', '1682': '#DC208C', '1683': '#3CD639', '1684': '#FCD7BF', '1685': '#341F65', '1686': '#9AC15D', '1687': '#A7A0D6', '1688': '#C52ABB', '1689': '#D06AD5', '1690': '#E4A365', '1691': '#C6FEEC', '1692': '#C587C3', '1693': '#83B1DF', '1694': '#506073', '1695': '#F39351', '1696': '#72C73B', '1697': '#44537E', '1698': '#4EDC5B', '1699': '#5A7D63', '1700': '#7EA183', '1701': '#40099D', '1702': '#1C6916', '1703': '#E38BCE', '1704': '#8AEB9C', '1705': '#C64570', '1706': '#2E3EEE', '1707': '#04755C', '1708': '#F0F7DE', '1709': '#661C8B', '1710': '#993241', '1711': '#E722AD', '1712': '#4E9EB7', '1713': '#D0ADE9', '1714': '#825541', '1715': '#3799A9', '1716': '#BFC660', '1717': '#AC017B', '1718': '#6BD588', '1719': '#47A27A', '1720': '#236163', '1721': '#555E1C', '1722': '#15B575', '1723': '#EAC2D7', '1724': '#CDAF5D', '1725': '#C2D0C3', '1726': '#F4A2AB', '1727': '#795B7A', '1728': '#578DFC', '1729': '#44AFD6', '1730': '#58F5E6', '1731': '#28A1D1', '1732': '#49DAD2', '1733': '#980123', '1734': '#119240', '1735': '#927AA1', '1736': '#A7A472', '1737': '#A0E99B', '1738': '#EC31E0', '1739': '#7CE84D', '1740': '#2511BB', '1741': '#5AF3CD', '1742': '#66990E', '1743': '#7F1CBF', '1744': '#A7EBCB', '1745': '#DF5FE6', '1746': '#D190D2', '1747': '#309332', '1748': '#865FB6', '1749': '#5F373C', '1750': '#657A3D', '1751': '#0AAF0C', '1752': '#0513E0', '1753': '#3AA442', '1754': '#EF2B4E', '1755': '#D06F41', '1756': '#337107', '1757': '#74D7EE', '1758': '#658403', '1759': '#0952D0', '1760': '#F83A0C', '1761': '#699F59', '1762': '#C71F2A', '1763': '#BA80E2', '1764': '#DB3012', '1765': '#E78DC5', '1766': '#86103F', '1767': '#821BEF', '1768': '#D79326', '1769': '#22A484', '1770': '#0DE5F8', '1771': '#1C375B', '1772': '#6F93E3', '1773': '#FB533D', '1774': '#B0982F', '1775': '#DFF624', '1776': '#977295', '1777': '#589B52', '1778': '#D83763', '1779': '#BE8106', '1780': '#8CE187', '1781': '#F7BC2A', '1782': '#1D7630', '1783': '#5A0629', '1784': '#41B319', '1785': '#DFA95C', '1786': '#2C4078', '1787': '#A12AD3', '1788': '#66229D', '1789': '#E7E525', '1790': '#A18114', '1791': '#496FED', '1792': '#E8FEC2', '1793': '#F25CD9', '1794': '#1DB7CA', '1795': '#539D99', '1796': '#FBDCB2', '1797': '#9A2D3E', '1798': '#D0B95B', '1799': '#B581F8', '1800': '#9CD9FA', '1801': '#9D4332', '1802': '#AEEC60', '1803': '#BFCDC0', '1804': '#2D44A4', '1805': '#679E7B', '1806': '#E27F37', '1807': '#17BCC4', '1808': '#67D019', '1809': '#A0A987', '1810': '#852BDC', '1811': '#4EC18B', '1812': '#2550BB', '1813': '#041B86', '1814': '#11A70E', '1815': '#DA4E56', '1816': '#B805CB', '1817': '#125E6B', '1818': '#9547B2', '1819': '#060299', '1820': '#A0F00C', '1821': '#2545AC', '1822': '#06CD5A', '1823': '#880C72', '1824': '#3F5277', '1825': '#40AB02', '1826': '#E71934', '1827': '#0828E5', '1828': '#3B8613', '1829': '#6D4380', '1830': '#DC237D', '1831': '#F7A6E7', '1832': '#7F51B1', '1833': '#64D4C3', '1834': '#5E4A2F', '1835': '#35EDAE', '1836': '#530177', '1837': '#87DD43', '1838': '#3BFE21', '1839': '#C38AFC', '1840': '#C7146D', '1841': '#AF5BF9', '1842': '#5178B8', '1843': '#1B0097', '1844': '#D877BE', '1845': '#6C615A', '1846': '#9F489B', '1847': '#B8CC14', '1848': '#FD1175', '1849': '#19FB3C', '1850': '#26D33B', '1851': '#C0930E', '1852': '#12AFED', '1853': '#3CF1F2', '1854': '#CC15AC', '1855': '#F4A0C1', '1856': '#A3C758', '1857': '#67C7CF', '1858': '#4FCF5A', '1859': '#54A6E8', '1860': '#A61254', '1861': '#98DC59', '1862': '#944976', '1863': '#EE81B5', '1864': '#500995', '1865': '#20F079', '1866': '#A9D17E', '1867': '#37B552', '1868': '#8EEB09', '1869': '#E733ED', '1870': '#FE6C1C', '1871': '#7E4C54', '1872': '#66B8F0', '1873': '#73F886', '1874': '#E6A3FA', '1875': '#75F982', '1876': '#5EC89C', '1877': '#16AA89', '1878': '#E0F2BB', '1879': '#9C8BD8', '1880': '#B03D8B', '1881': '#138AAE', '1882': '#620192', '1883': '#3944E8', '1884': '#04FB7E', '1885': '#215A09', '1886': '#AC8037', '1887': '#4E3FA5', '1888': '#0917C8', '1889': '#0590C7', '1890': '#220B6E', '1891': '#0E29D0', '1892': '#1922C3', '1893': '#1345E4', '1894': '#00E92C', '1895': '#776FCA', '1896': '#73FDFC', '1897': '#552E0E', '1898': '#6993B8', '1899': '#1FFFEA', '1900': '#591F42', '1901': '#55D8C0', '1902': '#95B305', '1903': '#14F4F4', '1904': '#D7893F', '1905': '#6EB220', '1906': '#6AD868', '1907': '#4B40ED', '1908': '#7CEA7C', '1909': '#105AC2', '1910': '#76DA13', '1911': '#741104', '1912': '#0B1734', '1913': '#3FF143', '1914': '#5A61B1', '1915': '#4662DB', '1916': '#82F1F8', '1917': '#9F1B29', '1918': '#81D271', '1919': '#D72497', '1920': '#92F18E', '1921': '#9D60D2', '1922': '#A8BEAD', '1923': '#160629', '1924': '#3FF29C', '1925': '#6CBBC7', '1926': '#86588E', '1927': '#C6AEC1', '1928': '#116BC0', '1929': '#69A6F7', '1930': '#17142D', '1931': '#19F897', '1932': '#594635', '1933': '#C5F4A0', '1934': '#AB8AAB', '1935': '#157E62', '1936': '#1DC297', '1937': '#C4EB70', '1938': '#AADF78', '1939': '#1A1424', '1940': '#C64434', '1941': '#106F22', '1942': '#BAA617', '1943': '#3A7FAF', '1944': '#A450C9', '1945': '#0D6D2B', '1946': '#76C66D', '1947': '#BDDBE8', '1948': '#C8EBEA', '1949': '#A90269', '1950': '#6F7620', '1951': '#761CDF', '1952': '#332BE5', '1953': '#AD4695', '1954': '#687CFC', '1955': '#EF5F83', '1956': '#CBA07E', '1957': '#D5989D', '1958': '#11527C', '1959': '#4AD172', '1960': '#87A34B', '1961': '#8242F9', '1962': '#446BE5', '1963': '#031934', '1964': '#49EF65', '1965': '#C207B3', '1966': '#51CA14', '1967': '#40B65D', '1968': '#37F977', '1969': '#38E595', '1970': '#81C6E1', '1971': '#5D4E6F', '1972': '#15420F', '1973': '#D3528F', '1974': '#30FBF9', '1975': '#F3C863', '1976': '#CBB3FA', '1977': '#C2DF99', '1978': '#B1AD63', '1979': '#808ADA', '1980': '#272D4B', '1981': '#59529B', '1982': '#0349B5', '1983': '#350D18', '1984': '#1F0C07', '1985': '#5841BA', '1986': '#23372B', '1987': '#FE7519', '1988': '#A3BAA8', '1989': '#668DC3', '1990': '#40B04A', '1991': '#8EE161', '1992': '#71B5DD', '1993': '#54219A', '1994': '#60CBB8', '1995': '#CF652A', '1996': '#6D91D9', '1997': '#EAC3D6', '1998': '#EE7B03', '1999': '#DF9C84'}

        
        product_names = product_df['제품명'].tolist()
        product_names.append('제품 없음')
        product_names.append('셋업')
        product_names.append('휴식')
        
        
        colors = {}
        for i, product_name in enumerate(product_names):
            color = CC[str(i % len(CC))]    
            colors[product_name] = color
        
        
        df_combined = df_combined.reindex(columns=new_column_order)
        df_combined = df_combined.rename(columns={'Task': '작업'})
        df_combined = df_combined.rename(columns={'Resource': 'Task'})
        
        fig = ff.create_gantt(df_combined, colors=colors, index_col='작업', show_colorbar=True, group_tasks=True, showgrid_x=True)
        
        fig.update_layout(
            title={
                'text': "김해공장 라인 작업 간트 차트<br><sub>비교 간트차트</sub>",  
                'y':0.9,  
                'x':0.5,
                'xanchor': 'center',
                'yanchor': 'top'},
        )
        pio.write_html(fig, file=f"Gantt_chart_by_line_{self.senario[1]}.html", auto_open=False)


if __name__ == "__main__":
    p_filepath = ['제품입력 데이터.xlsm', '목록']
    d_filepath = ['제품생산정보_폼.xlsm', '목록']
    
    scheduler = productionScheduler()
    
    scheduler.load_data(p_filepath,d_filepath)
    
    scheduler.data_structures()
    
    scheduler.setup_constraints()
    

