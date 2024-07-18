from ortools.sat.python import cp_model
from openpyxl import load_workbook
from dateutil.parser import parse
import calendar
import os
current_dir=os.path.dirname(os.path.abspath(__file__))
model=cp_model.CpModel()

# read data from excel(文件名)
def read_data_from_excel(filename):
    workbook=load_workbook(filename,data_only=True)
    
    lastweek_assignments=[]
    for row in workbook['上月班次衔接'].iter_rows(min_row=2,min_col=2,max_col=4,values_only=True):
        lastweek_assignments.append(tuple(row))
    
    employees=[]
    for row in workbook['capital'].iter_rows(min_row=2,min_col=1,max_col=1,values_only=True):
        employees.append(row[0])
    work_days=workbook['capital']['C2'].value
    num_days=workbook['capital']['F2'].value
    
    #date=parse(date)
    #num_days=calendar.monthrange(date.year,date.month)
    holi_days=[]
    for row in workbook['capital'].iter_rows(min_row=2,max_row=6,min_col=7,max_col=7,values_only=True):
        if row[0] is not None:
            holi_days.append(row[0])
    min_M910=workbook['capital']['P2'].value
    min_night=workbook['capital']['Q2'].value

    dayly_cover_demands=[]
    for row in workbook['班次要求'].iter_rows(min_row=2,min_col=1,max_col=10,values_only=True):
        if row[0] is not None:
            dayly_cover_demands.append(tuple(row))
    
    requests=[]
    for row in workbook['班次意见'].iter_rows(min_row=2,min_col=2,max_col=5,values_only=True):
        if row[0] is not None:
            requests.append(tuple(row))
    
    fixed_assignments=[]
    for row in workbook['固定班次'].iter_rows(min_row=2,min_col=2,max_col=4,values_only=True):
        fixed_assignments.append(tuple(row))

    workbook.close()
    return lastweek_assignments, employees, work_days, num_days, holi_days, min_M910, min_night, dayly_cover_demands, requests, fixed_assignments

# 读入各项参数
filename='mvtm_zx_nohalf.xlsm'
file_path=os.path.join(current_dir,filename)
lastweek_assignments, employees, work_days, num_days, holi_days, min_M910, min_night, dayly_cover_demands, requests, fixed_assignments=read_data_from_excel(file_path)
num_employees=len(employees)
shifts=['休','年假','M1','M2','M3','M4','M5','M6','M7','M8','M9','M10']
num_shifts=len(shifts)

obj_bool_vars=[]
obj_bool_coeffs=[]
# 定义排班变量
work={}
for e in range(num_employees):
    for s in range(num_shifts):
        for d in range(-6,num_days):
            work[e,d,s]=model.NewBoolVar('work%i_%i_%i' % (e,d,s))

# fixed_assignments(e,s,d)
for e,d,s in fixed_assignments:
    model.Add(work[e,d,s]==1)
# lastweek_assignments=[]
for e,d,s in lastweek_assignments:
    model.Add(work[e,d,s]==1)
# requests(e,s,d,w)
for e,d,s,w in requests:
    obj_bool_vars.append(work[e,d,s])
    obj_bool_coeffs.append(w)




# 每个员工每天只能上一个班次
for e in range(num_employees):
    for d in range(num_days):
        model.Add(sum(work[e,d,s] for s in range(num_shifts))==1)



# 最大连续工作天数不超过7天
for e in range(num_employees):
    for i in range(num_days):
        model.Add(sum(work[e,d,s] for s in range(2,num_shifts) for d in range(i-6,i))<7)
# 早班班个数约束。从外部读取

# 班次衔接约束

# 休息天数约束。从外部读取
rest_days=num_days-work_days

for e in range(num_employees):
    model.Add(sum(work[e,d,0] for d in range(num_days)  if d not in holi_days)==rest_days)

# 每日班次个数要求.从外部读取
#dayly_cover_demands=[]
for s in range(2,num_shifts):
    for d in range(num_days):
        works=[work[e,d,s] for e in range(num_employees)]
        min_demand=dayly_cover_demands[d][s-2]
        model.Add(sum(works)>=min_demand)

# 规划目标
model.Maximize(
    sum(obj_bool_vars[i]*obj_bool_coeffs[i] for i in range(len(obj_bool_vars)))
)
solver=cp_model.CpSolver()
solver.parameters.max_time_in_seconds=400

solution_printer=cp_model.ObjectiveSolutionPrinter()
status=solver.SolveWithSolutionCallback(model,solution_printer)



# print
if status==cp_model.OPTIMAL or status==cp_model.FEASIBLE:
    print()
    header='      '
    for d in range(num_days):
        header += str(d) + '  ' 
    print(header)
    for e in range(num_employees):
        schedule=''
        for d in range(num_days):
            for s in range(num_shifts):
                if solver.BooleanValue(work[e,d,s]):
                    schedule += shifts[s] + '  '
        print('worker %i:%s' % (e,schedule))
    print()
    print('Penalties:')

# 下面尝试写一个输出多个可行解的class，与上面的输出可能冲突
class PartialSolutionPrinter(cp_model.CpSolverSolutionCallback):
    '''Print intermediate solutions.'''
    def __init__(self,shifts,employees,num_days,limit) -> None:
        cp_model.CpSolverSolutionCallback.__init__(self)
        self._shifts=shifts
        self._employees=employees
        self._num_days=num_days
        self._solution_count=0
        self._solution_limit=limit

    def on_solution_callback(self):
        self._solution_count+=1
        print(f"排班方案_{self._solution_count}")
        # 这里写输出的逻辑

        if self._solution_count>=self._solution_limit:
            print(f"停止！已找到{self._solution_limit}个方案。")
            self.stop_search()
    def solutionCount(self):
        return self._solution_count
    
# Display the first five solutions.
solution_limit=5
# 用重写的Printer替换默认的
solution_printer=PartialSolutionPrinter(shifts,employees,num_days,solution_limit)
solver.solve(model,solution_printer)

# Statistics.
print("\nStatistics")
print(f" - conflicts       : {solver.num_conflicts}")
print(f" - branches        : {solver.num_branches}")
print(f" - wall time       : {solver.wall_time}")
print(f" - solutions found : {solution_printer.solutionCount()}")
    