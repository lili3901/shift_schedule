from ortools.sat.python import cp_model
import xlrd
model=cp_model.CpModel()

# read data from excel(文件名，表序号)
def read_data_from_excel(filename):
    workbook=xlrd.open_workbook(filename)
    sheet0=workbook.sheet_by_index(0)
    lastweek_assignments=[sheet0.cell_value(r,c) for r in range(1,sheet0.nrows) for c in range(1,4)]
    sheet1=workbook.sheet_by_index(1)
    employees=[sheet1.cell_value(r,0) for r in range(1,sheet1.nrows)]
    work_days=sheet1.cell_value(1,2)
    num_days=sheet1.cell_value(1,5)
    holi_days=[sheet1.cell_value(r,7) for r in range(1,sheet1.nrows)]
    min_M910=sheet1.cell_value(1,15)
    min_night=sheet1.cell_value(1,16)
    sheet2=workbook.sheet_by_index(2)
    dayly_cover_demands=[sheet2.cell_value(r,c) for r in range(1,sheet2.nrows) for c in range(0,10)]
    sheet3=workbook.sheet_by_index(3)
    requests=[sheet3.cell_value(r,c) for r in range(1,sheet3.nrows) for c in range(1,5)]
    sheet4=workbook.sheet_by_index(4)
    fixed_assignments=[sheet4.cell_value(r,c) for r in range(1,sheet4.nrows) for c in range(1,4)]
    return lastweek_assignments, employees, work_days, num_days, holi_days, min_M910, min_night, dayly_cover_demands, requests, fixed_assignments

# 读入各项参数
filename='mvtm_zx_nohalf.xls'
lastweek_assignments, employees, work_days, num_days, holi_days, min_M910, min_night, dayly_cover_demands, requests, fixed_assignments=read_data_from_excel(filename)
num_employees=len(employees)
shifts=['休','年假','M1','M2','M3','M4','M5','M6','M7','M8','M9','M10']
num_shifts=len(shifts)

obj_bool_vars=[]
obj_bool_coeffs=[]
# 定义排班变量
work={}
for e in range(num_employees):
    for s in range(num_shifts):
        for d in range(num_days):
            work[e,s,d]=model.NewBoolVal('work%i_%i_%i' % (e,s,d))

# fixed_assignments(e,s,d)
for e,s,d in fixed_assignments:
    model.Add(work[e,s,d]==1)
# lastweek_assignments=[]
for e,s,d in lastweek_assignments:
    model.Add(work[e,s,d]==1)
# requests(e,s,d,w)
for e,s,d,w in requests:
    obj_bool_vars.append(work[e,s,d])
    obj_bool_coeffs.append(w)




# 每个员工每天只能上一个班次
for e in range(num_employees):
    for d in range(num_days):
        model.Add(sum(work[e,s,d] for s in range(num_shifts))==1)



# 最大连续工作天数不超过7天
for e in range(num_employees):
    for i in range(num_days):
        model.Add(sum(work[e,s,d] for s in range(2,num_shifts) for d in range(i-6,i))<7)
# 早班班个数约束。从外部读取

# 班次衔接约束

# 休息天数约束。从外部读取
rest_days=num_days-work_days

for e in range(num_employees):
    model.Add(sum(work[e,0,d] for d in range(num_days)  if d not in holi_days)==rest_days)

# 每日班次个数要求.从外部读取
dayly_cover_demands=[]
for s in range(2,num_shifts):
    for d in range(num_days):
        works=[work[e,s,d] for e in range(num_employees)]
        min_demand=dayly_cover_demands[d][s-2]
        model.Add(sum(works)>=min_demand)

# 规划目标
model.Maximize(
    sum(obj_bool_vars[i]*obj_bool_coeffs[i] for i in range(len(obj_bool_vars)))
)
solver=cp_model.CpSolver()
solver.paramters.max_time_in_seconds=400
solution_printer=cp_model.ObjectiveSolutionPrinter()
status=solver.SolveWithSolutionCallback(model,solution_printer)



# print
if status==cp_model.OPTIMAL or status==cp_model.FEASIBLE:
    print()
    header='      '
    for d in range(num_days):
        header += d + '  ' 
    print(header)
    for e in range(num_employees):
        schedule=''
        for d in range(num_days):
            for s in range(num_shifts):
                if solver.BooleanValue(work(e,s,d)):
                    schedule += shifts[s] + '  '
        print('worker %i:%s' % (e,schedule))
    print()
    print('Penalties:')
