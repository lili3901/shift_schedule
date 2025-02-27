# windows10 python3.12 ortools9.10 不能运行，原因未知。
from ortools.sat.python import cp_model
from openpyxl import load_workbook,Workbook
from openpyxl.styles import Font, PatternFill, Alignment,Border,Side

import os
current_dir=os.path.dirname(os.path.abspath(__file__))
model=cp_model.CpModel()

# read data from excel(文件名)
def read_data_from_excel(filename):
    workbook=load_workbook(filename,data_only=True)
    duty_count=0
    rest_count=0
    print("读取……上月排班")
    lastweek_assignments=[]
    for row in workbook['上月班次衔接'].iter_rows(min_row=2,min_col=2,max_col=4,values_only=True):
        lastweek_assignments.append(tuple(row))
    print("读取……排班名单")
    employees=[]
    for row in workbook['capital'].iter_rows(min_row=2,min_col=1,max_col=1,values_only=True):
        employees.append(row[0])
    print("读取……工作日和总天数")
    work_days=workbook['capital']['C2'].value
    num_days=workbook['capital']['F2'].value
    print("读取……节日")
    holi_days=[]
    for row in workbook['capital'].iter_rows(min_row=2,max_row=6,min_col=8,max_col=8,values_only=True):
        if row[0] is not None:
            holi_days.append(row[0])
    print("读取……晚班限制")
    min_M910=workbook['capital']['P2'].value
    min_night=workbook['capital']['Q2'].value
    print("读取……班次需求")
    dayly_cover_demands=[]
    for row in workbook['班次要求'].iter_rows(min_row=2,min_col=1,max_col=10,values_only=True):
        if row[0] is not None:
            dayly_cover_demands.append(tuple(row))
    print("读取……排班意见")
    requests=[]
    for row in workbook['班次意见'].iter_rows(min_row=2,min_col=2,max_col=5,values_only=True):
        # 因为sheet中已用行数可能大于requests行数，所以要判断。班次要求同理。
        if row[0] is not None:
            requests.append(tuple(row))
            if row[2]==0:
                rest_count+=1
            else:
                duty_count+=1
    print("读取……固定班次")
    fixed_assignments=[]
    for row in workbook['固定班次'].iter_rows(min_row=2,min_col=2,max_col=4,values_only=True):
        fixed_assignments.append(tuple(row))
    print("读取……班次衔接")
    penalized_transitions=[]
    for row in workbook['班次衔接'].iter_rows(min_row=2,min_col=1,max_col=2,values_only=True):
        penalized_transitions.append(tuple(row))
    print("读取……班次")    
    shifts=[cell for row in workbook['班次衔接'].iter_rows(min_row=2,min_col=5,max_col=5,values_only=True) for cell in row if cell is not None]
    n_shifts=[]
    for row in workbook['班次衔接'].iter_rows(min_row=2,min_col=7,max_col=8,values_only=True):
        if row[0] is not None:
            n_shifts.append(row)
    workbook.close()
    return lastweek_assignments, employees, work_days, num_days, holi_days, min_M910, min_night, dayly_cover_demands, requests, fixed_assignments,rest_count,duty_count,penalized_transitions,shifts,n_shifts

# 读入各项参数
filename='mvtm_zx_nohalf.xlsm'
file_path=os.path.join(current_dir,filename)
lastweek_assignments, employees, work_days, num_days, holi_days, min_M910, min_night, dayly_cover_demands, requests, fixed_assignments,rest_count,duty_count,penalized_transitions,shifts,n_shifts=read_data_from_excel(file_path)
print("排班意见统计：")
print(f"-休息意见  :{rest_count}")
print(f"-上班意见  :{duty_count}")
num_employees=len(employees)
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
# 检查固定班次个数与班次个数要求是否冲突，很重要。
for e,d,s in fixed_assignments:
    model.Add(work[e,d,s]==1)
# 限制固定班次的个数（除‘休’外），不能多排。很重要
for e in range(num_employees):
    for d in range(num_days):
        for s in range(num_shifts-2,num_shifts):
            if (e,d,s) not in fixed_assignments:
                model.Add(work[e,d,s]==0)
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

def condition_1():
    # 最大连续工作天数不超过7天
    for e in range(num_employees):
        for i in range(num_days):
            model.Add(sum(work[e,d,s] for s in range(num_shifts) if s not in [0,num_shifts-2] for d in range(i-6,i+1))<7)

def condition_2():
    # 晚班个数约束。从外部读取
    for s, s_num in n_shifts:
        ns = s.split(",")
        nights = [int(n) for n in ns]
        for e in range(num_employees):
            model.Add(sum(work[e,d,s] for d in range(num_days) for s in nights) < (s_num + 1))
            # 这个是临时的，因为指定的VM3个数超过平均晚班数
            if s == "7,8,9,10":
                model.Add(sum(work[e,d,s] for d in range(num_days) for s in nights) >= (s_num - 1))

def condition_3():
    # 班次衔接约束
    # penalized_transitions=[(前一个班次，后一个班次)]
    for previous_shift,next_shift in penalized_transitions:
        for e in range(num_employees):
            for d in range(num_days):
                transition=[
                    work[e,d-1,previous_shift].Not(),work[e,d,next_shift].Not()
                ]
                model.AddBoolOr(transition)

def condition_4():
    # 休息天数约束。从外部读取
    rest_days=num_days-work_days-len(holi_days)
    for e in range(num_employees):
        model.Add(sum(work[e,d,0] for d in range(num_days)  if d not in holi_days)==rest_days)

def condition_5():
    # 法定假日不能全部上班
    for e in range(num_employees):
        model.Add(sum(work[e,d,s] for d in holi_days) > 0)

def condition_6():
    # 每日班次个数要求.从外部读取
    #dayly_cover_demands=[]
    # range(1,num_shifts-2)仅为可变班次
    for s in range(1,num_shifts-2):
        for d in range(num_days):
            works=[work[e,d,s] for e in range(num_employees)]
            min_demand=dayly_cover_demands[d][s-1]
            model.Add(sum(works)==min_demand)
def condition_7():
    # 禁止休-班-休模式
    # 当前因为上周期班表没有设为变量，所以d-1无法用AddBoolOr
    # 下面这个是变通解决上周最后两天分别是 休，休-班这两种情况
    for e in range(num_employees):
        for d in range(num_days-1):
            for s in range(1,num_shifts-2):
                if d == 0:
                    if (e,d-1,0) in lastweek_assignments:
                        model.AddBoolOr([work[e,d+1,0].Not(),work[e,d,s].Not()])
                    elif (e,d-2,0) in lastweek_assignments:
                        model.AddBoolOr([work[e,d,0].Not()])
                else:
                    #model.Add(sum([work[e,d-1,0],work[e,d,s],work[e,d+1,0]])<3)
                    model.AddBoolOr([work[e,d-1,0].Not(),work[e,d+1,0].Not(),work[e,d,s].Not()])


condition_1()  # 最大连续工作天数不超过7天

condition_2()  # 晚班个数约束

condition_3()  # 班次衔接约束

condition_4()  # 休息天数约束

#condition_5()  # 法定假日不能全部上班

condition_6()  # 每日班次个数要求

condition_7()  # 禁止休-班-休模式








print("\nstart solver")
# 规划目标
model.Maximize(
    sum(obj_bool_vars[i]*obj_bool_coeffs[i] for i in range(len(obj_bool_vars)))
)
solver=cp_model.CpSolver()
# 初始化解计数器
solution_count=[0]

# 定义字体样式，颜色为红色
font_style = Font(color='FF0000') # 红色的十六进制代码是 FF0000
# 定义填充样式，颜色为黄色
fill_style = PatternFill(start_color='F2FD99', end_color='F2FD99', fill_type='solid')
fill_style_blue = PatternFill(start_color='CC99FF', end_color='CC99FF', fill_type='solid')
# 创建对齐方式对象并设置水平和垂直居中
align_center = Alignment(horizontal='center', vertical='center')
# 创建边框样式对象，使用默认的细线和黑色
border_style = Border(
    left=Side(style='thin'), 
    right=Side(style='thin'), 
    top=Side(style='thin'), 
    bottom=Side(style='thin')
)
class PartialSolutionPrinter(cp_model.CpSolverSolutionCallback):
    #Print intermediate solutions.
    def __init__(self,shifts,employees,num_days) -> None:
        cp_model.CpSolverSolutionCallback.__init__(self)
        self._shifts=shifts
        self._employees=employees
        self._num_days=num_days
    
    # 输出结果
    def print_value(self):
        rest_meet=0
        duty_meet=0
        # 这里写输出的逻辑
        # 创建一个新的工作簿
        wb = Workbook()
        # 选择活动工作表
        ws = wb.active
        for d in range(self._num_days):
            ws.cell(row=6, column=2+d).value = d+1
            ws.cell(row=6, column=2+d).alignment=align_center
        for e in range(len(self._employees)):
            ws.cell(row=7+e, column=1).value = self._employees[e]
            ws.cell(row=7+e, column=1).alignment=align_center
            ws.cell(row=7+e, column=1).border=border_style
            for d in range(self._num_days):
                for s in range(len(self._shifts)):
                    #这里的取值写法要注意，目前还是自己乱写的
                    if self.BooleanValue(work[e,d,s]):
#                    if solver.BooleanValue(work[e,d,s]):
                        ws.cell(row=7+e, column=2+d).value = self._shifts[s]
                        ws.cell(row=7+e, column=2+d).alignment=align_center
                        ws.cell(row=7+e, column=2+d).border=border_style
                        if any(r[0]==e and r[1]==d and r[2]==s for r in fixed_assignments):
                            ws.cell(row=7+e, column=2+d).fill=fill_style_blue
                        else:
                            if any(r[0]==e and r[1]==d for r in requests):
                                ws.cell(row=7+e, column=2+d).font=font_style
                            if any(r[0]==e and r[1]==d and r[2]==s for r in requests):
                                ws.cell(row=7+e, column=2+d).fill=fill_style
                                if s==0:
                                    rest_meet+=1
                                else:
                                    duty_meet+=1         
        ws.cell(row=1, column=1).value ="score:"
        ws.cell(row=1, column=3).value =self.ObjectiveValue()
        ws.cell(row=2, column=1).value ="conflicts:"
        ws.cell(row=2, column=3).value =self.NumConflicts()
        ws.cell(row=3, column=1).value ="branches:"
        ws.cell(row=3, column=3).value =self.NumBranches()
        ws.cell(row=4, column=1).value ="rest-meet:"
        ws.cell(row=4, column=3).value =str(rest_meet)+" of"+str(rest_count)
        ws.cell(row=5, column=1).value ="shifts-meet:"
        ws.cell(row=5, column=3).value =str(duty_meet)+" of"+str(duty_count)
        filename = f"solution_00.xlsx"
        # 保存工作簿
        wb.save(filename)
        print(f"\nsolution_{self._solution_count} found")
        print(f"-got score    :{self.ObjectiveValue()}")
        print(f"-rest meet    :{rest_meet} of {rest_count}")
        print(f"-shifts meet  :{duty_meet} of {duty_count}")

        wb.close()

    def on_solution_callback(self):
        solution_count[0] += 1
        print(f"找到第{solution_count[0]}个可行解：")
        #print(f"x={self.Value(x)})
        print("目标函数值 =", self.ObjectiveValue())
        print()

# 设置求解时间限制
solver.parameters.max_time_in_seconds = 100.0

# 求解
solution_printer=PartialSolutionPrinter(shifts,employees,num_days)
status=solver.Solve(model,solution_printer)

#检查求解结果
if status == cp_model.OPTIMAL:
    print("找到最优解：")
    print("最优目标函数值 =", solver.ObjectiveValue())
    solution_printer.print_value() #输出结果
elif status == cp_model.FEASIBLE:
    print("找到可行解：")
    print("目标函数值 =", solver.ObjectiveValue())
    solution_printer.print_value() #输出结果    
elif status == cp_model.INFEASIBLE:
    print("无可行解：")
else:
    print("求解未完成。")
