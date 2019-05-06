from openpyxl import load_workbook

wb = load_workbook('students.xlsx')
ws1 = wb['1']
ws2 = wb['2']
ws3 = wb['低保户']
ws4 = wb['挂包人']

for i in range(201,250):
    name1 = ws1["B%d" % i].value #学生姓名
    parent = ws1["G%d" % i].value #监护人
    poor = 0
    low_flag = 0
    for j in range(1,435): #判断是否属于贫困户
        name2 = ws2["C%d" % j].value
        if name1 == name2:
            ws1["J%d" % i].value = "是"
            poor = 1
    if poor == 0:
        ws1["J%d" % i].value = "否"
    if poor == 1:
        for m in range(1,109):
            people = ws4["B%d" % m].value
            if parent == people:
                ws1["L%d" % i].value = ws4["C%d" % m].value
    
    for k in range(1,23): #判断是否属于低保户
        low = ws3["A%d" % k].value
        if parent == low:
            ws1["K%d" % i].value = "是"
            low_flag = 1
    if low_flag == 0:
        ws1["K%d" % i].value = "否"

            
#print(num)
wb.save("students.xlsx")
