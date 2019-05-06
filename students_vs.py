from openpyxl import load_workbook
from openpyxl.styles import Font
wb = load_workbook('children.xlsx')
ws1 = wb['1']
ws2 = wb['2']

font = Font(color='ff0000')
for i in range(1,237):
    name2 = ws2["B%d" % i].value
  

    for j in range(1,236):
        name1 = ws1["B%d" % j].value
              
        if name1 == name2:
            ws1["B%d" % i].font = font
            ws2["B%d" % j].font = font
            
            break
          
wb.save("children.xlsx")
