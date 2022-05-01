from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import os


# COPY FILE TO NEW ONE
os.system('cp original.xlsx newV2.xlsx')

# FROM ORIGINAL
wb = load_workbook('newV2.xlsx')
ws = wb.active

# NUM TX
print(len(ws['A'])-2, ' transactions loaded')


# FILL ROWS COLORS
def printRow(row, color):

    if color == 'red':
        for i in range(65,76):
            ws[chr(i)+str(row)].fill = PatternFill("solid", start_color="FF0000")
    if color == 'green':
        for i in range(65,76):
            ws[chr(i)+str(row)].fill = PatternFill("solid", start_color="00FF00")

B = 0
D = 0
F = 0
group = []
delete_rows = []
flag_first_in_group = True


# DETECT SAME TIME OPERATIONS
for row in (range(4,len(ws['A'])+1)):

    # CHECK IF SAME DAY AND HOUR  
    if (ws['K'+str(row)].value[0:10] == ws['K'+str(row-1)].value[0:10]) \
        and (ws['C'+str(row)].value == ws['C'+str(row-1)].value) \
        and (ws['E'+str(row)].value == ws['E'+str(row-1)].value) \
        and (ws['H'+str(row)].value == ws['H'+str(row-1)].value) \
        and (ws['A'+str(row)].value == 'Operación' or ws['A'+str(row)].value == 'Otras comisiones'):

        if flag_first_in_group:
            group.append(row-1)
            printRow(row-1, 'red')
            delete_rows.append(row-1)
            flag_first_in_group = False
            try:
                B += ws['B'+str(row-1)].value
            except:
                B = None
            try:
                D += ws['D'+str(row-1)].value
            except:
                D = None
            try:
                F += ws['F'+str(row-1)].value
            except:
                F = None
            
        group.append(row)
        printRow(row, 'red')
        delete_rows.append(row)
        # SUM
        try:
            B += ws['B'+str(row)].value
        except:
            B = None
        try:
            D += ws['D'+str(row)].value
        except:
            D = None
        try:
            F += ws['F'+str(row)].value
        except:
            F = None

    else: 
        # WE GET NO MORE IN SAME GROUP
        if len(group)>0:
            group = []
            flag_first_in_group = True
            # INSERT NEW ROW
            ws.insert_rows(row)
            ws['A'+str(row)] = ws['A'+str(row-1)].value
            ws['B'+str(row)] = B
            ws['C'+str(row)] = ws['C'+str(row-1)].value
            ws['D'+str(row)] = D
            ws['E'+str(row)] = ws['E'+str(row-1)].value
            ws['F'+str(row)] = F
            ws['G'+str(row)] = ws['G'+str(row-1)].value
            ws['H'+str(row)] = ws['H'+str(row-1)].value
            ws['I'+str(row)] = ws['I'+str(row-1)].value
            ws['J'+str(row)] = ws['J'+str(row-1)].value
            ws['K'+str(row)] = ws['K'+str(row-1)].value
            printRow(row, 'green')
            B = 0
            D = 0
            F = 0


wb.save("changesV2.xlsx")


c = 0
## NOW DELETE RED ROWS ON new.xlsx
for row_to_delete in delete_rows:
    ws.delete_rows(row_to_delete-c)
    c+=1


print(len(ws['A'])-2, ' transactions now')
wb.save("newV2.xlsx")
wb.close()
