import openpyxl

B3_53 = []
C3_53 = []
D3_53 = []
I3_53 = []
E3_50 = []
F3_50 = []
G1_50 = []
H1_50 = []

for i in range(3, 60):
    B3_53.append('B' + str(i))
# print(B3_53)

student_name = []
name = '23大队第十二周个人周统表.xlsx'
wb = openpyxl.load_workbook(name)
worksheet = wb.worksheets
sheet1 = worksheet[10]
for i in range(len(B3_53)):
    if sheet1[B3_53[i]].value is not None:
        student_name.append(sheet1[B3_53[i]].value)
# print(student_name)

for i in range(len(student_name) + 3):
    E3_50.append('E' + str(i + 3))
# print(E1_50)

for i in range(len(student_name) + 3):
    F3_50.append('F' + str(i + 3))
# print(F1_50)

for i in range(len(student_name) + 2):
    G1_50.append('G' + str(i + 1))
# print(G1_50)

for i in range(len(student_name) + 2):
    H1_50.append('H' + str(i + 1))
# print(H1_50)
for i in range(len(student_name) + 2):
    C3_53.append('C' + str(i + 3))
# print(C3_53)

for i in range(len(student_name) + 3):
    D3_53.append('D' + str(i + 3))

for i in range(len(student_name) + 3):
    I3_53.append('I' + str(i + 3))
# print(D3_53)
sheet1['E1'] = '('
sheet1['E2'] = '-0.15'
sheet1['F1'] = ')'
sheet1['G1'] = '、'
sheet1['G2'] = '扣分'
sheet1['H2'] = '通报'
wb.save(name)
print(student_name)

for i in range(len(student_name)):

    # 扣分
    if sheet1[C3_53[i]].value is not None:
        # print(sheet1[C3_53[i]].value)
        sheet1[E3_50[i]] = f'=C{i + 3}*$E$2'
        sheet1[G1_50[i+2]] = f'=E{i + 3}&($E$1&$G$2&$F$1)'
    # 通报
    if sheet1[D3_53[i]].value is not None:
        sheet1[F3_50[i]] = f'=D{i + 3}*$E$2'
        sheet1[H1_50[i+2]] = f'=F{i + 3}&($E$1&$H$2&$F$1)'

    # 通报
    wb.save(name)
    if sheet1[E3_50[i]].value is not None and sheet1[F3_50[i]].value is not None:
        sheet1[I3_53[i]] = f'=(E{i + 3}+F{i + 3})&($E$1&$G$2&$G$1&$H$2&$F$1)'
    wb.save(name)
