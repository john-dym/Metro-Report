# Python 3.7
from tkinter import filedialog
from xlrd import open_workbook
import openpyxl
from openpyxl.styles import Border, GradientFill, Alignment, Side, borders
from os.path import getmtime
from time import localtime, asctime

# Doors, column labels, metro numbers' variables (Planning to replace with ini file)
met1 = 'MET015'
met2 = 'MET021'
met3 = 'MET031'
a1 = 'a7'
a2 = 'a8'
a3 = 'a9'
cPartNo = 'Part No'
cPartName = 'Part Name'
cEO = 'EO No'
cLot = 'LOT No'
cLoc = 'Loc. No'
cCaseNo = 'Case No'

#Opens a file dialog and sets the variable wbFile with the path name
wbFile = filedialog.askopenfilename()
wb = open_workbook(wbFile)
ws = wb.sheet_by_index(0)

# Finds the column number for Part No, Part Name, EO, Lot, Loc, Case No.
for col_index in range(ws.ncols):
    if ws.cell(0,col_index).value == cPartNo:
        cPartNo = int(col_index)
        break
for col_index in range(ws.ncols):
    if ws.cell(0,col_index).value == cPartName:
        cPartName = int(col_index)
        break
for col_index in range(ws.ncols):
    if ws.cell(0,col_index).value == cEO:
        cEO = int(col_index)
        break
for col_index in range(ws.ncols):
    if ws.cell(0,col_index).value == cLot:
        cLot = int(col_index)
        break
for col_index in range(ws.ncols):
    if ws.cell(0,col_index).value == cLoc:
        cLoc = int(col_index)
        break

#Defines the list variables
met = [met1, met2, met3]
combPartNos = []
combPartNames = []
combPartLocs = []
combPartQtys = []
combPartEOs = []

# Iterates through the rows to find unique parts for matching locations
for x in range(len(met)):
#Declares and erases the lists
    partNos = []
    partNames = []
    partLocs = []
    partQtys = []
    partEOs = []
    for i in range(ws.nrows):
        if ws.cell(i, cLoc).value == met[x]:
         #Increases the box count for the part
            if ws.cell(i, cPartNo).value in partNos:
                partQtys[partNos.index(ws.cell(i, cPartNo).value)] = partQtys[partNos.index(ws.cell(i, cPartNo).value)] + 1
         #Adds new part to the list and sets the box count to 1 and EO count to 0
            else:
                partNames.append(ws.cell(i, cPartName).value)
                partNos.append(ws.cell(i, cPartNo).value)
                partLocs.append(ws.cell(i, cLoc).value)
                partQtys.append(1)
                partEOs.append(0)
        #Checks to see if the EO or Lot cell is blank, If not blank it adds 1 to the EO/Lot Qty
            if (ws.cell(i, cEO).value != '' or ws.cell(i, cLot).value != ''):
                partEOs[partNos.index(ws.cell(i, cPartNo).value)] = partEOs[partNos.index(ws.cell(i, cPartNo).value)] + 1
#Combines lists from previous loops
    combPartNos = combPartNos + partNos
    combPartNames = combPartNames + partNames
    combPartLocs = combPartLocs + partLocs
    combPartQtys = combPartQtys + partQtys
    combPartEOs = combPartEOs + partEOs

wb = openpyxl.Workbook()
tmpfile = 'tmp.xlsx'
ws = wb.active
ws['A1'] = 'Metro Box List'
ws.merge_cells('A1:F1')
ws['G1'] =  asctime(localtime(getmtime(wbFile)))
ws['A2'] = 'Part Number'
ws['B2'] = 'Part Name'
ws['C2'] = 'Boxes'
ws['D2'] = 'EO/Lot Qty'
ws['E2'] = 'Loc. No'
ws['F2'] = 'Door'
ws['A1']. = Border(left=Side(style='medium'), right=Side(style='medium'), top=Side(style='medium'), bottom=Side(style='medium'))

for i in range(3,len(combPartNos)):
    ws['A'+str(i)] = combPartNos[i - 3]
    ws['B'+str(i)] = combPartNames[i - 3]
    ws['C'+str(i)] = combPartQtys[i - 3]
    if combPartEOs[i - 3] != 0:
        ws['D'+str(i)] = combPartEOs[i - 3]
    ws['E'+str(i)] = combPartLocs[i - 3]
    # ws['F'+str(i)] = combPartDoors[i - 3]

wb.save(filename = tmpfile)

# #For debugging
# print()
# for z in range(len(combPartNos)):
#     print(str(combPartNos[z]) +', ' + str(combPartNames[z]) + ', ' + str(combPartQtys[z]) + ', ' + str(combPartEOs[z]) + ', ' + str(combPartLocs[z]))

