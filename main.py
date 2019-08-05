# Python 3.7
from tkinter import filedialog
from xlrd import open_workbook
from os.path import getmtime
from os import system
from time import localtime, asctime

import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment

# Doors, column labels, metro numbers' variables (Planning to replace with ini file)
met1 = 'MET015'
met2 = 'MET021'
met3 = 'MET031'
a1 = 'A7'
a2 = 'A8'
a3 = 'A9'
met = [met1, met2, met3]
door = [a1, a2, a3]
xlCol = ['A','B','C','D','E','F']
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
combPartNos = []
combPartNames = []
combPartLocs = []
combPartQtys = []
combPartEOs = []
combPartDoors = []

# Iterates through the rows to find unique parts for matching locations
for x in range(len(met)):
#Declares and erases the lists
    partNos = []
    partNames = []
    partLocs = []
    partQtys = []
    partEOs = []
    partDoors = []
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
                partDoors.append(door[x])
        #Checks to see if the EO or Lot cell is blank, If not blank it adds 1 to the EO/Lot Qty
            if (ws.cell(i, cEO).value != '' or ws.cell(i, cLot).value != ''):
                partEOs[partNos.index(ws.cell(i, cPartNo).value)] = partEOs[partNos.index(ws.cell(i, cPartNo).value)] + 1
#Combines lists from previous loops
    combPartNos = combPartNos + partNos
    combPartNames = combPartNames + partNames
    combPartLocs = combPartLocs + partLocs
    combPartQtys = combPartQtys + partQtys
    combPartEOs = combPartEOs + partEOs
    combPartDoors = combPartDoors + partDoors

wb = openpyxl.Workbook()
ws = wb.active
thin = Side(border_style='thin', color='000000') #Creates a variable for setting borders


ws['A1'] = 'Metro Box List ' + str(asctime(localtime(getmtime(wbFile)))) #Writes the date/timestamp of the file exported from database in G1
ws['A1'].alignment = Alignment(vertical='center',horizontal='center')
ws.merge_cells('A1:F1')
ws['A2'] = 'Part Number'
ws['B2'] = 'Part Name'
ws['C2'] = 'Boxes'
ws['D2'] = 'EO/Lot Qty'
ws['E2'] = 'Loc. No'
ws['F2'] = 'Door'

for x in range(ws.max_column):
    for i in range(ws.max_row):
        ws[xlCol[x]+str(i+1)].fill = PatternFill(start_color='e6e6e6', fill_type='solid')

for i in range(3,len(combPartNos)):
    ws['A'+str(i)] = combPartNos[i - 3]
    ws['B'+str(i)] = combPartNames[i - 3]
    ws['C'+str(i)] = combPartQtys[i - 3]
    if combPartEOs[i - 3] != 0:         #Ignores writing a 0 in the cell
        ws['D'+str(i)] = combPartEOs[i - 3]
        ws['D'+str(i)].fill = PatternFill(start_color='e6e6e6', fill_type='solid')
    ws['E'+str(i)] = combPartLocs[i - 3]
    ws['F'+str(i)] = combPartDoors[i - 3]

#Apply Borders to all cells
for x in range(ws.max_column):
    for i in range(ws.max_row):
        ws[xlCol[x]+str(i+1)].border = Border(top=thin, bottom=thin, left=thin, right=thin)

ws.column_dimensions['A'].width = 15
ws.column_dimensions['B'].width = 45
ws.column_dimensions['C'].width = 6
ws.column_dimensions['D'].width = 10
ws.column_dimensions['E'].width = 10
ws.column_dimensions['F'].width = 10

ws.print_options.horizontalCentered = True
ws.print_options.verticalCentered = True
ws.page_margins.top = 0.3
ws.page_margins.bottom = 0.3
ws.page_margins.left = 0.3
ws.page_margins.right = 0.3
ws.page_margins.header = 0
ws.page_margins.footer = 0

wb.save(filename = 'tmp.xlsx')
system('start ' + 'tmp.xlsx')
