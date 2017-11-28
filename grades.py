from openpyxl import Workbook
class Category:
    def __init__( self, name, num, weight ):
        self.name = name
        self.num = num
        self.weight = weight
sheetName = raw_input("Name of Course: ")
numCategories = raw_input("Number of categories: ")
categories = list()
for x in range(0, int(numCategories)):
    name = raw_input("Category " + str(x) + ": ")
    num = raw_input("Number of elements " + str(x) + ": ")
    weight = raw_input("Weight of category (Ex: 15 = 15%) " + str(x) + ": ")
    categories.append(Category(name, num, weight))
wb = Workbook()
ws = wb.active
ws.title = sheetName
col = 1
j = int(1)
ws.cell(row=int(categories[0].num) + 4, column=2).value = '=SUM('
for i in range( 1, int(numCategories)+1):
    ws.cell( row=1, column=col ).value = categories[i-1].name
    for j in range(1, int(categories[i-1].num)+1):
        ws.cell(row=j+1, column=col).value = categories[i-1].name + " " + str(j)
    ws.cell(row=j+2, column=col).value = categories[i-1].name + " Average: "
    ws.cell(row=j+2, column=col+1).value = "=AVERAGE(" + ws.cell(row=2, column=col+1).coordinate + ":" + ws.cell(row=j+1, column=col+1).coordinate + ")"
    ws.cell(row=int(categories[0].num) + 4, column=2).value += ',' + ws.cell(row=j+2, column=col+1).coordinate + '*' + str(float(categories[i-1].weight)/100)
    col += 2
ws.cell(row=int(categories[0].num) + 4, column=2).value += ')' 
ws.cell(row=int(categories[0].num) + 4, column=1).value = 'Final Grade'
wb.save(sheetName + ".xlsx")
