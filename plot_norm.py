import vidget_norm
import rozr_ok
from rozr_ok import *
import pandas as pd
import xlsxwriter
from openpyxl import Workbook
from openpyxl.chart import (Reference,
    Series,
    BarChart3D,)
from openpyxl import load_workbook
from openpyxl.styles import NamedStyle
from openpyxl.styles import Font, Color, Alignment, Border, Side, colors
from openpyxl import cell
#from openpyxl.cell import get_column_letter, column_index_from_string

file_name = str(period.get()) + '.xlsx'
wb = load_workbook(file_name )
sheet = wb.active
ws = wb.active

sheet = wb.get_sheet_by_name('Ан')
tuple(sheet['A1':'D35'])

for rowOfCellObjects in sheet['A1':'D35']:
    max_length = 0
    column = rowOfCellObjects[0].column
    
    for cellObj in rowOfCellObjects:
        if len(str(cellObj.value)) > max_length:
            max_length = len(cellObj.value)
        
            adjusted_width = (max_length + 2) * 1.2
            ws.column_dimensions[column].width = adjusted_width




'''
for rowOfCellObjects in sheet['A1':'D35']:
    for cellObj in rowOfCellObjects:
       if cellObj is not None:
           ws.column_dimensions['A'].width = 35
       elif cellObj is not None:
           ws.row_dimensions[1].width = 15
       else:
            pass
               
'''

'''
for col in sheet.columns:
    for j in  range(len(col)):
        if j==0:
            ws.column_dimensions.width = 10
        elif j>=4:
            ws.column_dimensions.width = 35
            

'''

'''
ws.column_dimensions['A'].width = 35
ws.column_dimensions['B'].width = 20
ws.column_dimensions['c'].width = 20
ws.column_dimensions['D'].width = 20
'''




wb.save(file_name)
'''
ns = NamedStyle(name='highlight')
ns.font = Font(bold=True, size=14)
border = Side(style='thick', color='000000')
ns.border = Border(left=border, top=border, right=border, bottom=border)


wb.add_named_style(ns)

for col in sheet.columns:
        for cell in col:
            sheet.style = 'highlight'



'''






'''work
ws.column_dimensions['A'].width = 35
ws.column_dimensions['B'].width = 25
'''
'''
ws.column_dimensions['A'].width = 35
ws.row_dimensions[1].width = 25
'''
'''
# Create a few styles
bold_font = Font(bold=True)

center_aligned_text = Alignment(horizontal="center")
double_border_side = Side(border_style="double")
square_border = Border(top=double_border_side,
                        right=double_border_side,
                        bottom=double_border_side,
                        left=double_border_side)


for col in sheet.columns:
        for cell in col:
            #column_dimensions.width = 35
            #row_dimensions.width = 25
            alignment = center_aligned_text
            border = square_border
          
            
'''

'''
 # Style some cells!
sheet['A2'].font = bold_font

sheet['A2'].alignment = center_aligned_text
sheet['A2'].border = square_border

'''








wb1 = load_workbook(file_name )
ws = wb1.active

rows = [
    ('Собівартість',vpluv_cost ),
    ('Адміністративні витрати', vpluv_admin),
    ('Інші операційні витрати', vpluv_others_expenses),
    ( 'Витрати на збут', vpluv_trade_expenses),
    ( 'Інші операційні доходи', vpluv_others_profit),
    ( 'Інші фінансові доходи', vpluv_others_finprofit),
    ( 'Дохід від реалізації' , vpluv_profit)
]

data = Reference(ws, min_col=2, min_row=26, max_col=2, max_row=33)
titles = Reference(ws, min_col=1, min_row=27, max_row=33)
chart = BarChart3D()
chart.title = "Вплив факторів"

chart.add_data( data,titles_from_data=True)
chart.set_categories(titles)

ws.add_chart(chart, 'H4')
wb1.save(file_name)

