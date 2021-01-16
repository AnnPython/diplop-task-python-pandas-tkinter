import vidget_norm
import rozr_ok
from rozr_ok import *
import pandas as pd
#import xlsxwriter
from openpyxl import Workbook
from openpyxl.chart import (Reference,Series,BarChart3D,)
from openpyxl import load_workbook
#from openpyxl.styles import NamedStyle
from openpyxl.styles import Font, Color, Alignment, Border, Side, colors
from openpyxl import cell


file_name = str(period.get()) + '.xlsx'
wb = load_workbook(file_name )
sheet = wb.active
ws = wb.active

sheet.column_dimensions["A"].width = 27
sheet.column_dimensions["B"].width = 15
sheet.column_dimensions["C"].width = 15
sheet.column_dimensions["D"].width = 15
sheet.column_dimensions["E"].width = 15

thin = Side(border_style="thin", color="303030") 
black_border = Border(top=thin, left=thin, right=thin, bottom=thin)
font = Font(name='Times New Roman', size=10, bold=False, color='07101c')
align = Alignment(horizontal="center", wrap_text= True, vertical="center")

for label in ["A", "B", "C", "D", "E"]: 
    for col_idx in range(34):
        idx = label + str(col_idx + 1) 
        sheet[idx].alignment = align 
        sheet[idx].font = font 

for row in ws['A2:E10']:
    for cell in row:
        cell.border = black_border


for row in ws['A15:D19']:
    for cell in row:
        cell.border = black_border

for row in ws['A23:C24']:
    for cell in row:
        cell.border = black_border


for row in ws['A27:B34']:
    for cell in row:
        cell.border = black_border         
'''
df.to_excel(writer, sheet_name="Sheet 1"
workbook = load_workbook(file_name )
sheet = workbook.active
format1 = workbook.add_format({'num_format': '0.00'})
sheet.set_column('E:E', 20, format1)
workbook.save(file_name)

'''        
wb.save(file_name)      


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

