import vidget_norm
import rozr_ok
from rozr_ok import *
import pandas as pd
import xlsxwriter
from openpyxl import Workbook
from openpyxl.chart import (Reference,
    Series,
    BarChart3D,)



file_name = str(period.get()) + '.xlsx'


wb = Workbook()
ws = wb.active
cs = wb.create_chartsheet()

rows = [
    ('Собівартість',vpluv_cost ),
    ('Адміністративні витрати', vpluv_admin),
    ('Інші операційні витрати', vpluv_others_expenses),
    ( 'Витрати на збут', vpluv_trade_expenses),
    ( 'Інші операційні доходи', vpluv_others_profit),
    ( 'Інші фінансові доходи', vpluv_others_finprofit),
    ( 'Дохід від реалізації' , vpluv_profit)
]

for row in rows:
    ws.append(row)

data = Reference(ws, min_col=2, min_row=1, max_col=3, max_row=7)
titles = Reference(ws, min_col=1, min_row=1, max_row=7)
chart = BarChart3D()
chart.title = "Вплив факторів"

chart.add_data(data=data, titles_from_data=True)
chart.set_categories(titles)

ws.add_chart(chart, 'E5')
wb.save('graf.xlsx')















'''
wb = Workbook()
ws = wb.active

rows = [
    ('Собівартість',vpluv_cost ),
    ('Адміністративні витрати', vpluv_admin),
    ('Інші операційні витрати', vpluv_others_expenses),
    ( 'Витрати на збут', vpluv_trade_expenses),
    ( 'Інші операційні доходи', vpluv_others_profit),
    ( 'Інші фінансові доходи', vpluv_others_finprofit),
    ( 'Дохід від реалізації' , vpluv_profit)
]





#sheet=wb['Ан']
#sheet['A35'] = ws


for row in rows:
    ws.append(row)
   

data = Reference(ws, min_col=1, min_row=1, max_col=3, max_row=7)
titles = Reference(ws, min_col=1, min_row=2, max_row=7)
chart = BarChart3D()
chart.title = "Вплив факторів"
chart.add_data({
    'categories': '=Ан!$A$27:$A$33',
    'values':     '=Ан!$B$27:$B$33',
    
})
#chart.add_data(data=data, titles_from_data=False)
#chart.set_categories(titles)

ws.add_chart(chart, 'E5')
wb.save(file_name)



'''






'''
#df_pl = pd.DataFrame({ 'lab':['Собівартість', 'Адміністративні витрати','Інші операційні витрати','Витрати на збут',
                                                                # 'Інші операційні доходи', 'Інші фінансові доходи','Дохід від реалізації'],
                     # 'data':[vpluv_cost, vpluv_admin,vpluv_others_expenses,vpluv_trade_expenses,vpluv_others_profit,vpluv_others_finprofit,vpluv_profit]})


df_pl=pd.DataFrame({"data": [vpluv_cost,vpluv_admin,vpluv_others_expenses,
                 vpluv_trade_expenses,vpluv_others_profit,vpluv_others_finprofit,
                 vpluv_profit]})
         
          

file_name = str(period.get()) + '.xlsx'

writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
#writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
df_pl.to_excel(writer, sheet_name='Ан')
# Get xlsxwriter objects
workbook = writer.book


worksheet = writer.sheets['Ан']

# Create a 'column' chart
chart = workbook.add_chart({'type': 'column'})
# select the values of the series and set a name for the series
chart.add_series({
    'lab':['Собівартість', 'Адміністративні витрати','Інші операційні витрати','Витрати на збут',
                                                                 'Інші операційні доходи', 'Інші фінансові доходи','Дохід від реалізації'], 
    'values': [vpluv_cost,vpluv_admin,vpluv_others_expenses,
                 vpluv_trade_expenses,vpluv_others_profit,vpluv_others_finprofit,
                 vpluv_profit], 
    "name": "My Series's Name"
})
# Insert the chart into the worksheet in the D2 cell
worksheet.insert_chart('Ан1', chart)
writer.save()
'''
