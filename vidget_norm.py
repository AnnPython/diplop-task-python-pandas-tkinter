import tkinter as tk
from tkinter import filedialog
import tkinter.ttk as ttk
from tkinter import IntVar
import pandas as pd
import sys
import os
from tkinter import messagebox
import numpy as np

def export():
    
    dd={'Показники': ['Дохід від реалізації','Собівартість','Валовий','Адміністративні витрати','Витрати на збут','Інші витрати','Інші доходи',
                      'Інші фінансові доходи','Фінансовий результат'],
        'Попередній період': [float(ent1.get()),float(ent2.get()),float(ent3.get()),float(ent4.get()),float(ent5.get()),float(ent6.get()),float(ent7.get()),float(ent8.get()),float(ent9.get())] ,
        'Поточний період': [float(ent11.get()),float(ent22.get()), float(ent33.get()), float(ent44.get()), float(ent55.get()), float(ent66.get()),
                            float(ent77.get()), float(ent88.get()), float(ent99.get())]}
        
    df =pd.DataFrame(dd, columns=['Показники', 'Попередній період','Поточний період'])
    file_name = str(period.get()) + '.xlsx'
    df.to_excel(file_name, sheet_name='Ан', index = False, header=True)
    messagebox.showinfo(title=None,message='Дані завантажено')
    

def read_add_upload():
    file_name = str(period.get()) + '.xlsx'
    if file_name in os.listdir():
        excel_data_df = pd.read_excel(file_name)
        
        excel_data_df['Відхилення']=excel_data_df['Поточний період']-excel_data_df['Попередній період']
        excel_data_df=np.round(excel_data_df, 1)
        excel_data_df['Приріст']=excel_data_df['Поточний період']/excel_data_df['Попередній період']*100-100         
        excel_data_df.to_excel(file_name,sheet_name='Ан', index=False)       
        
        messagebox.showinfo(title=None,message='Дані завантажено')

def read_add_read():
    file_name = str(period.get()) + '.xlsx'
    if file_name in os.listdir():                
        excel_data_df = pd.read_excel(file_name)  
        excel_data_df['Відхилення']=excel_data_df['Поточний період']-excel_data_df['Попередній період']
        excel_data_df['Приріст']=excel_data_df['Поточний період']/excel_data_df['Попередній період']*100-100    
        excel_data_df.to_excel(file_name,sheet_name='Ан', index=False)
    
        messagebox.showinfo(title=None,message='Дані завантажено')
            
    
def data():
    global period

    global ent0
    global ent8
    global ent9
    global ent1 
    global ent2 
    global ent3 
    global ent4 
    global ent5 
    global ent6 
    global ent7
    global ent8
    global ent9    
    global ent11
    global ent22 
    global ent33 
    global ent44 
    global ent55 
    global ent66 
    global ent77
    global ent88 
    global ent99
          
    root1.destroy()
    root = tk.Tk()
    root.title('факторний ')
    root.geometry('650x500')

    
    profit= tk.StringVar()  
    cost_price = tk.StringVar()  
    val_profit= tk.StringVar()  
    admin_expenses = tk.StringVar()  
    trade_expenses = tk.StringVar()  
    others_expenses= tk.StringVar()  
    others_profit = tk.StringVar()
    others_finprofit = tk.StringVar()
    fin_rez = tk.StringVar() 

 

    profit1= tk.StringVar()  
    cost_price1 = tk.StringVar()  
    val_profit1= tk.StringVar()  
    admin_expenses1 = tk.StringVar()  
    trade_expenses1 = tk.StringVar()  
    others_expenses1= tk.StringVar()  
    others_profit1 = tk.StringVar()
    others_finprofit1 = tk.StringVar()
    fin_rez1 = tk.StringVar() 

    
   
    lab1=ttk.Label(root, text='Дохід від реалізації')
    lab2=ttk.Label(root, text='Собівартість')
    lab3=ttk.Label(root, text='Валовий')
    lab4=ttk.Label(root, text='Адміністративні витрати')
    lab5=ttk.Label(root, text='Витрати на збут')
    lab6=ttk.Label(root, text='Інші витрати')
    lab7=ttk.Label(root, text='Інші доходи')
    lab8=ttk.Label(root, text='Інші фінансові доходи')
    lab9=ttk.Label(root, text='Фінансовий результат') 
    


    lab10=ttk.Label(root, text='Попередній період')
    lab10.grid(row=0,column=5)
    lab11=ttk.Label(root, text='Поточний період')
    lab11.grid(row=0,column=9)
    
    
    lab1.grid(row=1,column=3)
    lab2.grid(row=2,column=3)
    lab3.grid(row=3,column=3)
    lab4.grid(row=4,column=3)
    lab5.grid(row=5,column=3)
    lab6.grid(row=6,column=3)
    lab7.grid(row=7,column=3)
    lab8.grid(row=8,column=3)
    lab9.grid(row=9,column=3)


    
    ent1 = tk.Entry(root, textvariable = profit)
    ent2 = tk.Entry(root, textvariable = cost_price)
    ent3 = tk.Entry(root, textvariable = val_profit)
    ent4 = tk.Entry(root, textvariable = admin_expenses)
    ent5 = tk.Entry(root, textvariable = trade_expenses)
    ent6 = tk.Entry(root, textvariable = others_expenses)
    ent7 = tk.Entry(root, textvariable = others_profit)
    ent8 = tk.Entry(root, textvariable = others_finprofit)
    ent9 = tk.Entry(root, textvariable = fin_rez)
  
     
    ent11 = ttk.Entry(root, textvariable = profit1)
    ent22 = ttk.Entry(root, textvariable = cost_price1)
    ent33 = ttk.Entry(root, textvariable = val_profit1)
    ent44 = ttk.Entry(root, textvariable = admin_expenses1)
    ent55 = ttk.Entry(root, textvariable = trade_expenses1)
    ent66 = ttk.Entry(root, textvariable = others_expenses1)
    ent77 = ttk.Entry(root, textvariable = others_profit1)
    ent88 = ttk.Entry(root, textvariable = others_finprofit1)
    ent99 = ttk.Entry(root, textvariable = fin_rez1) 


   
    ent1.grid(row = 1, column = 5, columnspan = 3, padx = 5, pady = 5, ipady = 5,  sticky = 'we')
    ent2.grid(row = 2, column = 5, columnspan = 3, padx = 5, pady = 5, ipady = 5,  sticky = 'we')
    ent3.grid(row = 3, column = 5, columnspan = 3, padx = 5, pady = 5, ipady = 5,  sticky = 'we')
    ent4.grid(row = 4, column = 5, columnspan = 3, padx = 5, pady = 5, ipady = 5,  sticky = 'we')
    ent5.grid(row = 5, column = 5, columnspan = 3, padx = 5, pady = 5, ipady = 5,  sticky = 'we')
    ent6.grid(row = 6, column = 5, columnspan = 3, padx = 5, pady = 5, ipady = 5,  sticky = 'we')
    ent7.grid(row = 7, column = 5, columnspan = 3, padx = 5, pady = 5, ipady = 5,  sticky = 'we')
    ent8.grid(row = 8, column = 5, columnspan = 3, padx = 5, pady = 5, ipady = 5,  sticky = 'we')
    ent9.grid(row = 9, column = 5, columnspan = 3, padx = 5, pady = 5, ipady = 5,  sticky = 'we')
 
    
    ent11.grid(row = 1, column = 8, columnspan = 3, padx = 5, pady = 5, ipady = 5,  sticky = 'we')
    ent22.grid(row = 2, column = 8, columnspan = 3, padx = 5, pady = 5, ipady = 5,  sticky = 'we')
    ent33.grid(row = 3, column = 8, columnspan = 3, padx = 5, pady = 5, ipady = 5,  sticky = 'we')
    ent44.grid(row = 4, column = 8, columnspan = 3, padx = 5, pady = 5, ipady = 5,  sticky = 'we')
    ent55.grid(row = 5, column = 8, columnspan = 3, padx = 5, pady = 5, ipady = 5,  sticky = 'we')
    ent66.grid(row = 6, column = 8, columnspan = 3, padx = 5, pady = 5, ipady = 5,  sticky = 'we')
    ent77.grid(row = 7, column = 8, columnspan = 3, padx = 5, pady = 5, ipady = 5,  sticky = 'we')
    ent88.grid(row = 8, column = 8, columnspan = 3, padx = 5, pady = 5, ipady = 5,  sticky = 'we')
    ent99.grid(row = 9, column = 8, columnspan = 3, padx = 5, pady = 5, ipady = 5,  sticky = 'we') 


    
    btn1 = ttk.Button(root, text = 'Завантажити', command=export)
    btn1.grid(row = 10, column = 7)

    btn3 = ttk.Button(root, text = 'розрахувати', command=read_add_upload)
    btn3.grid(row = 14, column =7)

    root.mainloop()
       
    
root1 = tk.Tk()
root1.title('Вибір')
root1.geometry('300x150')
period=tk.StringVar() 
ent0=ttk.Entry(root1,textvariable = period )
ent0.grid(row = 0, column = 2, columnspan = 50, padx = 5, pady = 5, ipady = 5,  sticky = 'we')
lab0=ttk.Label(root1, text='Назва файла')
lab0.grid(row=0,column=0)

write = ttk.Radiobutton(root1, text='Занести дані',  command=data)
upload = ttk.Radiobutton(root1, text='завантажити з файлу',  command=read_add_read)
write.grid(row = 1, column = 12, sticky = 'w', pady = 10)
upload.grid(row = 2, column = 12,sticky = 'w', pady = 10)

root1.mainloop()


