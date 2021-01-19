import tkinter as tk
import tkinter.ttk as ttk
import pandas as pd
import sys
import os
from tkinter import messagebox
import numpy as np
from tkinter import font
import logging

logging.basicConfig(filename='factorn.log', format='%(asctime)s-%(levelname)s-%(message)s', datefmt='%Y:%m:%d:%H:%M:%S',level=logging.DEBUG)



def export():
    dd={'Показники': ['Дохід від реалізації','Собівартість','Валовий','Адміністративні витрати','Витрати на збут','Інші витрати','Інші доходи',
                      'Інші фінансові доходи','Фінансовий результат'],
        'Попередній період': [float(ent1.get()),float(ent2.get()),float(ent3.get()),float(ent4.get()),float(ent5.get()),float(ent6.get()),float(ent7.get()),float(ent8.get()),float(ent9.get())] ,
        'Поточний період': [float(ent11.get()),float(ent22.get()), float(ent33.get()), float(ent44.get()), float(ent55.get()), float(ent66.get()),
                            float(ent77.get()), float(ent88.get()), float(ent99.get())]}
        
    df =pd.DataFrame(dd, columns=['Показники', 'Попередній період','Поточний період'])
    
    file_name = str(period.get()) + '.xlsx'    
    df.to_excel(file_name, sheet_name='Аналіз', index = False, header=True)
    messagebox.showinfo(title=None,message='Дані завантажено')
    logging.debug(f'Вхідні дані записано до файлу {file_name} ')
    
        
def read_add_upload():
    
    file_name = str(period.get()) + '.xlsx'
    if file_name in os.listdir():
        excel_data_df = pd.read_excel(file_name)
        excel_data_df=np.round(excel_data_df, 1)
        excel_data_df['Відхилення']=excel_data_df['Поточний період']-excel_data_df['Попередній період']        
        excel_data_df['Приріст']=excel_data_df['Поточний період']/excel_data_df['Попередній період']*100-100        
        excel_data_df.to_excel(file_name,sheet_name='Аналіз',float_format="%.1f", index=False)       
        
        messagebox.showinfo(title=None,message='Розрахунок проведено')
        logging.debug(f'Файл {file_name} з записаними вхідними даними зчитано та доданий розрахунок до таблиці з вхідними даними')
    else:
        messagebox.showinfo(title=None,message='Перевірте чи заповнені всі показники! \nПовторіть введення даних')
        logging.debug(f'Файл {file_name} дані не записано, некоректні вхідні дані') 

        
            
def read_add_read():
    file_name = str(period.get()) + '.xlsx'
    if file_name in os.listdir():                
        excel_data_df = pd.read_excel(file_name)
        excel_data_df['Відхилення']=excel_data_df['Поточний період']-excel_data_df['Попередній період']
        excel_data_df['Приріст']=excel_data_df['Поточний період']/excel_data_df['Попередній період']*100-100    
        excel_data_df.to_excel(file_name,sheet_name='Аналіз',float_format="%.1f",  index=False)    
        messagebox.showinfo(title=None,message='Дані завантажено')
        logging.debug(f'Файл {file_name}  з вхідними даними зчитано та доданий розрахунок до таблиці з вхідними даними')
        
    else:        
        messagebox.showinfo(title=None,message='Файл не знайдено') 
        logging.debug(f'Файл з вхідними даними не знайдено')
     
    
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

    if period.get() != '':
        root1.destroy()        
    else:
         messagebox.showinfo(title=None,message='Назва файла не внесено! \nВнесіть назву та продовжуйте')
         logging.debug(f'Дані не записано, не вказано назву файла')         
    def close():
        root.destroy()

    root = tk.Tk()
    root.title('Факторний аналіз')
    w = root.winfo_screenwidth()
    h = root.winfo_screenheight()
    w = w//2
    h = h//2
    w = w - 300
    h = h - 300
    root.geometry('600x500+{}+{}'.format(w,h))
    root["bg"] = "#51856c"

    frame = tk.Frame(root, bg='#7AA899')
    frame['borderwidth'] = 3   
    frame['relief'] = 'ridge'
    frame['border']=5   
    frame.pack(padx=1,pady=1)

    frame1 = tk.Frame(root, bg='#51856c')    
    frame1['borderwidth'] = 20    
    frame1['relief'] = 'flat'
    frame1['border']=10   
    frame1.pack(padx=1,pady=1)


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
    
   
    lab1=tk.Label(frame, text='Дохід від реалізації, тис. грн',bg='#7AA899', font=('Calibri',10, 'bold') )
    lab2=tk.Label(frame, text='Собівартість тис. грн',bg='#7AA899', font=('Calibri',10, 'bold')  )
    lab3=tk.Label(frame, text='Валовий тис. грн ', bg='#7AA899',font=('Calibri',10, 'bold') )
    lab4=tk.Label(frame, text='Адміністративні витрати тис. грн ', bg='#7AA899',font=('Calibri',10, 'bold')  )
    lab5=tk.Label(frame, text='Витрати на збут тис. грн',bg='#7AA899', font=('Calibri',10, 'bold') )
    lab6=tk.Label(frame, text='Інші витрати тис. грн ',bg='#7AA899', font=('Calibri',10, 'bold') )
    lab7=tk.Label(frame, text='Інші доходи тис. грн ',bg='#7AA899', font=('Calibri',10, 'bold') )
    lab8=tk.Label(frame, text='Інші фінансові доходи тис. грн',bg='#7AA899', font=('Calibri',10, 'bold')  )
    lab9=tk.Label(frame, text='Фінансовий результат тис. грн',bg='#7AA899', font=('Calibri',10, 'bold') ) 
    


    lab10=tk.Label(frame, text='Попередній період', bg='#7AA899', font=('Calibri',10, 'bold'))
    lab10.grid(row=0,column=5)
    lab11=tk.Label(frame, text='Поточний період', bg='#7AA899', font=('Calibri',10, 'bold'))
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


    
    ent1 = tk.Entry(frame, textvariable = profit)
    ent2 = tk.Entry(frame, textvariable = cost_price)
    ent3 = tk.Entry(frame, textvariable = val_profit)
    ent4 = tk.Entry(frame, textvariable = admin_expenses)
    ent5 = tk.Entry(frame, textvariable = trade_expenses)
    ent6 = tk.Entry(frame, textvariable = others_expenses)
    ent7 = tk.Entry(frame, textvariable = others_profit)
    ent8 = tk.Entry(frame, textvariable = others_finprofit)
    ent9 = tk.Entry(frame, textvariable = fin_rez)
  
     
    ent11 = ttk.Entry(frame, textvariable = profit1)
    ent22 = ttk.Entry(frame, textvariable = cost_price1)
    ent33 = ttk.Entry(frame, textvariable = val_profit1)
    ent44 = ttk.Entry(frame, textvariable = admin_expenses1)
    ent55 = ttk.Entry(frame, textvariable = trade_expenses1)
    ent66 = ttk.Entry(frame, textvariable = others_expenses1)
    ent77 = ttk.Entry(frame, textvariable = others_profit1)
    ent88 = ttk.Entry(frame, textvariable = others_finprofit1)
    ent99 = ttk.Entry(frame, textvariable = fin_rez1) 


   
    ent1.grid(row = 1, column = 5, columnspan = 2, padx = 5, pady = 5, ipady = 5,  sticky = 'we')
    ent2.grid(row = 2, column = 5, columnspan = 2, padx = 5, pady = 5, ipady = 5,  sticky = 'we')
    ent3.grid(row = 3, column = 5, columnspan = 2, padx = 5, pady = 5, ipady = 5,  sticky = 'we')
    ent4.grid(row = 4, column = 5, columnspan = 2, padx = 5, pady = 5, ipady = 5,  sticky = 'we')
    ent5.grid(row = 5, column = 5, columnspan = 2, padx = 5, pady = 5, ipady = 5,  sticky = 'we')
    ent6.grid(row = 6, column = 5, columnspan = 2, padx = 5, pady = 5, ipady = 5,  sticky = 'we')
    ent7.grid(row = 7, column = 5, columnspan = 2, padx = 5, pady = 5, ipady = 5,  sticky = 'we')
    ent8.grid(row = 8, column = 5, columnspan = 2, padx = 5, pady = 5, ipady = 5,  sticky = 'we')
    ent9.grid(row = 9, column = 5, columnspan = 2, padx = 5, pady = 5, ipady = 5,  sticky = 'we')
 
    
    ent11.grid(row = 1, column = 8, columnspan = 3, padx = 5, pady = 5, ipady = 5,  sticky = 'e')
    ent22.grid(row = 2, column = 8, columnspan = 3, padx = 5, pady = 5, ipady = 5,  sticky = 'e')
    ent33.grid(row = 3, column = 8, columnspan = 3, padx = 5, pady = 5, ipady = 5,  sticky = 'e')
    ent44.grid(row = 4, column = 8, columnspan = 3, padx = 5, pady = 5, ipady = 5,  sticky = 'e')
    ent55.grid(row = 5, column = 8, columnspan = 3, padx = 5, pady = 5, ipady = 5,  sticky = 'e')
    ent66.grid(row = 6, column = 8, columnspan = 3, padx = 5, pady = 5, ipady = 5,  sticky = 'e')
    ent77.grid(row = 7, column = 8, columnspan = 3, padx = 5, pady = 5, ipady = 5,  sticky = 'e')
    ent88.grid(row = 8, column = 8, columnspan = 3, padx = 5, pady = 5, ipady = 5,  sticky = 'e')
    ent99.grid(row = 9, column = 8, columnspan = 3, padx = 5, pady = 5, ipady = 5,  sticky = 'e') 


    
    btn1 = tk.Button(frame1, text = 'Завантажити',relief = 'groove', border = 4, font=('Calibri',12, 'bold'), bg='#7AA899',fg='black',command=export)
    btn1.grid(row = 16, column = 12, ipadx = 3, ipady = 3)
    logging.debug('Вхідні дані завантажено')
    btn3 = tk.Button(frame1, text = 'Розрахувати',relief = 'groove', border = 4,font=('Calibri',12, 'bold'), bg='#7AA899',fg='black',command=read_add_upload)
    btn3.grid(row = 16, column =14, ipadx = 3, ipady = 3)
    btn3.config(command=read_add_upload )

    btn4 = tk.Button(frame1, text = 'Закрити',relief = 'groove', border = 4,font=('Calibri',10, 'bold'),bg='#7AA899', fg='red',command=close)
    btn4.grid(row = 16, column =25, ipadx = 3, ipady = 3,sticky = 'se' )
    btn4.config(command=close )


    
    root.mainloop()
       

def close():
        root1.destroy()
    
root1 = tk.Tk()
root1.title('Вибір шляху введення даних')

w = root1.winfo_screenwidth()
h = root1.winfo_screenheight()
w = w//2
h = h//2
w = w - 200
h = h - 200
root1.geometry('300x200+{}+{}'.format(w,h))
root1["bg"] = "#7AA899"

period=tk.StringVar() 
ent0=tk.Entry(root1,textvariable = period )
ent0.grid(row = 0, column = 2, columnspan = 50, padx = 5, pady = 5, ipady = 5,  sticky = 'we')
lab0=t=tk.Label(root1, text='Назва файла', bg='#7AA899' ,font=('Calibri',12, 'bold') )
lab0.grid(row=0,column=0, )



write = ttk.Radiobutton(root1, text='Занести дані',  command=data)
upload = ttk.Radiobutton(root1, text='Завантажити з файлу',   command=read_add_read)
write.grid(row = 1, column = 12, sticky = 'w', pady = 10)
upload.grid(row = 2, column = 12,sticky = 'w', pady = 10)
logging.info('Вибір способу внесення даних зроблено')
btn5 = tk.Button(root1, text = 'Закрити',relief = 'groove', border = 4,font=('Calibri',10, 'bold'), bg='#7AA899', fg='red',command=close)
btn5.grid(row = 18, column =20)
btn5.config(command=close )





root1.mainloop()


