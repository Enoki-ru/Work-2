```python
import pandas as pd
import numpy as np
import os 
import shutil
import openpyxl

files = os.listdir() # Выведем список файлов в корневой папке
excels = list(filter(lambda x: x.endswith('.xlsx'), files)) #оставим среди них те, что являются excel файлом
excels = list(filter(lambda x: x.startswith('склад'), excels)) #оставим только наш файл среди всех

assert len(excels)==1 , 'В корневой папке содержится больше чем одна таблица с названием, начинающемся на (склад)' #выведем ошибку, если есть проблемы

excel=excels[0]
name_main_file=excel
print(f'Найден основной файл! {excel}')

name_folder = 'saves'
os.makedirs(name_folder, exist_ok=True)

db=pd.read_excel(excel)

last_column=db.iloc[:,-1].name # Посмотрим, как называется последний столбец
new_last='Примечание' 
db=db.rename(columns={last_column:new_last}) #переименуем его
last_column=new_last
#db[last_column]
db[last_column]=db[last_column].notna()
db=db[db[last_column]==True]
db.drop(db.columns[2:10], axis=1, inplace=True)
db=db.fillna(0)
db=db.reset_index(drop=True)

print(db.head(2))

name_read_file='Предыдущая таблица.xlsx'
files=os.listdir(path=name_folder)
excels = list(filter(lambda x: x.endswith('.xlsx'), files)) #оставим среди них те, что являются excel файлом
excels = list(filter(lambda x: x.startswith(name_read_file), excels)) #оставим только наш файл среди всех
print(excels)
path_read_file=name_folder+'/'+name_read_file

def save_file(db,name_file):
    from openpyxl import Workbook 
    # Добавляем библиотеку для записи в Excel с изменением названий листов
    #for name, db in db.items():
    with pd.ExcelWriter(name_file) as writer:  
        db.to_excel(writer, sheet_name='TDsheet')
        
def change_width(full_path):
    wb=openpyxl.load_workbook(filename=full_path)
    sheets=wb.sheetnames
    print(f"---------------------\nМодуль с подгоном ширины столбов Excel")
    for i in range (len(sheets)):
        sheet=sheets[i]
        wb.active=i
        ws=wb.active
        for column_cells in ws.columns:
            length = max(len(str(cell.value)) for cell in column_cells)
            print(f"Ширина столбца в {sheet} = {length}")
            if length>60:
                length=60
            ws.column_dimensions[column_cells[0].column_letter].width = length*1.2
    print(f"Конец модуля\n------------------------------")
    wb.active=0
    wb.save(full_path)

# def word_count(string):
#     return(len(string.strip().split(" ")))

path_read_file=name_folder+'/'+name_read_file

if len(excels)==0:
    print('С чем сравнивать то?')
    print('Вставим необработанные данные...')
    shutil.copyfile(name_main_file,path_read_file)
    # name_template=name_folder+'/'+name_main_file
    # os.rename(name_template,path_read_file)

    #shutil.move(name_read_file,name_folder)

    print('Успешно!\n-------------')

    print(f'Внимание! Первоначальные данные сохранены. Отчет будет создан при внесении следующего файла. Найти предыдущие файлы можете в папке {name_folder}')
    #os.system("pause")
else:
    db_read=pd.read_excel(path_read_file)
    last_column=db_read.iloc[:,-1].name # Посмотрим, как называется последний столбец
    new_last='Примечание' 
    db_read=db_read.rename(columns={last_column:new_last}) #переименуем его
    last_column=new_last
    #db_read[last_column]
    db_read[last_column]=db_read[last_column].notna()
    db_read=db_read[db_read[last_column]==True]
    db_read.drop(db_read.columns[2:10], axis=1, inplace=True)
    db_read=db_read.fillna(0)
    db_read=db_read.reset_index(drop=True)

    print(db_read.head(2))

    def similar_finder(db_main,db_read,dic,name_dic):
        dic[name_dic] = pd.DataFrame(columns=["Номенклатура",'Было','Стало',"Разница","Примечание"])
        similar=list(set(db_main['Номенклатура']) & set(db_read['Номенклатура']))
        for i in range(len(db_main)):
            word=db_main['Номенклатура'][i]
            con=''
            if word in similar:
                row_main = db_main[db_main['Номенклатура'] == word].index.tolist()[0]
                row_read= db_read[db_read['Номенклатура'] == word].index.tolist()[0]
                count_main=db_main['1.Основной склад "ВентЭл"'][row_main]
                count_read=db_read['1.Основной склад "ВентЭл"'][row_read]
                delta=-count_read+count_main
                
                if delta!=0:
                    dic2=pd.DataFrame({"Номенклатура":[word],'Было':[count_read],'Стало':count_main,"Разница":[delta],"Примечание":con })
                    dic[name_dic]=dic[name_dic].append(dic2, ignore_index= True)

        not_similar=list(set(db_main['Номенклатура']) - set(db_read['Номенклатура']))
        for i in range(len(db_main)):
            word=db_main['Номенклатура'][i]
            if word in not_similar:
                row_main=db_main[db_main['Номенклатура'] == word].index.tolist()[0]
                count_main=db_main['1.Основной склад "ВентЭл"'][row_main]
                count_read=0
                con='Товара раньше в списке не было'
                delta=-count_read+count_main
                dic2=pd.DataFrame({"Номенклатура":[word],'Было':[count_read],'Стало':count_main,"Разница":[delta],"Примечание":con })
                dic[name_dic]=dic[name_dic].append(dic2, ignore_index= True)
        not_similar=list(set(db_read['Номенклатура']) - set(db_main['Номенклатура']))
        for i in range(len(db_read)):
            word=db_read['Номенклатура'][i]
            if word in not_similar:
                row_read=db_read[db_read['Номенклатура'] == word].index.tolist()[0]
                count_read=db_read['1.Основной склад "ВентЭл"'][row_read]
                count_main=0
                con='Товара в текущем списке нет'
                delta=-count_read+count_main
                dic2=pd.DataFrame({"Номенклатура":[word],'Было':[count_read],'Стало':count_main,"Разница":[delta],"Примечание":con })
                dic[name_dic]=dic[name_dic].append(dic2, ignore_index= True)            
                
        return dic

    import warnings
    warnings.filterwarnings('ignore') # Добавил тк без этого будет выводиться надпись, что в будущем метод pd.append не будет использоваться

    dic={}
    dic=similar_finder(db,db_read,dic,'Изменения')
    save_file(dic['Изменения'],'Готовый отчет.xlsx')
    change_width('Готовый отчет.xlsx')

    #os.remove(path_read_file)
 
    path_last_file=name_folder+'/'+'Запасная таблица.xlsx'
    shutil.copyfile(path_read_file,path_last_file)
    shutil.copyfile(name_main_file,path_read_file)

#os.system("pause")
```

    Найден основной файл! склад 12.01.xlsx
                      Номенклатура  1.Основной склад "ВентЭл"  Примечание
    0                          LFT                        441        True
    1  LFT(d) 2E-146-67 Вентилятор                        236        True
    ['Предыдущая таблица.xlsx']
                      Номенклатура  1.Основной склад "ВентЭл"  Примечание
    0                          LFT                        441        True
    1  LFT(d) 2E-146-67 Вентилятор                        236        True
    ---------------------
    Модуль с подгоном ширины столбов Excel
    Ширина столбца в TDsheet = 4
    Ширина столбца в TDsheet = 12
    Ширина столбца в TDsheet = 4
    Ширина столбца в TDsheet = 5
    Ширина столбца в TDsheet = 7
    Ширина столбца в TDsheet = 10
    Конец модуля
    ------------------------------
    


```python

```
