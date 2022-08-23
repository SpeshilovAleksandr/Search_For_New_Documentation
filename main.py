import pandas as pd
import datetime
import time


PATH_OLD = '.\\archive of notifications DSO\\notifications DSO 2022-05-26.xlsx'
PATH_NEW = '.\\central archive OTD\\notifications DSO.xlsx'
PATH_DIFF = '.\\new receipts of DSO'

time_start = time.time()

data_old = pd.read_excel(PATH_OLD, skiprows=[0])
data_new = pd.read_excel(PATH_NEW, skiprows=[0])

##print(data_old.columns, end='\n\n')
##print(data_old.index, end='\n\n')
##print(data_old.values, end='\n\n')

##print(data_old['Обозначение документа'][0])
##print(len(data_old.index))

doc = pd.Series(['АБВГ', '666', 'Наимен', 'Москва', 'ЦНИИХМ', '0', '0'],
                index=['Обозначение документа', '№  изм.', 'Наименование',
                       'Город', 'Организация',
                       'Дата ввода информации (фильтр новых поступлений)',
                       'Примечание'])
##print(doc, end='\n\n')

##data_diff = data_old.append(doc, ignore_index=True)

data_diff = pd.DataFrame(
    columns=['Обозначение документа', '№  изм.', 'Наименование', 'Город',
             'Организация', 'Дата ввода информации (фильтр новых поступлений)',
             'Примечание'])
##data_diff = data_diff.append(doc, ignore_index=True)
##print(data_diff)
##data_diff.to_excel('data_diff.xlsx', sheet_name='Новые поступления',
##                   index=False)

### Последний элемент
##print(data_diff.loc[len(data_diff.index) - 1], end='\n\n')  

for i in range(len(data_new.index)):
    match = False  # Совпадение с записью из старого списка.
    for j in range(len(data_old.index)):
        if data_new['Обозначение документа'][i] == \
           data_old['Обозначение документа'][j]:
            match = True
            if data_new['№  изм.'][i] != data_old['№  изм.'][j]:
                data_diff = pd.concat([data_diff,
                                       pd.DataFrame(data_new.loc[i]).T],
                                      ignore_index=True)

    # Если нет записи в старом списке, значит документ - новый.
    if not match:
        data_diff = pd.concat([data_diff, pd.DataFrame(data_new.loc[i]).T],
                              ignore_index=True)

today = datetime.date.today()
name_diff = PATH_DIFF + '\\new receipts of DSO ' + str(today) + '.xlsx'

writer = pd.ExcelWriter(name_diff)
data_diff.to_excel(writer, sheet_name=str(today), index=False)
# Auto-adjust columns' width
for column in data_diff:
    column_width = max(data_diff[column].astype(str).map(len).max(), len(column))
    col_idx = data_diff.columns.get_loc(column)
    writer.sheets[str(today)].set_column(col_idx, col_idx, column_width)
writer.save()

time_finish = time.time()
time_diff = time_finish - time_start
print('Complete! Time:', time_diff, 'секунд.')




















