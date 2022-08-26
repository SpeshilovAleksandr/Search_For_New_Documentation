import sys
import time
import datetime
import pandas as pd


PATH_OLD = '.\\archive of notifications DSO\\notifications DSO 2022-07-29.xlsx'
PATH_NEW = '.\\central archive OTD\\notifications DSO.xlsx'
##PATH_OLD = '.\\archive of notifications DSO\\test_file_old.xlsx'
##PATH_NEW = '.\\central archive OTD\\test_file_new.xlsx'
PATH_DIFF = '.\\new receipts of DSO'


def frame_to_dict(frame: pd.DataFrame, keys: str):
    """Convert frame to dictionary of dictionaries by keys."""
    dict_of_dict = dict()
    for i in range(len(frame.index)):
        key = frame[keys][i]
        if key in dict_of_dict and \
           frame['Дата ввода информации (фильтр новых поступлений)'][i] <= \
           dict_of_dict[key]['Дата ввода информации (фильтр новых поступлений)']:
            continue
        else:  # date in frame > date in dictionary, or not in dictionary
            dict_of_dict[key] = dict()
            for column in frame.columns:
                if column != keys:
                    dict_of_dict[key][column] = frame[column][i]
    return dict_of_dict


def trivial_difference(data_new: pd.DataFrame, data_old: pd.DataFrame):
    """Asymptotic O(N^2)."""
    data_diff = pd.DataFrame(columns=data_new.columns)
    
    for i in range(len(data_new.index)):
        match = False  # Match an entry from the old list.
        for j in range(len(data_old.index)):
            if data_new['Обозначение документа'][i] == \
               data_old['Обозначение документа'][j]:
                match = True
                if data_new['№  изм.'][i] != data_old['№  изм.'][j]:
                    data_diff = pd.concat([data_diff,
                                           pd.DataFrame(data_new.loc[i]).T],
                                          ignore_index=True)
        # If there is no entry in the old list, then the document is new.
        if not match:
            data_diff = pd.concat([data_diff, pd.DataFrame(data_new.loc[i]).T],
                                  ignore_index=True)
    return data_diff


def frame_difference(frame_1: pd.DataFrame, frame_2: pd.DataFrame, keys: str):
    """Return frame_1 - frame_2."""
    dictionary = frame_to_dict(frame_2, keys)
    data_diff = pd.DataFrame(columns=frame_1.columns)
    for i in range(len(frame_1.index)):
        key = frame_1[keys][i]
        if key in dictionary and \
           frame_1['Дата ввода информации (фильтр новых поступлений)'][i] <= \
           dictionary[key]['Дата ввода информации (фильтр новых поступлений)']:
            continue
        else:
            data_diff = pd.concat([data_diff, pd.DataFrame(frame_1.loc[i]).T],
                                  ignore_index=True)
    return data_diff


def write_frame_to_excel(frame: pd.DataFrame, name: str):
    """Write data frame to excel file."""
    today = datetime.date.today()
    first_sheet = 'New as of ' + str(today)
    
    writer = pd.ExcelWriter(name, datetime_format='dd.mm.yyyy')
    frame.to_excel(writer, sheet_name=first_sheet, index=False)
    workbook = writer.book
    worksheet = writer.sheets[first_sheet]
    
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'border': 1,
        'valign': 'top',
        'font': 'Times New Roman',
        'size': '14'})
    cell_format_center = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'text_wrap': True,
        'font': 'Times New Roman',
        'size': '14'})
    cell_format_left = workbook.add_format({
        'align': 'left',
        'valign': 'vcenter',
        'text_wrap': True,
        'font': 'Times New Roman',
        'size': '14'})
    
    # Formatting header.
    for col_num, value in enumerate(frame.columns.values):
        worksheet.write(0, col_num, value, header_format)

    # Cell merging.
    columns = ['Город', 'Организация']
    for column in columns:
        identical_cells = find_identical_cells(frame, column)
        for cells in identical_cells:
            worksheet.merge_range(*cells, cell_format_center)

    # Format column.
    worksheet.set_column(0, 1, 30, cell_format_center)
    worksheet.set_column(2, 3, 30, cell_format_left)

    # Auto-adjust columns' width.
    for column in frame:
        column_width = max(frame[column].astype(str).map(len).max(),
                           len(column)) + 4
        col_idx = frame.columns.get_loc(column)
        worksheet.set_column(col_idx, col_idx, column_width)
        
    writer.save()


def find_identical_cells(df: pd.DataFrame, column: str):
    """Find identical cells in a column in a data frame.
    Return list of tuples of start cell, finish cell and value.
    """
    identical_cells = []
    col_idx = df.columns.get_loc(column)
    start = 0
    finish = 0
    value = None
    first_match = True
    group = False
    print(1, df[column][0])
    for i in range(1, len(df.index)):
        print(i+1, df[column][i])
        if df[column][i] == df[column][i-1] and first_match:
            value = df[column][i-1]
            start = i-1
            finish = i
            first_match = not first_match
            group = True
            continue
        elif df[column][i] == df[column][i-1] and not first_match:
            finish = i
            continue
        elif df[column][i] != df[column][i-1] and not first_match:
            cells = (start+1, col_idx, finish+1, col_idx, value)
            identical_cells.append(cells)
            first_match = not first_match
            group = False
        elif df[column][i] != df[column][i-1] and first_match:
            group = False
            continue
    if group:        
        cells = (start+1, col_idx, finish+1, col_idx, value)
        identical_cells.append(cells)
            
    for cells in identical_cells:
        print(cells)

    return identical_cells


def main():

    time_start = time.time()

    data_old = pd.read_excel(PATH_OLD, skiprows=[0])
    data_new = pd.read_excel(PATH_NEW, skiprows=[0])

    data_diff = frame_difference(data_new, data_old, 'Обозначение документа')
    
    columns = ['Город', 'Организация', 'Обозначение документа',
               'Наименование']
    data_new = pd.DataFrame()
    for column in columns:
        data_new[column] = data_diff[column]
##    print(data_new)

    # Write excel file.
    today = datetime.date.today()
    file_name = PATH_DIFF + '\\new receipts of DSO ' + str(today) + '.xlsx'   
    write_frame_to_excel(data_new, file_name)



##    # Test size of
##    data_test = pd.read_excel('test_file.xlsx')
##    print(data_test, end='\n\n')
##    print('size of data test', sys.getsizeof(data_test), end='\n\n')
##
##    dict_test = frame_to_dict(data_test, 'Обозначение')
##    print(dict_test, end='\n\n')
##    print('size of dict test', sys.getsizeof(dict_test), end='\n\n')


    time_finish = time.time()
    time_diff = time_finish - time_start
    print('\nComplete! Time:', '{:.2f}'.format(time_diff), 'sec.')


if __name__ == '__main__':
    main()

















