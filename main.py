import sys
import time
import datetime
import pandas as pd


##PATH_OLD = '.\\archive of notifications DSO\\notifications DSO 2022-07-29.xlsx'
##PATH_NEW = '.\\central archive OTD\\notifications DSO.xlsx'
PATH_OLD = '.\\archive of notifications DSO\\test_file_old.xlsx'
PATH_NEW = '.\\central archive OTD\\test_file_new.xlsx'
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
    """Write frame to excel file."""
    today = datetime.date.today()
    writer = pd.ExcelWriter(name)
    frame.to_excel(writer, sheet_name='New as of ' + str(today), index=False)
    # Auto-adjust columns' width
    for column in frame:
        column_width = max(frame[column].astype(str).map(len).max(),
                           len(column)) + 5
        col_idx = frame.columns.get_loc(column)
        writer.sheets['New as of ' + str(today)].set_column(col_idx, col_idx,
                                                            column_width)
    writer.save()


def main():

    time_start = time.time()

    data_old = pd.read_excel(PATH_OLD, skiprows=[0])
    data_new = pd.read_excel(PATH_NEW, skiprows=[0])

    data_diff = frame_difference(data_new, data_old, 'Обозначение документа')

    # Write excel file.
    today = datetime.date.today()
    name_diff = PATH_DIFF + '\\new receipts of DSO ' + str(today) + '.xlsx'   
    write_frame_to_excel(data_diff, name_diff)



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

















