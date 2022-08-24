import pandas as pd
import datetime
import time
import sys


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
        dict_of_dict[key] = dict()
        for column in frame.columns:
            if column != keys:
                dict_of_dict[key][column] = frame[column][i]
    return dict_of_dict


def trivial_difference(data_new: pd.DataFrame, data_old: pd.DataFrame):
    """Asymptotic O(N^2)."""
    data_diff = pd.DataFrame(columns=data_new.columns)
    
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
    return data_diff


def frame_difference(frame1: pd.DataFrame, frame2: dict, keys: str):
    """Return frame1 - frame2."""
    data_diff = pd.DataFrame(columns=frame1.columns)
    for i in range(len(frame1.index)):
        key = frame1[keys][i]

        if key in frame2 and frame1['№  изм.'][i] == frame2[key]['№  изм.']:
            continue
        else:
            data_diff = pd.concat([data_diff, pd.DataFrame(frame1.loc[i]).T],
                                  ignore_index=True)
    return data_diff


def main():

    time_start = time.time()

    data_old = pd.read_excel(PATH_OLD, skiprows=[0])
    data_new = pd.read_excel(PATH_NEW, skiprows=[0])
    
##    data_diff = frame_difference(
##        data_new, frame_to_dict(data_old, 'Обозначение документа'),
##        keys='Обозначение документа')

    data_diff = trivial_difference(data_new, data_old)


    today = datetime.date.today()
    name_diff = PATH_DIFF + '\\new receipts of DSO ' + str(today) + '.xlsx'

    writer = pd.ExcelWriter(name_diff)
    data_diff.to_excel(writer, sheet_name='New as of ' + str(today),
                       index=False)
    # Auto-adjust columns' width
    for column in data_diff:
        column_width = max(data_diff[column].astype(str).map(len).max(),
                           len(column))
        col_idx = data_diff.columns.get_loc(column)
        writer.sheets['New as of ' + str(today)].set_column(col_idx, col_idx,
                                                            column_width)
    writer.save()



##    data_test = pd.read_excel('test_file.xlsx')
##    print(data_test, end='\n\n')
##    print('size of data test', sys.getsizeof(data_test), end='\n\n')
##
##    dict_test = frame_to_dict(data_test, 'Обозначение')
##    print(dict_test, end='\n\n')
##    print('size of dict test', sys.getsizeof(dict_test), end='\n\n')


    time_finish = time.time()
    time_diff = time_finish - time_start
    print('Complete! Time:', time_diff, 'секунд.')


if __name__ == '__main__':
    main()

















