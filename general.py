from typing import List, Any

import pandas as pd
import operator as op
from pandas import DataFrame


def reader(filename='None'):
    dataframe_name = pd.read_excel(filename)
    del dataframe_name['NO']
    dataframe_name = sorter(dataframe_name, ['TANGGAL BERTOLAK', 'JAM BERTOLAK'])

    return dataframe_name


def sorter(dataframe=None, sorted_by=None):
    dataframe.sort_values(by=sorted_by, inplace=True)

    return dataframe


def contain(data=None, type_name=None):
    return all(isinstance(item, type_name) for item in data) or any(isinstance(item, type_name) for item in data)


def classifier(dataframe=None, column_name=None, value=None, model: str = None):
    model.lower()
    if model == 'ssib':
        filtered: DataFrame = dataframe.loc[dataframe[column_name] <= value]
        filtered.reset_index(drop=True, inplace=True)

        return filtered

    elif model == 'bsib':
        filtered: DataFrame = dataframe.loc[dataframe[column_name] > value]
        filtered.reset_index(drop=True, inplace=True)

        return filtered

    elif model == 'other':
        filtered: DataFrame = dataframe.loc[dataframe[column_name] == value]
        filtered.reset_index(drop=True, inplace=True)

        return filtered

    else:
        print('Specify correct data label!')


def semicolon_list(dict_data):
    result: list[Any] = []
    for key in dict_data:
        if contain(dict_data[key], str):
            for item in dict_data[key]:
                try:
                    op.contains(item, '; ')
                except TypeError:
                    item = str(item)
                finally:
                    if op.contains(item, '; '):
                        result.append(key)
                        break
                break

    return result
