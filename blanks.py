import json
import general
import operator as op

def clearance(dataframe=None):
    source_dict = dataframe.to_dict(orient='list')
    openjson = open('json/clearance.json')
    destination_dict, len_source = json.load(openjson), 0

    # Collect length of value's list
    for key in source_dict:
        if len(source_dict[key]) > 0:
            len_source = len(source_dict[key])
            break

    # Add necessary new column(s)
    for key in destination_dict:
        if key not in source_dict:
            destination_dict[key] = ['--' * len_source]
        if key in source_dict:
            destination_dict[key] = source_dict[key]

    # Add a new column (Number) to destination
    destination_dict['NUMBER'] = list(range(1, len_source+1))

    # Mark keys in dictionary that contains ";"
    marked_keys = general.semicolon_list(destination_dict)

    # Using marked_keys to add new rows for ';' occurrence

