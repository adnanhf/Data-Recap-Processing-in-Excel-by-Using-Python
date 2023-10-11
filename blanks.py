import json
import general


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

    # Add new row everytime found ";"
    for key in destination_dict:
        # Check if the entire list contain only string
        if general.contain_only(destination_dict[key], str):
            print('yes')
            print(destination_dict[key])
            # for item in destination_dict[key]:
                # if item.find('; '):
                    # print('Yes')
