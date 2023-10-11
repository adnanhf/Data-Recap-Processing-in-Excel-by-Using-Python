import general
import blanks
from pandas import DataFrame


def main():
    data_from_excel: DataFrame = general.reader(filename='files/File Input.xlsx')
    # with pd.ExcelWriter('files/File Output.xlsx', engine='xlsxwriter') as writer:
    blanks.clearance(data_from_excel)


if __name__ == '__main__':
    main()
