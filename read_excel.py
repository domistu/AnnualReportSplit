import pandas as pd
import os
from get_file import get_file_name

def read_excel_file():

    # Column index for the room opening/closing data
    col_index = 1

    file_name = '0515-07 week (34).xlsx'
    df = pd.read_excel(file_name, sheet_name=None, header=None)

    dateRangeFile = 'dates.csv'
    dates_df = pd.read_csv(dateRangeFile, header=None)
    dates_lists = dates_df.values.tolist()

    # Get a list of sheets from the df dictionary
    sheetList = []

    for key in df.keys():
        sheetList.append(key)

    # Get the index of the end of the header portion for each sheet
    headerEndIndex = []

    for x in range(len(sheetList)):
        headerEndIndex.append(df[sheetList[x]].index[(df[sheetList[x]][col_index].str.contains('time', case=False) == True)][0])

    os.makedirs('Output Files', exist_ok=True)

    # Get filtered data for a date range
    for i in range(len(dates_lists)):

        startDateList = dates_lists[i][0].split('/')
        if len(startDateList[0]) < 2:
            startDateList[0] = f'0{startDateList[0]}'
        startDate = f'{startDateList[2]}-{startDateList[1]}-{startDateList[0]}'

        endDateList = dates_lists[i][6].split('/')
        if len(endDateList[0]) < 2:
            endDateList[0] = f'0{endDateList[0]}'
        endDate = f'{endDateList[2]}-{endDateList[1]}-{endDateList[0]}'

        outputFileName = f'Output Files/{startDate}_to_{endDate}.xlsx'

        writer = pd.ExcelWriter(outputFileName)

        for x in range(len(sheetList)):
            df_header = df[sheetList[x]].iloc[:headerEndIndex[x]+1]
            df_header.to_excel(writer, sheet_name=sheetList[x], index=False, header=None)
        
        writer.close()

        writer_append = pd.ExcelWriter(outputFileName, mode='a', if_sheet_exists='overlay')

        for x in range(len(sheetList)):
            dateFiltered_df = df[sheetList[x]].loc[(df[sheetList[x]][col_index].str.match('|'.join(dates_lists[i]), case=False) == True)]
            dateFiltered_df.to_excel(writer_append, sheet_name=sheetList[x], index=False, header=None, startrow=headerEndIndex[x]+1)       
        
        writer_append.close()

read_excel_file()