
import pandas as pd
import os
import sys
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from matplotlib.ticker import StrMethodFormatter

pd.set_option('display.float_format', lambda x: '%.1f' % x)

def merge_sheets():
    while True:
        try:
            path = input('Please enter the path for the folder that contains the sheets to be merged: \n')
            path = path.replace('"', '')
            if not os.path.isdir:
                print('Invalid Path, Please Try Again!')
                continue
            break
        except:
            print('Invalid Path, Please Try Again!')

    df = pd.DataFrame()
    try:
        files = os.listdir(path)
        filename = ''
        for file in files:
            # skipping opened files
            if '~$' in file: continue
            if file[-3:] == 'csv':
                filename = file
                # handling csv format
                df_file = pd.read_csv(path + "\\" + file, encoding = "ISO-8859-1")
                extension = '.csv'
                df = df.append(df_file)
            elif file[-4:] == 'xlsx':
                filename = file
                # handling Excel format
                df_file = pd.read_excel(path + "\\" + file, engine='openpyxl')
                extension = '.xlsx'
                df = df.append(df_file)

        if extension == '.csv':
            name = path + "\\" + filename[:-4] + '_merged' + extension
            df.to_csv(name, encoding='UTF-8', index=False)
        else:
            name = path+ "\\" + filename[:-5] + '_merged' + extension
            df.to_excel(name, index=False)
    except Exception as err:
        print('The below error occurred during the processing of the sheet:')
        print(err)
        input()
        sys.exit(1)

def split_sheet():
    while True:
        n = input('Please enter the number of sub files to be generated: ')
        try:
            n = int(n)
        except:
            print('Invalid Input, Please Try Again!')
            continue
        if n > 0:
            break
        else:
            print('Invalid Input, Please Try Again!')

    while True:
        try:
            path = input('Please enter the file path: \n')
            path = path.replace('"', '')
            if not os.path.isfile:
                print('Invalid Input, Please Try Again!')
                continue
            if path[-3:] != 'csv' and path[-3:] != 'lsx':
                print('Path must be for XLSX or CSV files')      
            else:
                break
        except:
            print('Invalid Input, Please Try Again!')


    try:
        if path[-3:] == 'csv':
            # handling csv format
            df = pd.read_csv(path, encoding = "ISO-8859-1")
            extension = '.csv'
        else:
            # handling Excel format
            df = pd.read_excel(path, engine='openpyxl')
            extension = '.xlsx'
    except Exception as err:
        print('The below error occurred during the processing of the sheet:')
        print(err)
        sys.exit(1)

    nrows = df.shape[0]
    strt, end = 0, 0
    step = int(nrows/n)
    if step == 0:
        step = 1
    for i in range(n):
        if i > nrows: break
        if end == 0:
            strt = 0
            end = step
        elif i == n - 1:
            strt = end
            end = nrows
        else:
            strt = end
            end = end + step
        df2 = df.iloc[strt:end, :]

        if extension == '.csv':
            name = path[:-4] + f'_{i+1}' + extension
            df2.to_csv(name, encoding='UTF-8', index=False)
        else:
            name = path[:-5] + f'_{i+1}' + extension
            df2.to_excel(name, index=False)
            
def process_sheet():

    while True:
        try:
            path = input('Please enter the file path: \n')
            path = path.replace('"', '')
            if not os.path.isfile:
                print('Invalid Input, Please Try Again!')
                continue
            if path[-3:] != 'csv' and path[-3:] != 'lsx':
                print('Path must be for XLSX or CSV files')      
            else:
                break
        except:
            print('Invalid Input, Please Try Again!')

    print('-'*110)
    print('Processing the sheet ...')
    print('-'*110)
    try:
        if path[-3:] == 'csv':
            # handling csv format
            df = pd.read_csv(path, encoding = "ISO-8859-1", low_memory=False)
        else:
            # handling Excel format
            df = pd.read_excel(path, engine='openpyxl', low_memory=False)
    except Exception as err:
        print('The below error occurred during the processing of the sheet:')
        print(err)
        sys.exit(1)

    for col in df.columns:
        df[col] = df[col].astype(float)

    print(df.info(verbose=True, show_counts=False))

    for col in df.columns:
        if 'Coefficient' in col:
            col1 = col
        elif 'Temperature' in col:
            col2 = col       
            col2 = col2.replace('(K)', '(C)')           
            df[col] = df[col].apply(lambda x: x-273)
            df = df.rename(columns={col: col2})

    print('-'*110)
    print('Descriptive statistics')
    print('-'*110)
    print(df[[col1, col2]].describe(percentiles=[]))
    print('-'*110)

    fig, (ax1, ax2) = plt.subplots(1,2, figsize=(14,5))
    sns.histplot(df[col1], ax=ax1, bins=50, log_scale=True)
    ax1.get_xaxis().set_major_formatter(StrMethodFormatter('{x:.3f}'))
    sns.histplot(df[col2],ax=ax2, bins=50)
    plt.show()

if __name__ == "__main__":

    while True:
        print("Please select the required operation\n")
        print("1. Merging Sheets")
        print("2. Splitting Sheet")
        print("3. Processing sheet\n")
        func = input()
        try:
            func = int(func)
        except:
            print("Invalid Input, please input a number from 1-3")
            continue
        if func != 1 and func != 2 and func != 3:
            print("Invalid Input, please input a number from 1-3")
            continue
        break

    try:
        if func == 1:
            merge_sheets()
        elif func == 2:
            split_sheet()
        else:
            process_sheet()
        input('Process completed successfully, press any key to exit.')
    except Exception as err:
        print('The following error occurred:')
        print(err)
        input('Press any key to exit.')



