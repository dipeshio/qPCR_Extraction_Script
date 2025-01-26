from tkinter import filedialog
from tkinter import Tk
import pandas as pd
import glob
import os

def file_open():
    path, new_path = extraction_loc()
    excel_files = [path]
    n_files = []
    count = 0

    for f in excel_files:
        df = pd.concat(pd.read_excel(f, sheet_name=None, index_col=None, na_values=['NA']).values(), ignore_index=True)
        df = df[~df.applymap(lambda x: 'Std' in str(x)).any(axis=1)]
        df = df[~df.applymap(lambda x: 'Water' in str(x)).any(axis=1)]
        df = df[df.columns.drop(list(df.filter(regex='Standard')))]
        df = df[df.columns.drop(list(df.filter(regex='Replicate')))]

        if len(df.columns) == 8:
            df.drop(df.columns[7], axis=1, inplace=True)

        output_file = os.path.join(new_path, os.path.basename(f).replace('.xlsx', '_extracted.xlsx'))
        df.to_excel(output_file, sheet_name='Extracted Data', index=False)
        n_files.append(output_file)
        count += 1

    print(n_files)

def extraction_short(f):
    df = pd.concat(pd.read_excel(f, sheet_name=None, index_col=None, na_values=['NA'], usecols=[0,1,2,3,4,5,6]).values(), ignore_index=True)
    df = df[~df.applymap(lambda x: 'Std' in str(x)).any(axis=1)]
    df = df[~df.applymap(lambda x: 'Water' in str(x)).any(axis=1)]
    df = df[df.columns.drop(list(df.filter(regex='Standard')))]
    df = df[df.columns.drop(list(df.filter(regex='Replicate')))]

def extraction_loc():
    root = Tk()
    root.withdraw()
    path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    new_path = os.path.join(os.path.dirname(path), 'Extractions')

    if not os.path.exists(new_path):
        os.mkdir(new_path)
        print('Extracted excel files are located in the folder named Extractions')
    return path, new_path

file_open()