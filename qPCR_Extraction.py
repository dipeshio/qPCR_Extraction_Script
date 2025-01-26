from tkinter import filedialog
from tkinter import Tk
import pandas as pd
import glob
import os


def file_open():
    root = Tk()
    root.withdraw()
    path = filedialog.askdirectory()
    excel_files = glob.glob(os.path.join(path, "*.xlsx"))

    for f in excel_files:
        all_sheets = pd.read_excel(f, sheet_name=None, index_col=None, na_values=['NA'],
                                   usecols=[1, 2, 3, 10, 11, 12, 13, 14])
        filtered_sheets = []

        for sheet_name, df in all_sheets.items():
            df = df[~df.map(lambda x: 'Std' in str(x)).any(axis=1)]
            df = df[~df.map(lambda x: 'Water' in str(x)).any(axis=1)]
            df = df[df.columns.drop(list(df.filter(regex='Standard')))]
            df = df[df.columns.drop(list(df.filter(regex='Replicate')))]

            if len(df.columns) == 8:
                df.drop(df.columns[7], axis=1, inplace=True)

            filtered_sheets.append(df)

        combined_df = pd.concat(filtered_sheets, ignore_index=True)
        output_file = os.path.join(path, os.path.basename(f).replace('.xlsx', '_extracted.xlsx'))
        combined_df.to_excel(output_file, sheet_name='Extracted Data', index=False)


file_open()