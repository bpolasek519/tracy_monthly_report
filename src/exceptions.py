import pandas as pd
import os


def read_excel_with_exception(filename, sheet_name):
    try:
        df = pd.read_excel(filename, sheet_name=sheet_name, engine='openpyxl')
        test = df.dtypes
        return df
    except Exception as e:
        file = os.path.basename(filename)
        raise Exception(f'Unable to read {file}. Please check your file and make adjustments as needed. '
                        f'(Most common error is #DIV/0! error.)')
