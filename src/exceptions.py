import pandas as pd
import os


def read_excel_with_exception(filename, sheet_name):
    try:
        df = pd.read_excel(filename, sheet_name=sheet_name, engine='openpyxl')
        test = df.dtypes
        return df
    except Exception as e:
        file = os.path.basename(filename)
        raise Exception(f'Unable to read {file}. Check that filters are turned off and there are no #DIV/0!')
