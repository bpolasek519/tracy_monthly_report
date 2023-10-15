import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side, PatternFill, Font
from openpyxl.utils import get_column_letter
from src.constants import (CENTER_STYLE, PERCENTAGE_STYLE, DATE_STYLE, ACCOUNTING_STYLE, COLUMN_WIDTHS,
                           ACCOUNTING_STYLE_NO_BORDER, NUMBER_STYLE, STYLE_MAPPINGS)
import src.dataframe_helper as dh


def create_llc_wip_sheet(df: pd.DataFrame, wb: Workbook) -> None:
    final_df = dh.create_wip_df(df=df, title='LLC USPS WIP', for_fs=False)

    ws = wb.create_sheet(title='LLC WIP')

    center_algined_cols = ['Type: \nJOC, HB', 'Contract', 'Proj. #', 'Prob. \n C/O #']

    # Setting a specified height to all the cells
    ws.sheet_format.defaultRowHeight = 30
    ws.sheet_format.customHeight = True

    # Applying styles based on the column
    for col, col_name in enumerate(final_df.columns, start=1):
        cell = ws.cell(row=1, column=col, value=col_name)
        cell.style = CENTER_STYLE
        cell.alignment = cell.alignment.copy(wrapText=True)

    # Apply the styling to the columns
    col_names = final_df.columns
    col_indices = {}
    for col_name, style in STYLE_MAPPINGS.items():
        if col_name in col_names:
            col_idx = col_names.get_loc(col_name)
            col_indices[col_name] = col_idx
            if col_name == 'Comment':
                set_style(df=final_df, ws=ws, col_name=col_name, style=style, idx=col_idx, wrapText=True)
            else:
                set_style(df=final_df, ws=ws, col_name=col_name, style=style, idx=col_idx)

    # This adds the accounting styling for the total sum row without adding anything else.
    last_row = final_df.iloc[-1]

    last_row_style_cols = ['Balance WIP', 'Awd $', 'Bill $', '$ Previously Paid', '$ Outstanding']
    for col in last_row_style_cols:
        col_idx = final_df.columns.get_loc(col)
        value = last_row[col]
        cell = ws.cell(row=len(final_df) + 1, column=col_idx + 1, value=value)
        cell.style = ACCOUNTING_STYLE_NO_BORDER

    for col_name in center_algined_cols:
        col_idx = final_df.columns.get_loc(col_name) + 1
        for row, value in enumerate(final_df[col_name][:-1], start=2):
            cell = ws.cell(row=row, column=col_idx, value=value)
            cell.style = CENTER_STYLE

    # This is the general styling for the cells
    for row, row_data in enumerate(final_df.itertuples(index=False), start=2):
        if row <= len(final_df):
            for col, value in enumerate(row_data[:-1], start=1):
                if (col != col_indices['Awd'] + 1 and col != col_indices['Substantial Complete'] + 1 and
                        col != col_indices['Billed Date'] + 1 and col != col_indices['Awd $'] + 1 and
                        col != col_indices['Bill $'] + 1 and col != col_indices['Contract Comp. Date'] + 1
                        and col != col_indices['$ Previously Paid'] + 1 and col != col_indices['%'] + 1 and
                        col != col_indices['$ Outstanding'] + 1 and col != col_indices['Balance WIP'] + 1):
                    cell = ws.cell(row=row, column=col, value=value)
                    cell.border = Border(left=Side(border_style='thin'),
                                         right=Side(border_style='thin'),
                                         top=Side(border_style='thin'),
                                         bottom=Side(border_style='thin'))
        elif row == len(final_df) + 1:
            for col, value in enumerate(row_data, start=1):
                cell = ws.cell(row=row, column=col, value=value)

    apply_column_widths(final_df, ws)
    header_settings(final_df, ws)
    alternate_color_fill(final_df, ws)

    wb.save('data/test.xlsx')


def create_llc_outstanding_sheet(df: pd.DataFrame, wb: Workbook) -> None:
    final_df = dh.create_outstanding_df(df, title='LLC USPS Outstanding')

    ws = wb.create_sheet(title='LLC Outstanding')

    center_algined_cols = ['Type: \nJOC, HB', 'Contract', 'Proj. #', 'Prob. \n C/O #']

    # Setting a specified height to all the cells
    ws.sheet_format.defaultRowHeight = 30
    ws.sheet_format.customHeight = True

    # Applying styles based on the column
    for col, col_name in enumerate(final_df.columns, start=1):
        cell = ws.cell(row=1, column=col, value=col_name)
        cell.style = CENTER_STYLE
        cell.alignment = cell.alignment.copy(wrapText=True)

    # Apply the styling to the columns
    col_names = final_df.columns
    col_indices = {}
    for col_name, style in STYLE_MAPPINGS.items():
        if col_name in col_names:
            col_idx = col_names.get_loc(col_name)
            col_indices[col_name] = col_idx
            if col_name == 'Comment':
                set_style(df=final_df, ws=ws, col_name=col_name, style=style, idx=col_idx, wrapText=True)
            else:
                set_style(df=final_df, ws=ws, col_name=col_name, style=style, idx=col_idx)

    # This adds the accounting styling for the total sum row without adding anything else.
    last_row = final_df.iloc[-1]

    last_row_style_cols = ['Balance Due', 'Awd $', 'Bill $', '$ Paid']
    for col in last_row_style_cols:
        col_idx = final_df.columns.get_loc(col)
        value = last_row[col]
        cell = ws.cell(row=len(final_df) + 1, column=col_idx + 1, value=value)
        cell.style = ACCOUNTING_STYLE_NO_BORDER

    for col_name in center_algined_cols:
        col_idx = final_df.columns.get_loc(col_name) + 1
        for row, value in enumerate(final_df[col_name][:-1], start=2):
            cell = ws.cell(row=row, column=col_idx, value=value)
            cell.style = CENTER_STYLE

    # This is the general styling for the cells
    for row, row_data in enumerate(final_df.itertuples(index=False), start=2):
        if row <= len(final_df):
            for col, value in enumerate(row_data[:-1], start=1):
                if (col != col_indices['Awd'] + 1 and col != col_indices['Substantial Complete'] + 1 and
                        col != col_indices['Billed Date'] + 1 and col != col_indices['Awd $'] + 1 and
                        col != col_indices['Bill $'] + 1 and col != col_indices['$ Paid'] + 1
                        and col != col_indices['Balance Due'] + 1 and col != col_indices['%'] + 1 and
                        col != col_indices['Comment'] + 1):
                    cell = ws.cell(row=row, column=col, value=value)
                    cell.border = Border(left=Side(border_style='thin'),
                                         right=Side(border_style='thin'),
                                         top=Side(border_style='thin'),
                                         bottom=Side(border_style='thin'))
        elif row == len(final_df) + 1:
            for col, value in enumerate(row_data, start=1):
                cell = ws.cell(row=row, column=col, value=value)

    apply_column_widths(final_df, ws)
    header_settings(final_df, ws)
    alternate_color_fill(final_df, ws)

    wb.save('data/test.xlsx')


def create_llc_paid_sheet(df: pd.DataFrame, wb: Workbook) -> None:
    final_df = dh.create_paid_df(df, 'LLC USPS Paid')

    ws = wb.create_sheet(title='LLC Paid')

    center_algined_cols = ['Type: \nJOC, HB', 'Contract', 'Proj. #', 'Prob. \n C/O #']

    # Setting a specified height to all the cells
    ws.sheet_format.defaultRowHeight = 30
    ws.sheet_format.customHeight = True

    # Applying styles based on the column
    for col, col_name in enumerate(final_df.columns, start=1):
        cell = ws.cell(row=1, column=col, value=col_name)
        cell.style = CENTER_STYLE
        cell.alignment = cell.alignment.copy(wrapText=True)

    # Apply the styling to the columns
    col_names = final_df.columns
    col_indices = {}
    for col_name, style in STYLE_MAPPINGS.items():
        if col_name in col_names:
            col_idx = col_names.get_loc(col_name)
            col_indices[col_name] = col_idx
            if col_name == 'Comment':
                set_style(df=final_df, ws=ws, col_name=col_name, style=style, idx=col_idx, wrapText=True)
            else:
                set_style(df=final_df, ws=ws, col_name=col_name, style=style, idx=col_idx)

    # This adds the accounting styling for the total sum row without adding anything else.
    last_row = final_df.iloc[-1]

    last_row_style_cols = ['Balance Due', 'Awd $', 'Bill $', '$ Previously Paid', '$ Paid']
    for col in last_row_style_cols:
        col_idx = final_df.columns.get_loc(col)
        value = last_row[col]
        cell = ws.cell(row=len(final_df) + 1, column=col_idx + 1, value=value)
        cell.style = ACCOUNTING_STYLE_NO_BORDER

    for col_name in center_algined_cols:
        col_idx = final_df.columns.get_loc(col_name) + 1
        for row, value in enumerate(final_df[col_name][:-1], start=2):
            cell = ws.cell(row=row, column=col_idx, value=value)
            cell.style = CENTER_STYLE

    # This is the general styling for the cells
    for row, row_data in enumerate(final_df.itertuples(index=False), start=2):
        if row <= len(final_df):
            for col, value in enumerate(row_data[:-1], start=1):
                if (col != col_indices['Awd'] + 1 and col != col_indices['Substantial Complete'] + 1 and
                        col != col_indices['Billed Date'] + 1 and col != col_indices['Paid/Closed'] + 1 and
                        col != col_indices['Awd $'] + 1 and col != col_indices['Bill $'] + 1
                        and col != col_indices['$ Paid'] + 1 and col != col_indices['$ Previously Paid'] + 1 and
                        col != col_indices['Balance Due'] + 1 and col != col_indices['%'] + 1 and
                        col != col_indices['Comment'] + 1):
                    cell = ws.cell(row=row, column=col, value=value)
                    cell.border = Border(left=Side(border_style='thin'),
                                         right=Side(border_style='thin'),
                                         top=Side(border_style='thin'),
                                         bottom=Side(border_style='thin'))
        elif row == len(final_df) + 1:
            for col, value in enumerate(row_data, start=1):
                cell = ws.cell(row=row, column=col, value=value)

    apply_column_widths(final_df, ws)
    header_settings(final_df, ws)
    alternate_color_fill(final_df, ws)

    wb.remove_sheet(wb.get_sheet_by_name('Sheet'))
    wb.save('data/test.xlsx')


def create_fs_paid_sheet(usps_df: pd.DataFrame, hcde_df: pd.DataFrame, misc_df: pd.DataFrame, buyboard_df: pd.DataFrame,
                         pca_df: pd.DataFrame, friendswood_df: pd.DataFrame, wb: Workbook) -> None:
    # Retrieving the Excel sheets needed and matching the column names and creating the paid tables
    final_buyboard_df, final_friendswood_df, final_hcde_df, final_misc_df, final_pca_df, final_usps_df = (
        clean_col_names(buyboard_df, friendswood_df, hcde_df, misc_df, pca_df, usps_df, func=dh.create_paid_df,
                        name='Paid'))

    # Create a blank row DataFrame with NaN values
    blank_row = pd.DataFrame({col: [np.nan] for col in final_usps_df.columns})

    # Merge the multiple dataframes and note where exlcuded rows are for styling. This includes the 'total' row,
    # the blank row, and the next header row.
    df_lists = [final_usps_df, final_misc_df, final_buyboard_df, final_hcde_df, final_pca_df, final_friendswood_df]
    style_skipped_rows = []
    total_rows = []
    final_df = pd.DataFrame()
    for df in df_lists:
        if not df.empty:
            if final_df.empty:
                final_df = df
                style_skipped_rows.append(len(final_df.index) - 1)
                total_rows.append(len(final_df.index) - 1)
                style_skipped_rows.append(len(final_df.index))
                style_skipped_rows.append(len(final_df.index) + 1)
            else:
                final_df = pd.concat([final_df, blank_row, df], ignore_index=True)
                style_skipped_rows.append(len(final_df.index) - 1)
                total_rows.append(len(final_df.index) - 1)
                style_skipped_rows.append(len(final_df.index))
                style_skipped_rows.append(len(final_df.index) + 1)
        else:
            continue

    total_df = final_df.iloc[total_rows]

    # Adds the total amount of money
    awd_sum = total_df['Awd $'].sum()
    bill_sum = total_df['Bill $'].sum()
    prev_sum = total_df['$ Previously Paid'].sum()
    paid_sum = total_df['$ Paid'].sum()
    balance_due_sum = total_df['Balance Due'].sum()
    total_row = {'Type: \nJOC, HB': f'Total Paid July 2023', 'Awd $': awd_sum, 'Bill $': bill_sum,
                 '$ Previously Paid': prev_sum, '$ Paid': paid_sum, 'Balance Due': balance_due_sum}
    total_row_df = pd.DataFrame(total_row, index=[0])
    final_df = pd.concat([final_df, blank_row, total_row_df], ignore_index=True)

    final_df.reset_index(drop=True, inplace=True)

    ws = wb.create_sheet(title='FS Paid')

    center_algined_cols = ['Type: \nJOC, HB', 'Contract', 'Proj. #', 'Prob. \n C/O #']

    # Setting a specified height to all the cells
    ws.sheet_format.defaultRowHeight = 30
    ws.sheet_format.customHeight = True

    # Centers the header row
    for col, col_name in enumerate(final_df.columns, start=1):
        cell = ws.cell(row=1, column=col, value=col_name)
        cell.style = CENTER_STYLE
        cell.alignment = cell.alignment.copy(wrapText=True)

    # Apply the styling to the columns
    col_names = final_df.columns
    col_indices = {}
    for col_name, style in STYLE_MAPPINGS.items():
        if col_name in col_names:
            col_idx = col_names.get_loc(col_name)
            col_indices[col_name] = col_idx
            if col_name == 'Comment':
                set_style(df=final_df, ws=ws, col_name=col_name, style=style, idx=col_idx, wrapText=True)
            else:
                set_style(df=final_df, ws=ws, col_name=col_name, style=style, idx=col_idx)

    last_row_style_cols = ['Balance Due', 'Awd $', 'Bill $', '$ Previously Paid', '$ Paid']
    for col in last_row_style_cols:
        col_idx = final_df.columns.get_loc(col)
        for row in style_skipped_rows[:-2]:
            value = final_df.iloc[row, col_idx]
            cell = ws.cell(row=row + 2, column=col_idx + 1, value=value)
            cell.style = ACCOUNTING_STYLE_NO_BORDER
        value = final_df.iloc[-1][col]
        cell = ws.cell(row=len(final_df) + 1, column=col_idx + 1, value=value)
        cell.style = ACCOUNTING_STYLE_NO_BORDER

    for col_name in center_algined_cols:
        col_idx = final_df.columns.get_loc(col_name) + 1
        for row, value in enumerate(final_df[col_name][:-1], start=2):
            cell = ws.cell(row=row, column=col_idx, value=value)
            cell.style = CENTER_STYLE

    # Need to shift the rows down 2 due to the way the Excel sheet locates them.
    shifted_style_skipped_rows = [x + 2 for x in style_skipped_rows]
    # This is the general styling for the cells
    for row, row_data in enumerate(final_df.itertuples(index=False), start=2):
        if row not in shifted_style_skipped_rows:
            for col, value in enumerate(row_data, start=1):
                if (col != col_indices['Awd'] + 1 and col != col_indices['Substantial Complete'] + 1 and
                        col != col_indices['Billed Date'] + 1 and col != col_indices['Paid/Closed'] + 1 and
                        col != col_indices['Awd $'] + 1 and col != col_indices['Bill $'] + 1
                        and col != col_indices['$ Paid'] + 1 and col != col_indices['$ Previously Paid'] + 1 and
                        col != col_indices['Balance Due'] + 1 and col != col_indices['%'] + 1 and
                        col != col_indices['Comment'] + 1):
                    cell = ws.cell(row=row, column=col, value=value)
                    cell.border = Border(left=Side(border_style='thin'),
                                         right=Side(border_style='thin'),
                                         top=Side(border_style='thin'),
                                         bottom=Side(border_style='thin'))
        else:
            for col, value in enumerate(row_data, start=1):
                cell = ws.cell(row=row, column=col, value=value)

    apply_column_widths(final_df, ws)
    header_settings(final_df, ws, shifted_style_skipped_rows)
    alternate_color_fill(final_df, ws, shifted_style_skipped_rows)

    wb.save('data/test.xlsx')


def create_fs_outstanding_sheet(usps_df: pd.DataFrame, hcde_df: pd.DataFrame, misc_df: pd.DataFrame,
                                buyboard_df: pd.DataFrame, pca_df: pd.DataFrame, friendswood_df: pd.DataFrame,
                                wb: Workbook) -> None:
    # Retrieving the Excel sheets needed and matching the column names and creating the paid tables
    final_buyboard_df, final_friendswood_df, final_hcde_df, final_misc_df, final_pca_df, final_usps_df = (
        clean_col_names(buyboard_df, friendswood_df, hcde_df, misc_df, pca_df, usps_df, func=dh.create_outstanding_df,
                        name='Outstanding'))

    # Create a blank row DataFrame with NaN values
    blank_row = pd.DataFrame({col: [np.nan] for col in final_usps_df.columns})

    # Merge the multiple dataframes and note where exlcuded rows are for styling. This includes the 'total' row,
    # the blank row, and the next header row.
    df_lists = [final_usps_df, final_misc_df, final_buyboard_df, final_hcde_df, final_pca_df, final_friendswood_df]
    style_skipped_rows = []
    total_rows = []
    final_df = pd.DataFrame()
    for df in df_lists:
        if not df.empty:
            if final_df.empty:
                final_df = df
                style_skipped_rows.append(len(final_df.index) - 1)
                total_rows.append(len(final_df.index) - 1)
                style_skipped_rows.append(len(final_df.index))
                style_skipped_rows.append(len(final_df.index) + 1)
            else:
                final_df = pd.concat([final_df, blank_row, df], ignore_index=True)
                style_skipped_rows.append(len(final_df.index) - 1)
                total_rows.append(len(final_df.index) - 1)
                style_skipped_rows.append(len(final_df.index))
                style_skipped_rows.append(len(final_df.index) + 1)
        else:
            continue

    total_df = final_df.iloc[total_rows]

    # Adds the total amount of money
    awd_sum = total_df['Awd $'].sum()
    bill_sum = total_df['Bill $'].sum()
    paid_sum = total_df['$ Paid'].sum()
    balance_due_sum = total_df['Balance Due'].sum()
    total_row = {'Type: \nJOC, HB': f'Total Outstanding July 2023', 'Awd $': awd_sum, 'Bill $': bill_sum,
                 '$ Paid': paid_sum, 'Balance Due': balance_due_sum}
    total_row_df = pd.DataFrame(total_row, index=[0])
    final_df = pd.concat([final_df, blank_row, total_row_df], ignore_index=True)

    final_df.reset_index(drop=True, inplace=True)

    ws = wb.create_sheet(title='FS Outstanding')

    center_algined_cols = ['Type: \nJOC, HB', 'Contract', 'Proj. #', 'Prob. \n C/O #']

    # Setting a specified height to all the cells
    ws.sheet_format.defaultRowHeight = 30
    ws.sheet_format.customHeight = True

    # Centers the header row
    for col, col_name in enumerate(final_df.columns, start=1):
        cell = ws.cell(row=1, column=col, value=col_name)
        cell.style = CENTER_STYLE
        cell.alignment = cell.alignment.copy(wrapText=True)

    # Apply the styling to the columns
    col_names = final_df.columns
    col_indices = {}
    for col_name, style in STYLE_MAPPINGS.items():
        if col_name in col_names:
            col_idx = col_names.get_loc(col_name)
            col_indices[col_name] = col_idx
            if col_name == 'Comment':
                set_style(df=final_df, ws=ws, col_name=col_name, style=style, idx=col_idx, wrapText=True)
            else:
                set_style(df=final_df, ws=ws, col_name=col_name, style=style, idx=col_idx)

    last_row_style_cols = ['Balance Due', 'Awd $', 'Bill $', '$ Paid']
    for col in last_row_style_cols:
        col_idx = final_df.columns.get_loc(col)
        for row in style_skipped_rows[:-2]:
            value = final_df.iloc[row, col_idx]
            cell = ws.cell(row=row + 2, column=col_idx + 1, value=value)
            cell.style = ACCOUNTING_STYLE_NO_BORDER
        value = final_df.iloc[-1][col]
        cell = ws.cell(row=len(final_df) + 1, column=col_idx + 1, value=value)
        cell.style = ACCOUNTING_STYLE_NO_BORDER

    for col_name in center_algined_cols:
        col_idx = final_df.columns.get_loc(col_name) + 1
        for row, value in enumerate(final_df[col_name][:-1], start=2):
            cell = ws.cell(row=row, column=col_idx, value=value)
            cell.style = CENTER_STYLE

    # Need to shift the rows down 2 due to the way the Excel sheet locates them.
    shifted_style_skipped_rows = [x + 2 for x in style_skipped_rows]
    # This is the general styling for the cells
    for row, row_data in enumerate(final_df.itertuples(index=False), start=2):
        if row not in shifted_style_skipped_rows:
            for col, value in enumerate(row_data, start=1):
                if (col != col_indices['Awd'] + 1 and col != col_indices['Substantial Complete'] + 1 and
                        col != col_indices['Billed Date'] + 1 and col != col_indices['Awd $'] + 1 and
                        col != col_indices['Bill $'] + 1 and col != col_indices['$ Paid'] + 1
                        and col != col_indices['Balance Due'] + 1 and col != col_indices['%'] + 1 and
                        col != col_indices['Comment'] + 1):
                    cell = ws.cell(row=row, column=col, value=value)
                    cell.border = Border(left=Side(border_style='thin'),
                                         right=Side(border_style='thin'),
                                         top=Side(border_style='thin'),
                                         bottom=Side(border_style='thin'))
        else:
            for col, value in enumerate(row_data, start=1):
                cell = ws.cell(row=row, column=col, value=value)

    apply_column_widths(final_df, ws)
    header_settings(final_df, ws, shifted_style_skipped_rows)
    alternate_color_fill(final_df, ws, shifted_style_skipped_rows)

    wb.save('data/test.xlsx')


def create_fs_wip_sheet(usps_df: pd.DataFrame, hcde_df: pd.DataFrame, misc_df: pd.DataFrame, buyboard_df: pd.DataFrame,
                        pca_df: pd.DataFrame, friendswood_df: pd.DataFrame, wb: Workbook) -> None:
    # Retrieving the Excel sheets needed and matching the column names and creating the paid tables
    final_buyboard_df, final_friendswood_df, final_hcde_df, final_misc_df, final_pca_df, final_usps_df = (
        clean_col_names(buyboard_df, friendswood_df, hcde_df, misc_df, pca_df, usps_df, func=dh.create_wip_df,
                        name='WIP'))

    # Create a blank row DataFrame with NaN values
    blank_row = pd.DataFrame({col: [np.nan] for col in final_usps_df.columns})

    # Merge the multiple dataframes and note where exlcuded rows are for styling. This includes the 'total' row,
    # the blank row, and the next header row.
    df_lists = [final_usps_df, final_misc_df, final_buyboard_df, final_hcde_df, final_pca_df, final_friendswood_df]
    style_skipped_rows = []
    total_rows = []
    final_df = pd.DataFrame()
    for df in df_lists:
        if not df.empty:
            if final_df.empty:
                final_df = df
                style_skipped_rows.append(len(final_df.index) - 1)
                total_rows.append(len(final_df.index) - 1)
                style_skipped_rows.append(len(final_df.index))
                style_skipped_rows.append(len(final_df.index) + 1)
            else:
                final_df = pd.concat([final_df, blank_row, df], ignore_index=True)
                style_skipped_rows.append(len(final_df.index) - 1)
                total_rows.append(len(final_df.index) - 1)
                style_skipped_rows.append(len(final_df.index))
                style_skipped_rows.append(len(final_df.index) + 1)
        else:
            continue

    total_df = final_df.iloc[total_rows]

    # Adds the total amount of money
    awd_sum = total_df['Awd $'].sum()
    bill_sum = total_df['Bill $'].sum()
    paid_sum = total_df['$ Previously Paid'].sum()
    outstanding_sum = total_df['$ Outstanding'].sum()
    balance_due_sum = total_df['Balance WIP'].sum()
    total_row = {'Type: \nJOC, HB': f'Total WIP July 2023', 'Awd $': awd_sum, 'Bill $': bill_sum,
                 '$ Previously Paid': paid_sum, '$ Outstanding': outstanding_sum, 'Balance WIP': balance_due_sum}
    total_row_df = pd.DataFrame(total_row, index=[0])
    final_df = pd.concat([final_df, blank_row, total_row_df], ignore_index=True)

    final_df.reset_index(drop=True, inplace=True)

    ws = wb.create_sheet(title='FS WIP')

    center_algined_cols = ['Type: \nJOC, HB', 'Contract', 'Proj. #', 'Prob. \n C/O #']

    # Setting a specified height to all the cells
    ws.sheet_format.defaultRowHeight = 30
    ws.sheet_format.customHeight = True

    # Centers the header row
    for col, col_name in enumerate(final_df.columns, start=1):
        cell = ws.cell(row=1, column=col, value=col_name)
        cell.style = CENTER_STYLE
        cell.alignment = cell.alignment.copy(wrapText=True)

    # Apply the styling to the columns
    col_names = final_df.columns
    col_indices = {}
    for col_name, style in STYLE_MAPPINGS.items():
        if col_name in col_names:
            col_idx = col_names.get_loc(col_name)
            col_indices[col_name] = col_idx
            if col_name == 'Comment':
                set_style(df=final_df, ws=ws, col_name=col_name, style=style, idx=col_idx, wrapText=True)
            else:
                set_style(df=final_df, ws=ws, col_name=col_name, style=style, idx=col_idx)

    last_row_style_cols = ['Balance WIP', 'Awd $', 'Bill $', '$ Previously Paid', '$ Outstanding']
    for col in last_row_style_cols:
        col_idx = final_df.columns.get_loc(col)
        for row in style_skipped_rows[:-2]:
            value = final_df.iloc[row, col_idx]
            cell = ws.cell(row=row + 2, column=col_idx + 1, value=value)
            cell.style = ACCOUNTING_STYLE_NO_BORDER
        value = final_df.iloc[-1][col]
        cell = ws.cell(row=len(final_df) + 1, column=col_idx + 1, value=value)
        cell.style = ACCOUNTING_STYLE_NO_BORDER

    for col_name in center_algined_cols:
        col_idx = final_df.columns.get_loc(col_name) + 1
        for row, value in enumerate(final_df[col_name][:-1], start=2):
            cell = ws.cell(row=row, column=col_idx, value=value)
            cell.style = CENTER_STYLE

    # Need to shift the rows down 2 due to the way the Excel sheet locates them.
    shifted_style_skipped_rows = [x + 2 for x in style_skipped_rows]
    # This is the general styling for the cells
    for row, row_data in enumerate(final_df.itertuples(index=False), start=2):
        if row not in shifted_style_skipped_rows:
            for col, value in enumerate(row_data, start=1):
                if (col != col_indices['Awd'] + 1 and col != col_indices['Substantial Complete'] + 1 and
                        col != col_indices['Billed Date'] + 1 and col != col_indices['Awd $'] + 1 and
                        col != col_indices['Bill $'] + 1 and col != col_indices['$ Previously Paid'] + 1 and
                        col != col_indices['%'] + 1 and col != col_indices['$ Outstanding'] + 1 and
                        col != col_indices['Balance WIP'] + 1):
                    cell = ws.cell(row=row, column=col, value=value)
                    cell.border = Border(left=Side(border_style='thin'),
                                         right=Side(border_style='thin'),
                                         top=Side(border_style='thin'),
                                         bottom=Side(border_style='thin'))
        else:
            for col, value in enumerate(row_data, start=1):
                cell = ws.cell(row=row, column=col, value=value)

    apply_column_widths(final_df, ws)
    header_settings(final_df, ws, shifted_style_skipped_rows)
    alternate_color_fill(final_df, ws, shifted_style_skipped_rows)

    wb.save('data/test.xlsx')


def clean_col_names(buyboard_df, friendswood_df, hcde_df, misc_df, pca_df, usps_df, func, name):
    usps_df.rename(columns={'Facility Name': 'Client', 'Address': 'Location', '%': 'Billed %'}, inplace=True)
    final_usps_df = func(usps_df, title=f'USPS {name}')
    hcde_df.rename(columns={'Type: JOC/HB': 'Type:  JOC, CC, HB', 'Contract ': 'Contract', 'Billed $': 'Bill $'},
                   inplace=True)
    final_hcde_df = func(hcde_df, title=f'HCDE {name}')
    misc_col_names = misc_df.columns
    misc_df.rename(columns={'Type:  JOC, HB': 'Type:  JOC, CC, HB', misc_col_names[1]: 'Contract', 'Client ': 'Client',
                            'Billed $': 'Bill $', 'Comments': 'Comment'}, inplace=True)
    final_misc_df = func(misc_df, title=f'Misc. {name}')
    buyboard_df.rename(columns={'JOC/HB': 'Type:  JOC, CC, HB', 'Billed $': 'Bill $', 'Comments': 'Comment'},
                       inplace=True)
    final_buyboard_df = func(buyboard_df, title=f'Buyboard {name}')
    pca_df.rename(columns={'JOC/HB': 'Type:  JOC, CC, HB', 'Contract #': 'Contract', 'Billed $': 'Bill $'},
                  inplace=True)
    final_pca_df = func(pca_df, title=f'PCA {name}')
    friendswood_df.rename(columns={'JOC/HB': 'Type:  JOC, CC, HB', 'Contract #': 'Contract', 'Billed $': 'Bill $',
                                   'Billed       %': 'Billed %'}, inplace=True)
    final_friendswood_df = func(friendswood_df, title=f'Friendswood {name}')
    return final_buyboard_df, final_friendswood_df, final_hcde_df, final_misc_df, final_pca_df, final_usps_df


def apply_column_widths(final_df, ws):
    for header_name, width in COLUMN_WIDTHS.items():
        col_letter = None
        for col_idx, col in enumerate(final_df.columns, start=1):
            if col == header_name:
                col_letter = get_column_letter(col_idx)
                break
        if col_letter:
            ws.column_dimensions[col_letter].width = width


def set_style(df, ws, col_name, style, idx, wrapText: bool = False):
    for row, value in enumerate(df[col_name][:-1], start=2):
        cell = ws.cell(row=row, column=idx + 1, value=value)
        cell.style = style
        if wrapText:
            cell.alignment = cell.alignment.copy(wrapText=True)


def header_settings(df, ws, shifted_style_skipped_rows=None):
    # This applies the header color of light yellow
    if shifted_style_skipped_rows is None:
        shifted_style_skipped_rows = []

    header_fill_color = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
    for cell in ws[1]:
        cell.fill = header_fill_color
        cell.font = Font(bold=True)

    if not shifted_style_skipped_rows:
        for cell in ws[len(df.index) + 1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='left')
    else:
        for value in shifted_style_skipped_rows:
            for cell in ws[value]:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal='left')
                cell.border = Border()
    for cell in ws[2]:
        cell.font = Font(bold=True)
        cell.border = Border()
        cell.alignment = Alignment(horizontal='left')


def alternate_color_fill(df, ws, shifted_style_skipped_rows=None):
    """
    This function alternates a fill color for every other row. It excludes the header and the final row. If the sheet is
    a merge of multiple dataframes, it also excludes all the total rows that are located using shifted_df_lengths.
    :param df: The final dataframe to be written to the Excel sheet
    :param ws: The Excel sheet.
    :param shifted_style_skipped_rows: The rows to exclude. OPTIONAL.
    """
    if shifted_style_skipped_rows is None:
        shifted_style_skipped_rows = []

    fill_color = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')
    for row_index, row in df.iloc[2:-1:2].iterrows():
        if not shifted_style_skipped_rows:
            for col_index, value in enumerate(row):
                cell = ws.cell(row=row_index + 2, column=col_index + 1)  # +2 to account for 0-based index and header
                cell.value = value
                cell.fill = fill_color
        else:
            if row_index + 2 in shifted_style_skipped_rows:
                for col_index, value in enumerate(row):
                    cell = ws.cell(row=row_index + 2, column=col_index + 1)
                    cell.value = value
            else:
                for col_index, value in enumerate(row):
                    cell = ws.cell(row=row_index + 2, column=col_index + 1)
                    cell.value = value
                    cell.fill = fill_color
