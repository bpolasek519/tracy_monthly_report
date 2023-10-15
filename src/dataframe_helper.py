from src.constants import MONTH_TO_NUMBER
import pandas as pd


def create_wip_df(df: pd.DataFrame, title: str, month: str = '', for_fs: bool = True) -> pd.DataFrame:
    # Filter so job type is only JOC or HB
    mask_job_type = df['Type:  JOC, CC, HB'].str.contains('JOC|HB', case=False, na=False)
    df_filtered_for_job_type = df[mask_job_type].copy()
    df_filtered_for_job_type.reset_index(drop=True, inplace=True)

    # Filter so Awd col has no blanks
    mask_awd = df_filtered_for_job_type['Awd'].isna()
    df_awd_filter = df_filtered_for_job_type[~mask_awd].copy()
    df_awd_filter.reset_index(drop=True, inplace=True)

    # Filter out where Balance Due is not zero
    mask_balance_due = df_awd_filter['Balance Due'] == 0
    df_balance_due_filter = df_awd_filter[~mask_balance_due].copy()
    df_balance_due_filter.reset_index(drop=True, inplace=True)

    # Filter out where Percent is not 100%
    mask_percent = df_balance_due_filter['Billed %'] == 1
    df_percent_filter = df_balance_due_filter[~mask_percent].copy()
    df_percent_filter.reset_index(drop=True, inplace=True)

    # Certain cols to keep for final
    if for_fs:
        cols_to_keep = ['Type:  JOC, CC, HB', 'Contract', 'Proj. #', 'Prob. C/O #', 'Client', 'Location', 'Description',
                        'Awd', 'Awd $', 'Substantial Complete', 'Billed Date', 'Bill $', 'Billed %', 'Comment',
                        '$ Previously Paid']
    else:
        cols_to_keep = ['Type:  JOC, CC, HB', 'Contract', 'Proj. #', 'Prob. C/O #', 'Client', 'Location', 'Description',
                        'Awd', 'WO#', 'Awd $', 'Contract Comp. Date', 'Substantial Complete', 'Billed Date', 'Bill $',
                        'Billed %', 'Comment', '$ Previously Paid']

    final_df = df_percent_filter[cols_to_keep].copy()
    final_df.reset_index(drop=True, inplace=True)

    # Rename the columns to match what is final
    final_df.rename(columns={'Type:  JOC, CC, HB': 'Type: \nJOC, HB', 'Client': 'Facility Name', 'Location': 'Address',
                             'Billed %': '%', 'Prob. C/O #': 'Prob. \n C/O #'},
                    inplace=True)

    final_df['$ Outstanding'] = final_df['Bill $'] - final_df['$ Previously Paid']
    final_df['Balance WIP'] = final_df['Awd $'] - final_df['$ Previously Paid'] - final_df['$ Outstanding']

    if final_df.empty:
        return final_df

    extra_title_row = {'Type: \nJOC, HB': title}
    final_df.loc[-1] = extra_title_row
    final_df.index = final_df.index + 1
    final_df = final_df.sort_index()

    # Adds the total amount of money
    awd_sum = final_df['Awd $'].sum()
    bill_sum = final_df['Bill $'].sum()
    paid_sum = final_df['$ Previously Paid'].sum()
    outstanding_sum = final_df['$ Outstanding'].sum()
    balance_due_sum = final_df['Balance WIP'].sum()
    total_row = {'Type: \nJOC, HB': f'Total {title}', 'Awd $': awd_sum, 'Bill $': bill_sum,
                 '$ Previously Paid': paid_sum, '$ Outstanding': outstanding_sum, 'Balance WIP': balance_due_sum}
    total_row_df = pd.DataFrame(total_row, index=[0])
    final_df = pd.concat([final_df, total_row_df], ignore_index=True)

    return final_df


def create_outstanding_df(df: pd.DataFrame, title: str, month: str = '') -> pd.DataFrame:
    # Filter so job type is only JOC or HB
    mask_job_type = df['Type:  JOC, CC, HB'].str.contains('JOC|HB', case=False, na=False)
    df_filtered_for_job_type = df[mask_job_type].copy()
    df_filtered_for_job_type.reset_index(drop=True, inplace=True)

    # Filter for Paid/Closed that is blank
    mask_paid_closed_blank = df_filtered_for_job_type['Paid/   Closed'].isna()
    df_filtered_paid_closed = df_filtered_for_job_type[mask_paid_closed_blank].copy()
    df_filtered_paid_closed.reset_index(drop=True, inplace=True)

    # Filter that Billed Date is NOT blank
    mask_billed_date_blank = df_filtered_paid_closed['Billed Date'].isna()
    df_filtered_billed_date = df_filtered_paid_closed[~mask_billed_date_blank].copy()
    df_filtered_billed_date.reset_index(drop=True, inplace=True)

    # Filter for Billed - Paid > 0:
    df_filtered_billed_date['Billed-Paid'] = (df_filtered_billed_date['Bill $'] -
                                              (df_filtered_billed_date['$ Previously Paid'] +
                                               df_filtered_billed_date['$ Paid Current Month']))
    mask_diff_zero = df_filtered_billed_date['Billed-Paid'] > 0
    final_df = df_filtered_billed_date[mask_diff_zero].copy()
    final_df.reset_index(drop=True, inplace=True)
    final_df['$ Paid'] = final_df['$ Previously Paid'] + final_df['$ Paid Current Month']

    # Certain cols to keep for final
    cols_to_keep = ['Type:  JOC, CC, HB', 'Contract', 'Proj. #', 'Prob. C/O #', 'Client', 'Location', 'Description',
                    'Awd', 'Awd $', 'Substantial Complete', 'Billed Date', 'Bill $', 'Billed %', 'Comment',
                    '$ Paid', 'Billed-Paid']
    final_df = final_df[cols_to_keep].copy()

    # Rename the columns to match what is final
    final_df.rename(columns={'Type:  JOC, CC, HB': 'Type: \nJOC, HB', 'Client': 'Facility Name', 'Location': 'Address',
                             'Billed %': '%', 'Prob. C/O #': 'Prob. \n C/O #', 'Billed-Paid': 'Balance Due'},
                    inplace=True)

    if final_df.empty:
        return final_df

    extra_title_row = {'Type: \nJOC, HB': title}
    final_df.loc[-1] = extra_title_row
    final_df.index = final_df.index + 1
    final_df = final_df.sort_index()

    # Adds the total amount of money
    awd_sum = final_df['Awd $'].sum()
    bill_sum = final_df['Bill $'].sum()
    paid_sum = final_df['$ Paid'].sum()
    balance_due_sum = final_df['Balance Due'].sum()
    total_row = {'Type: \nJOC, HB': f'Total {title}', 'Awd $': awd_sum, 'Bill $': bill_sum,
                 '$ Paid': paid_sum, 'Balance Due': balance_due_sum}
    total_row_df = pd.DataFrame(total_row, index=[0])
    final_df = pd.concat([final_df, total_row_df], ignore_index=True)

    return final_df


def create_paid_df(df: pd.DataFrame, title: str, month: str) -> pd.DataFrame:
    # Filter so job type is only JOC or HB
    mask_job_type = df['Type:  JOC, CC, HB'].str.contains('JOC|HB', case=False, na=False)
    df_filtered_for_job_type = df[mask_job_type].copy()
    df_filtered_for_job_type.reset_index(drop=True, inplace=True)

    # Filter for paid in the current month
    df_filtered_for_job_type['Paid/   Closed'] = pd.to_datetime(df_filtered_for_job_type['Paid/   Closed'])
    mask_paid_current_month = df_filtered_for_job_type['Paid/   Closed'].dt.month == MONTH_TO_NUMBER[month]
    df_filtered_for_paid_current_month = df_filtered_for_job_type[mask_paid_current_month].copy()

    # Certain cols to keep for final
    cols_to_keep = ['Type:  JOC, CC, HB', 'Contract', 'Proj. #', 'Prob. C/O #', 'Client', 'Location', 'Description',
                    'Awd', 'Awd $', 'Substantial Complete', 'Billed Date', 'Bill $', 'Billed %', 'Comment',
                    '$ Previously Paid', '$ Paid Current Month', 'Balance Due', 'Paid/   Closed']

    # Remove any that haven't been awarded yet
    mask_awd_not_empty = pd.isna(df_filtered_for_paid_current_month['Awd'])
    final_df = df_filtered_for_paid_current_month[~mask_awd_not_empty].reset_index(drop=True)
    final_df = final_df[cols_to_keep].copy()

    if final_df.empty:
        return final_df

    # Rename the columns to match what is final
    final_df.rename(columns={'Type:  JOC, CC, HB': 'Type: \nJOC, HB', 'Client': 'Facility Name', 'Location': 'Address',
                             '$ Paid Current Month': '$ Paid', 'Billed %': '%', 'Paid/   Closed': 'Paid/Closed',
                             'Prob. C/O #': 'Prob. \n C/O #'},
                    inplace=True)
    extra_title_row = {'Type: \nJOC, HB': title}
    final_df.loc[-1] = extra_title_row
    final_df.index = final_df.index + 1
    final_df = final_df.sort_index()

    # Adds the total amount of money
    awd_sum = final_df['Awd $'].sum()
    bill_sum = final_df['Bill $'].sum()
    prev_sum = final_df['$ Previously Paid'].sum()
    paid_sum = final_df['$ Paid'].sum()
    balance_due_sum = final_df['Balance Due'].sum()
    total_row = {'Type: \nJOC, HB': f'Total {title}', 'Awd $': awd_sum, 'Bill $': bill_sum,
                 '$ Previously Paid': prev_sum, '$ Paid': paid_sum, 'Balance Due': balance_due_sum}
    total_row_df = pd.DataFrame(total_row, index=[0])
    final_df = pd.concat([final_df, total_row_df], ignore_index=True)

    return final_df
