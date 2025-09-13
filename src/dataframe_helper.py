from src.constants import MONTH_TO_NUMBER
import pandas as pd


def create_wip_df(df: pd.DataFrame, title: str, filename: str, month: str = '',
                  for_fs: bool = True) -> pd.DataFrame:
    if 'LLC USPS WIP' in title:
        billed_percent = '%'
    else:
        billed_percent = 'Billed %'

    # Filter so job type is only JOC or HB
    mask_job_type = df['Type:  JOC, CC, HB'].str.contains('JOC|HB', case=False, na=False)
    df_filtered_for_job_type = df[mask_job_type].copy()
    df_filtered_for_job_type.reset_index(drop=True, inplace=True)

    mask_awd_amt_not_zero = df_filtered_for_job_type['Awd $'] != 0
    mask_paid_closed_blank = df_filtered_for_job_type['Paid/   Closed'].isna()
    mask_awd_not_nan = ~df_filtered_for_job_type['Awd $'].isna()
    df_filtered_for_job_type[billed_percent] = round(df_filtered_for_job_type[billed_percent], 2)
    mask_billed_percent_not_100 = df_filtered_for_job_type[billed_percent] != 1

    df_percent_filter = df_filtered_for_job_type[mask_awd_amt_not_zero & mask_paid_closed_blank & mask_awd_not_nan &
                                                 mask_billed_percent_not_100].copy()

    df_percent_filter['Total Paid'] = df_percent_filter['$ Previously Paid'] + df_percent_filter['$ Paid Current Month']
    df_percent_filter['$ Outstanding'] = df_percent_filter['Bill $'] - df_percent_filter['Total Paid']
    df_percent_filter['Balance WIP'] = (df_percent_filter['Awd $'] - df_percent_filter['Total Paid'] -
                                        df_percent_filter['$ Outstanding'])

    # Certain cols to keep for final
    if for_fs:
        cols_to_keep = ['Type:  JOC, CC, HB', 'Contract', 'Proj. #', 'Prob. C/O #', 'Client', 'Location', 'Description',
                        'Awd', 'Awd $', 'Substantial Complete', 'Billed Date', 'Bill $', billed_percent, 'Comment',
                        'Total Paid', '$ Outstanding', 'Balance WIP']
    else:
        cols_to_keep = ['Type:  JOC, CC, HB', 'Contract', 'Proj. #', 'Prob. C/O #', 'Client', 'Location', 'Description',
                        'Awd', 'Awd $', 'Contract Comp. Date', 'Substantial Complete', 'Billed Date', 'Bill $',
                        billed_percent, 'Comment', 'Total Paid', '$ Outstanding', 'Balance WIP']

    missing_columns = [col for col in cols_to_keep if col not in df_percent_filter.columns]

    if len(missing_columns) > 0:
        raise Exception(f'Cannot find column(s): {", ".join(missing_columns)} in {filename} sheet to create {title}. '
                        f'Check to see if the column was renamed to something else or removed.')
    else:
        final_df = df_percent_filter[cols_to_keep].copy()

    final_df.reset_index(drop=True, inplace=True)

    # Rename the columns to match what is final
    final_df.rename(columns={'Type:  JOC, CC, HB': 'Type: \nJOC, HB', 'Client': 'Facility Name', 'Location': 'Address',
                             billed_percent: '%', 'Prob. C/O #': 'Prob. \n C/O #'},
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
    paid_sum = final_df['Total Paid'].sum()
    outstanding_sum = final_df['$ Outstanding'].sum()
    balance_due_sum = final_df['Balance WIP'].sum()
    total_row = {'Type: \nJOC, HB': f'Total {title}', 'Awd $': awd_sum, 'Bill $': bill_sum,
                 'Total Paid': paid_sum, '$ Outstanding': outstanding_sum, 'Balance WIP': balance_due_sum}
    total_row_df = pd.DataFrame(total_row, index=[0])
    final_df = pd.concat([final_df, total_row_df], ignore_index=True)

    return final_df


def create_outstanding_df(df: pd.DataFrame, title: str, filename: str, month: str = '') -> pd.DataFrame:
    if 'LLC USPS Outstanding' in title:
        billed_percent = '%'
    else:
        billed_percent = 'Billed %'

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

    df_filtered_billed_date['Bill $'] = df_filtered_billed_date['Bill $'].astype(float)
    df_filtered_billed_date['Bill $'] = round(df_filtered_billed_date['Bill $'], 2)
    df_filtered_billed_date['$ Previously Paid'] = round(df_filtered_billed_date['$ Previously Paid'], 2)
    df_filtered_billed_date['$ Paid Current Month'] = round(df_filtered_billed_date['$ Paid Current Month'].astype(float), 2)

    # Filter for Billed - Paid > 0:
    df_filtered_billed_date['Billed-Paid'] = (df_filtered_billed_date['Bill $'] -
                                              (df_filtered_billed_date['$ Previously Paid'] +
                                               df_filtered_billed_date['$ Paid Current Month']))
    df_filtered_billed_date['Billed-Paid'] = round(df_filtered_billed_date['Billed-Paid'], 2)
    mask_diff_zero = df_filtered_billed_date['Billed-Paid'] != 0
    final_df = df_filtered_billed_date[mask_diff_zero].copy()
    final_df.reset_index(drop=True, inplace=True)
    final_df['$ Paid'] = final_df['$ Previously Paid'] + final_df['$ Paid Current Month']

    # Certain cols to keep for final
    cols_to_keep = ['Type:  JOC, CC, HB', 'Contract', 'Proj. #', 'Prob. C/O #', 'Client', 'Location', 'Description',
                    'Awd', 'Awd $', 'Substantial Complete', 'Billed Date', 'Bill $', billed_percent, 'Comment',
                    '$ Paid', 'Billed-Paid']

    missing_columns = [col for col in cols_to_keep if col not in final_df.columns]

    if len(missing_columns) > 0:
        raise Exception(f'Cannot find column(s): {", ".join(missing_columns)} in {filename} sheet to create {title}. '
                        f'Check to see if the column was renamed to something else or removed.')
    else:
        final_df = final_df[cols_to_keep].copy()

    # Rename the columns to match what is final
    final_df.rename(columns={'Type:  JOC, CC, HB': 'Type: \nJOC, HB', 'Client': 'Facility Name', 'Location': 'Address',
                             billed_percent: '%', 'Prob. C/O #': 'Prob. \n C/O #', 'Billed-Paid': 'Balance Due'},
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


def create_paid_df(df: pd.DataFrame, title: str, month: str, filename: str) -> pd.DataFrame:
    # Filter so job type is only JOC or HB
    mask_job_type = df['Type:  JOC, CC, HB'].str.contains('JOC|HB', case=False, na=False)
    df_filtered_for_job_type = df[mask_job_type].copy()
    df_filtered_for_job_type.reset_index(drop=True, inplace=True)

    if 'LLC USPS Paid' in title:
        billed_percent = '%'
    else:
        billed_percent = 'Billed %'

    df_filtered_for_job_type['Paid/   Closed'] = pd.to_datetime(df_filtered_for_job_type['Paid/   Closed'])
    mask_current_month_date = df_filtered_for_job_type['Paid/   Closed'].dt.month == MONTH_TO_NUMBER[month]
    df_filtered_for_job_type[billed_percent] = round(df_filtered_for_job_type[billed_percent], 2)
    mask_billed_percent = df_filtered_for_job_type[billed_percent] == 1
    mask_paid_current_month_has_value = ((df_filtered_for_job_type['$ Paid Current Month'] != 0) &
                                         (~df_filtered_for_job_type['$ Paid Current Month'].isna()))
    mask_date_empty = df_filtered_for_job_type['Paid/   Closed'].isna()

    mask_current_month_and_billed_100_percent = mask_current_month_date & mask_billed_percent
    mask_paid_current_month_and_no_date = mask_paid_current_month_has_value & mask_date_empty

    df_filtered_for_paid_current_month = df_filtered_for_job_type[mask_current_month_and_billed_100_percent
                                                                  | mask_paid_current_month_and_no_date].copy()

    # Certain cols to keep for final
    cols_to_keep = ['Type:  JOC, CC, HB', 'Contract', 'Proj. #', 'Prob. C/O #', 'Client', 'Location', 'Description',
                    'Awd', 'Awd $', 'Substantial Complete', 'Billed Date', 'Bill $', billed_percent, 'Comment',
                    '$ Previously Paid', '$ Paid Current Month', 'Balance Due', 'Paid/   Closed']

    # Remove any that haven't been awarded yet
    mask_awd_not_empty = pd.isna(df_filtered_for_paid_current_month['Awd'])
    final_df = df_filtered_for_paid_current_month[~mask_awd_not_empty].reset_index(drop=True)

    missing_columns = [col for col in cols_to_keep if col not in final_df.columns]

    if len(missing_columns) > 0:
        raise Exception(f'Cannot find column(s): {", ".join(missing_columns)} in {filename} sheet to create {title}. '
                        f'Check to see if the column was renamed to something else or removed.')
    else:
        final_df = final_df[cols_to_keep].copy()

    # Rename the columns to match what is final
    final_df.rename(columns={'Type:  JOC, CC, HB': 'Type: \nJOC, HB', 'Client': 'Facility Name', 'Location': 'Address',
                             '$ Paid Current Month': '$ Paid', billed_percent: '%', 'Paid/   Closed': 'Paid/Closed',
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


def create_cp_completed(df: pd.DataFrame, month: str):
    pass
    # TODO: Paid that month or 100% billed is completed
    # TODO: Everything else is Active

    # Confirming that billed is 100%
    mask_100_per = df['Billed %'] >= 1
    mask_paid_current_month = df['Paid/   Closed'].dt.month == MONTH_TO_NUMBER[month]
    mask_paid_null = df['Paid/   Closed'].isna()

    completed_df = df[mask_100_per & (mask_paid_current_month | mask_paid_null)].copy()
    completed_df.reset_index(drop=True, inplace=True)

    mask_balance_due = df['Balance Due'] > 0
    active_df = df[~mask_100_per & (mask_paid_current_month | mask_paid_null) & mask_balance_due].copy()
    active_df.reset_index(drop=True, inplace=True)

    unique_contracts = df['Contract '].unique()

    return completed_df, active_df, unique_contracts
