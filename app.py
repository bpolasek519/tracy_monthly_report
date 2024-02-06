import pandas as pd
from openpyxl import Workbook
import src.wrangler as wr
from flask import Flask, render_template, request, send_file
import os
import src.constants as con
import shutil
import src.dataframe_helper as dh
import src.exceptions as e

app = Flask(__name__)


def read_usps_report(usps_file, hcde_file, misc_file, buyboard_file, pca_file, friendswood_file, month, year):
    usps_fmd_df = e.read_excel_with_exception(f'{usps_file}', sheet_name=f'{year} LTD (FMD)')
    hdce_df = e.read_excel_with_exception(f'{hcde_file}', sheet_name=f'{year}')
    llc_df = e.read_excel_with_exception(f'{usps_file}', sheet_name=f'{year} FMD - DPFS LLC')
    misc_df = e.read_excel_with_exception(f'{misc_file}', sheet_name=f'{year}')
    buyboard_df = e.read_excel_with_exception(f'{buyboard_file}', sheet_name=f'{year}')
    pca_df = e.read_excel_with_exception(f'{pca_file}', sheet_name=f'{year}')
    friendswood_df = e.read_excel_with_exception(f'{friendswood_file}', sheet_name=f'{year}')

    wb = Workbook()

    # Create FS Paid
    wr.create_fs_sheet(usps_df=usps_fmd_df, hcde_df=hdce_df, misc_df=misc_df, buyboard_df=buyboard_df,
                       pca_df=pca_df, friendswood_df=friendswood_df, wb=wb, month=month, year=year, fs_type='Paid',
                       df_creation_func=dh.create_paid_df, last_row_columns=con.FS_PAID_LAST_ROW_COLS,
                       columns_to_exclude_from_generic_styles=con.FS_PAID_COLS_TO_EXCLUDE)

    # Create FS Outstanding
    wr.create_fs_sheet(usps_df=usps_fmd_df, hcde_df=hdce_df, misc_df=misc_df, buyboard_df=buyboard_df,
                       pca_df=pca_df, friendswood_df=friendswood_df, wb=wb, month=month, year=year,
                       fs_type='Outstanding', df_creation_func=dh.create_outstanding_df,
                       last_row_columns=con.FS_OUTSTANDING_LAST_ROW_COLS,
                       columns_to_exclude_from_generic_styles=con.FS_OUTSTANDING_COLS_TO_EXCLUDE)

    # Create FS WIP
    wr.create_fs_sheet(usps_df=usps_fmd_df, hcde_df=hdce_df, misc_df=misc_df, buyboard_df=buyboard_df,
                       pca_df=pca_df, friendswood_df=friendswood_df, wb=wb, month=month, year=year, fs_type='WIP',
                       df_creation_func=dh.create_wip_df, last_row_columns=con.FS_WIP_LAST_ROW_COLS,
                       columns_to_exclude_from_generic_styles=con.FS_WIP_COLS_TO_EXCLUDE)

    # Create LLC Paid
    wr.create_llc_sheet(llc_df, wb, llc_type='Paid', last_row_cols=con.LLC_PAID_LAST_ROW_COLS,
                        cols_to_exclude=con.LLC_PAID_COLS_TO_EXCLUDE, month=month, year=year)

    # Create LLC Outstanding
    wr.create_llc_sheet(llc_df, wb, llc_type='Outstanding', last_row_cols=con.LLC_OUTSTANDING_LAST_ROW_COLS,
                        cols_to_exclude=con.LLC_OUTSTANDING_COLS_TO_EXCLUDE, month=month, year=year)

    # Create LLC WIP
    wr.create_llc_sheet(llc_df, wb, llc_type='WIP', last_row_cols=con.LLC_WIP_LAST_ROW_COLS,
                        cols_to_exclude=con.LLC_WIP_COLS_TO_EXCLUDE, month=month, year=year)

    wb.remove_sheet(wb.get_sheet_by_name('Sheet'))

    return wb


def create_choice_partners_report(hcde_file, year, month):
    wb = Workbook()
    sheet = wb.active

    hcde_df = e.read_excel_with_exception(f'{hcde_file}', sheet_name=f'{year}')
    dh.create_cp_completed(hcde_df, month=month)
    wr.create_cp_header(sheet)

    return wb


@app.route('/')
def index():
    return render_template('index.html', month_options=con.MONTH_OPTIONS)


@app.route('/process', methods=['POST'])
def process_spreadsheets():
    year = request.form['year']
    month = request.form['month']
    uploaded_files = request.files.getlist('spreadsheets')

    upload_dir = 'uploads'

    # Check if the upload directory exists and delete its contents if it does
    if os.path.exists(upload_dir):
        for file in os.listdir(upload_dir):
            file_path = os.path.join(upload_dir, file)
            try:
                if os.path.isfile(file_path):
                    os.unlink(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
            except Exception as e:
                print(f"Error deleting {file_path}: {e}")

    os.makedirs(upload_dir, exist_ok=True)

    usps_file = ''
    hcde_file = ''
    misc_file = ''
    buyboard_file = ''
    pca_file = ''
    friendswood_file = ''

    for uploaded_file in uploaded_files:
        if uploaded_file.filename != '':
            if 'USPS' in uploaded_file.filename:
                file_path = os.path.join('uploads', uploaded_file.filename)
                uploaded_file.save(file_path)
                usps_file = file_path
            elif 'HCDE' in uploaded_file.filename:
                file_path = os.path.join('uploads', uploaded_file.filename)
                uploaded_file.save(file_path)
                hcde_file = file_path
            elif 'MISC' in uploaded_file.filename:
                file_path = os.path.join('uploads', uploaded_file.filename)
                uploaded_file.save(file_path)
                misc_file = file_path
            elif 'BuyBoard' in uploaded_file.filename:
                file_path = os.path.join('uploads', uploaded_file.filename)
                uploaded_file.save(file_path)
                buyboard_file = file_path
            elif 'PCA' in uploaded_file.filename:
                file_path = os.path.join('uploads', uploaded_file.filename)
                uploaded_file.save(file_path)
                pca_file = file_path
            elif 'Friendswood' in uploaded_file.filename:
                file_path = os.path.join('uploads', uploaded_file.filename)
                uploaded_file.save(file_path)
                friendswood_file = file_path
            else:
                alert_message = f'Unknown file uploaded: {uploaded_file.filename}'
                return f"<script>alert('{alert_message}'); window.history.back();</script>"

    if (usps_file == '' or hcde_file == '' or misc_file == '' or pca_file == '' or buyboard_file == '' or
            friendswood_file == ''):
        alert_message = f'Missing one or more files to create the WIP file'
        return f"<script>alert('{alert_message}'); window.history.back();</script>"

    try:
        wb = read_usps_report(usps_file=usps_file, hcde_file=hcde_file, friendswood_file=friendswood_file, year=year,
                              misc_file=misc_file, pca_file=pca_file, buyboard_file=buyboard_file, month=month)
        # cp_wb = create_choice_partners_report(hcde_file=hcde_file)

        results_file = f'uploads/{con.MONTH_TO_ZERO_PADDED_NUMBER[month]}-{year} WIP.xlsx'
        # cp_file = f'uploads/Facilities Sources {con.MONTH_TO_ZERO_PADDED_NUMBER[month]} {year}.xlsx'
        wb.save(results_file)
        # cp_wb.save(cp_file)

        return send_file(results_file, as_attachment=True)

    except Exception as ex:
        return f"<script>alert('{ex}'); window.history.back();</script>"


if __name__ == '__main__':
    app.run(debug=True)
