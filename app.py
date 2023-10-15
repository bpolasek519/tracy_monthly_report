import pandas as pd
from openpyxl import Workbook
import src.wrangler as wr
import src.dataframe_helper as dh


# App needs to read in the sheet names, year, month


def read_usps_report():
    usps_fmd_df = pd.read_excel('data/USPS Status Reports Separated - 7.23 Revised.xlsx', sheet_name='2023 LTD (FMD)')
    hdce_df = pd.read_excel('data/87-HCDE Status Report - 7.23 Revised.xlsx', sheet_name='2023')
    llc_df = pd.read_excel('data/USPS Status Reports Separated - 7.23 Revised.xlsx', sheet_name='2023 FMD - DPFS LLC')
    misc_df = pd.read_excel('data/76 MISC. JOBS STATUS REPORT-7.23 Revised.xlsx', sheet_name='2023')
    buyboard_df = pd.read_excel('data/98 BuyBoard Status Report-7.23 Revised.xlsx', sheet_name='2023')
    pca_df = pd.read_excel('data/62 PCA STATUS REPORT-7.23 Revised.xlsx', sheet_name='2023')
    friendswood_df = pd.read_excel('data/21 Friendswood ISD STATUS REPORT-7.23 Revised.xlsx', sheet_name='2023')

    wb = Workbook()

    wr.create_llc_paid_sheet(llc_df, wb)
    wr.create_llc_outstanding_sheet(llc_df, wb)
    wr.create_llc_wip_sheet(llc_df, wb)
    wr.create_fs_paid_sheet(usps_df=usps_fmd_df, hcde_df=hdce_df, misc_df=misc_df, buyboard_df=buyboard_df,
                            pca_df=pca_df, friendswood_df=friendswood_df, wb=wb)
    wr.create_fs_outstanding_sheet(usps_df=usps_fmd_df, hcde_df=hdce_df, misc_df=misc_df, buyboard_df=buyboard_df,
                                   pca_df=pca_df, friendswood_df=friendswood_df, wb=wb)
    wr.create_fs_wip_sheet(usps_df=usps_fmd_df, hcde_df=hdce_df, misc_df=misc_df, buyboard_df=buyboard_df,
                           pca_df=pca_df, friendswood_df=friendswood_df, wb=wb)
