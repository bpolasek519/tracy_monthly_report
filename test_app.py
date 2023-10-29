from app import read_usps_report


def test_read_usps_report():

    read_usps_report(usps_file='data/current/USPS Status Reports Separated.xlsx',
                     hcde_file='data/current/87-HCDE Status Report - Current.xlsx',
                     friendswood_file='data/current/21 Friendswood ISD STATUS REPORT-CURRENT.xlsx',
                     buyboard_file='data/current/98 BuyBoard Status Report-New.xlsx',
                     pca_file='data/current/62 PCA STATUS REPORT-CURRENT.xlsx',
                     misc_file='data/current/76 MISC. JOBS STATUS REPORT-CURRENT.xlsx', year=2023, month='October')

    assert True
