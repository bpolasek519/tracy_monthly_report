from app import read_usps_report, create_choice_partners_report


def test_read_usps_report():

    wb = read_usps_report(usps_file='data/current/USPS Status Reports Separated.xlsx',
                     hcde_file='data/current/87-HCDE Status Report - Current.xlsx',
                     friendswood_file='data/current/21 Friendswood ISD STATUS REPORT-CURRENT.xlsx',
                     buyboard_file='data/current/98 BuyBoard Status Report-New.xlsx',
                     pca_file='data/current/62 PCA STATUS REPORT-CURRENT.xlsx',
                     misc_file='data/current/76 MISC. JOBS STATUS REPORT-CURRENT.xlsx', year=2024, month='June')

    results_file = f'data/current/test_WIP.xlsx'
    wb.save(results_file)

    assert True


def test_cp_file():
    wb = create_choice_partners_report(hcde_file='data/current/87-HCDE Status Report - Current.xlsx', year=2024,
                                       month='June')

    wb.save('data/cp/test.xlsx')
