from app import read_usps_report, create_choice_partners_report


def test_read_usps_report():

    wb = read_usps_report(
        usps_file='data/current/USPS Status Reports Separated.xlsx',
                     hcde_file='',
                     friendswood_file='',
                     buyboard_file='',
                     pca_file='',
                     misc_file='',
        year=2025, month='August')

    results_file = f'data/current/test_WIP.xlsx'
    wb.save(results_file)

    assert True


def test_cp_file():
    wb = create_choice_partners_report(hcde_file='data/current/87-HCDE Status Report - Current.xlsx', year=2024,
                                       month='June')

    wb.save('data/cp/test.xlsx')
