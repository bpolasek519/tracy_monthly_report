from openpyxl.styles import NamedStyle, Alignment, Border, Side

SHEET_ROW_HEIGHT = 30

# Creating the different formatting styles needed
DATE_STYLE = NamedStyle(name='date_style', number_format='mm/dd/yy')
DATE_STYLE.alignment = Alignment(horizontal='center')
DATE_STYLE.border = Border(left=Side(border_style='thin'),
                           right=Side(border_style='thin'),
                           top=Side(border_style='thin'),
                           bottom=Side(border_style='thin'))

ACCOUNTING_STYLE = NamedStyle(name='accounting_style',
                              number_format='_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"_);_(@_)')
ACCOUNTING_STYLE.alignment = Alignment(horizontal='center')
ACCOUNTING_STYLE.border = Border(left=Side(border_style='thin'),
                                 right=Side(border_style='thin'),
                                 top=Side(border_style='thin'),
                                 bottom=Side(border_style='thin'))

PERCENTAGE_STYLE = NamedStyle(name='percentage_style', number_format='0.00%')
PERCENTAGE_STYLE.alignment = Alignment(horizontal='center')
PERCENTAGE_STYLE.border = Border(left=Side(border_style='thin'),
                                 right=Side(border_style='thin'),
                                 top=Side(border_style='thin'),
                                 bottom=Side(border_style='thin'))

NUMBER_STYLE = NamedStyle(name='number_style', number_format='0.00')
NUMBER_STYLE.alignment = Alignment(horizontal='center')
NUMBER_STYLE.border = Border(left=Side(border_style='thin'),
                             right=Side(border_style='thin'),
                             top=Side(border_style='thin'),
                             bottom=Side(border_style='thin'))

CENTER_STYLE = NamedStyle(name='center_style')
CENTER_STYLE.alignment = Alignment(horizontal='center')
CENTER_STYLE.border = Border(left=Side(border_style='thin'),
                             right=Side(border_style='thin'),
                             top=Side(border_style='thin'),
                             bottom=Side(border_style='thin'))

ACCOUNTING_STYLE_NO_BORDER = NamedStyle(name='accounting_style2',
                                        number_format='_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"_);_(@_)')

# Setting specified widths to the columns:
COLUMN_WIDTHS = {
    'Type: \nJOC, HB': 10,
    'Contract': 10,
    'Proj. #': 10,
    'Prob. \n C/O #': 7,
    'Facility Name': 30,
    'Address': 30,
    'Description': 35,
    'Awd': 12,
    'Awd $': 16,
    'Substantial Complete': 12,
    'Contract Comp. Date': 12,
    'Billed Date': 12,
    'Bill $': 15,
    '%': 10,
    'Comment': 35,
    '$ Previously Paid': 15,
    '$ Paid': 15,
    'Balance Due': 15,
    '$ Outstanding': 15,
    'Balance WIP': 16,
    'Paid/Closed': 12
}

STYLE_MAPPINGS = {
    'Comment': CENTER_STYLE,
    'Awd': DATE_STYLE,
    'Substantial Complete': DATE_STYLE,
    'Billed Date': DATE_STYLE,
    'Paid/Closed': DATE_STYLE,
    'Awd $': ACCOUNTING_STYLE,
    'Bill $': ACCOUNTING_STYLE,
    '$ Paid': ACCOUNTING_STYLE,
    '$ Previously Paid': ACCOUNTING_STYLE,
    'Balance Due': ACCOUNTING_STYLE,
    '%': PERCENTAGE_STYLE,
    'WO#': NUMBER_STYLE,
    'Contract Comp. Date': DATE_STYLE,
    '$ Outstanding': ACCOUNTING_STYLE,
    'Balance WIP': ACCOUNTING_STYLE
}

MONTH_TO_NUMBER = {
    "January": 1,
    "February": 2,
    "March": 3,
    "April": 4,
    "May": 5,
    "June": 6,
    "July": 7,
    "August": 8,
    "September": 9,
    "October": 10,
    "November": 11,
    "December": 12
}

MONTH_TO_ZERO_PADDED_NUMBER = {
    "January": '01',
    "February": '02',
    "March": '03',
    "April": '04',
    "May": '05',
    "June": '06',
    "July": '07',
    "August": '08',
    "September": '09',
    "October": '10',
    "November": '11',
    "December": '12'
}

MONTH_OPTIONS = [
    "January", "February", "March", "April",
    "May", "June", "July", "August",
    "September", "October", "November", "December"
]

FS_PAID_LAST_ROW_COLS = ['Balance Due', 'Awd $', 'Bill $', '$ Previously Paid', '$ Paid']
FS_PAID_COLS_TO_EXCLUDE = ['Awd', 'Substantial Complete', 'Billed Date', 'Paid/Closed', 'Awd $', 'Bill $', '$ Paid',
                           'Balance Due', '%', 'Comment', '$ Previously Paid']

FS_OUTSTANDING_LAST_ROW_COLS = ['Awd $', 'Bill $', '$ Paid', 'Balance Due']
FS_OUTSTANDING_COLS_TO_EXCLUDE = ['Awd', 'Substantial Complete', 'Billed Date', 'Awd $', 'Bill $', '$ Paid',
                                  'Balance Due', '%', 'Comment']

FS_WIP_LAST_ROW_COLS = ['Balance WIP', 'Awd $', 'Bill $', 'Total Paid', '$ Outstanding']
FS_WIP_COLS_TO_EXCLUDE = ['Awd', 'Substantial Complete', 'Billed Date', 'Awd $', 'Bill $', 'Total Paid',
                          'Balance WIP', '%', '$ Outstanding']

LLC_WIP_LAST_ROW_COLS = ['Balance WIP', 'Awd $', 'Bill $', 'Total Paid', '$ Outstanding']
LLC_WIP_COLS_TO_EXCLUDE = ['Awd', 'Substantial Complete', 'Billed Date', 'Awd $', 'Bill $', 'Contract Comp. Date',
                           'Total Paid', '%', '$ Outstanding', 'Balance WIP']

LLC_OUTSTANDING_LAST_ROW_COLS = ['Balance Due', 'Awd $', 'Bill $', '$ Paid']
LLC_OUTSTANDING_COLS_TO_EXCLUDE = ['Awd', 'Substantial Complete', 'Billed Date', 'Awd $', 'Bill $', '$ Paid',
                                   'Balance Due', '%', 'Comment']

LLC_PAID_LAST_ROW_COLS = ['Balance Due', 'Awd $', 'Bill $', '$ Previously Paid', '$ Paid']
LLC_PAID_COLS_TO_EXCLUDE = ['Awd', 'Substantial Complete', 'Billed Date', 'Paid/Closed', 'Awd $', 'Bill $', '$ Paid',
                            'Balance Due', '%', 'Comment', '$ Previously Paid']
