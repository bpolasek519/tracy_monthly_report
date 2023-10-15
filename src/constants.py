from openpyxl.styles import NamedStyle, Alignment, Border, Side

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
