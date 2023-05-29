import pandas as pd
from openpyxl import load_workbook
from pandas import read_excel


# changing column names for raw data
def renaming():
    wb = load_workbook('data_input/All.xlsx')
    sheet1 = wb['Лист1']
    sheet1['A1'] = 'Company'
    sheet1['B1'] = 'Division'
    sheet1['C1'] = 'Transaction type'
    sheet1['D1'] = 'Expense type'
    sheet1['E1'] = 'Counter Party'
    sheet1['F1'] = 'January'
    sheet1['G1'] = 'February'
    sheet1['H1'] = 'March'
    sheet1['I1'] = 'April'
    sheet1['J1'] = 'May'
    # sheet1['K1'] = 'June'
    # sheet1['L1'] = 'July'
    # sheet1['M1'] = 'August'
    # sheet1['N1'] = 'September'
    # sheet1['O1'] = 'October'
    # sheet1['P1'] = 'November'
    # sheet1['Q1'] = 'December'

    # sheet1.delete_cols(14,1)
    wb.save('data_input/All.xlsx')

# call manually or refer from other scripts
renaming()
