import pandas as pd
from openpyxl import load_workbook
from pandas import read_excel
from openpyxl.styles import Font



def pasting_atmd(i, y):
    # читаем данные, устанавливаем expense id
    raw_data = read_excel('data_input/All.xlsx')
    raw_data.set_index(["Expense type"], inplace=True)
    raw_data['Counter Party'] = raw_data['Counter Party'].fillna('no_data')
    # читаем форматирование расходов по годовой модели, устанавливаем expense id
    formatting_expenses = read_excel('data_input/Format_expenses.xlsx')
    formatting_expenses["Expense type"] = formatting_expenses["Статья бюджета"].astype(str)
    formatting_expenses.set_index('Expense type', inplace=True)
    # читаем форматирование отделов по годовой модели, устанавливаем division id
    formatting_divisions = read_excel('data_input/Format_divisions.xlsx')
    formatting_divisions["Division id"] = formatting_divisions["Division"].astype(str)
    formatting_divisions.set_index('Division id', inplace=True)
    # получаем отформатированный доступ к данным
    data = raw_data.merge(formatting_expenses, left_index=True, right_index=True, how="outer")
    data["Expense type_"] = data.index
    # выбираем данные по Company
    data_atmd = data.loc[data['Company'] == 'АТМД']
    data_atmd_checking = data_atmd.loc[data_atmd['Transaction type'] == 'Расходование']

    # Выбираем и группируем доходы
    data_atmd_sales = data_atmd.loc[data_atmd['Transaction type'] == 'Поступление']

    # Корректировка ВГО и статей
    data_atmd.loc[(data_atmd['Transaction type'] == 'Расходование') & (
            data_atmd['Counter Party'] == 'Вольфагролес'), 'Expense 1'] = 'Услуги ВГО'
    data_atmd.loc[(data_atmd['Transaction type'] == 'Расходование') & (
            data_atmd['Статья бюджета'] == 'Штрафы, пени, неустойки'), 'Expense 1'] = 'Материальные затраты'
    data_atmd.loc[(data_atmd['Transaction type'] == 'Расходование') & (
            data_atmd['Статья бюджета'] == 'Продажа ОС (остаточная стоимость)'), 'Expense 1'] = 'Материальные затраты'

    # Объединяем данные по Expense 1, смотрим по необходимости
    expenses_atmd = data_atmd.groupby(['Expense 1', 'Статья бюджета']).sum().reset_index()
    summa_atmd_sales = data_atmd_sales.groupby(['Expense type_', 'Counter Party']).sum().reset_index()

    bdr_sales_atmd_AM = summa_atmd_sales.loc[(summa_atmd_sales['Expense type_'] == 'IT сопровождение') & (
            summa_atmd_sales['Counter Party'] == 'АРИЭЛЬ МЕТАЛЛ АО')]
    bdr_sales_atmd_AMD = summa_atmd_sales.loc[(summa_atmd_sales['Expense type_'] == 'IT сопровождение') & (
            summa_atmd_sales['Counter Party'] == 'АМД')]
    bdr_sales_atmd_VAL = summa_atmd_sales.loc[(summa_atmd_sales['Expense type_'] == 'IT сопровождение') & (
            summa_atmd_sales['Counter Party'] == 'Вольфагролес')]
    bdr_sales_atmd_OtherClients = summa_atmd_sales.loc[summa_atmd_sales['Expense type_'] == 'IT сопровождение']
    bdr_sales_atmd_OtherSales = summa_atmd_sales.loc[summa_atmd_sales['Expense type_'] != 'IT сопровождение']

    expenses_atmd_FOT = expenses_atmd.loc[expenses_atmd['Expense 1'] == 'ФОТ']
    expenses_atmd_ESN = expenses_atmd.loc[expenses_atmd['Expense 1'] == 'ЕСН']
    expenses_atmd_Personnel = expenses_atmd.loc[expenses_atmd['Expense 1'] == 'Прочие расходы на персонал']
    expenses_atmd_VGO = expenses_atmd.loc[expenses_atmd['Expense 1'] == 'Услуги ВГО']
    expenses_atmd_Material = expenses_atmd.loc[expenses_atmd['Expense 1'] == 'Материальные затраты']
    expenses_atmd_Amortization = expenses_atmd.loc[expenses_atmd['Expense 1'] == 'Амортизация']
    expenses_atmd_DebtInterest = expenses_atmd.loc[expenses_atmd['Expense 1'] == 'Проценты']
    expenses_atmd_Unknown = expenses_atmd.loc[expenses_atmd['Expense 1'] == '?!']
    expenses_atmd_Taxes = expenses_atmd.loc[expenses_atmd['Expense 1'] == ' Налог на прибыль ']

    expense_check = expenses_atmd_FOT[y].sum() + expenses_atmd_ESN[y].sum() + expenses_atmd_Personnel[y].sum() + \
                    expenses_atmd_VGO[y].sum() + expenses_atmd_Material[y].sum() + expenses_atmd_Amortization[y].sum() \
                    + expenses_atmd_DebtInterest[y].sum() + expenses_atmd_Taxes[y].sum()

    print(f'Ошибка при переносе данных АТМД в {y} равна {round(expense_check, 2) - round(data_atmd_checking[y].sum(), 2)}')

    model_E = load_workbook('data_output/Model.xlsx')
    sheet_atdm = model_E['АТМД']

    sheet_atdm[f'{i}15'] = bdr_sales_atmd_AM[y].sum()/1000
    sheet_atdm[f'{i}16'] = bdr_sales_atmd_AMD[y].sum()/1000
    sheet_atdm[f'{i}17'] = bdr_sales_atmd_VAL[y].sum()/1000
    sheet_atdm[f'{i}18'] = bdr_sales_atmd_OtherClients[y].sum()/1000 - bdr_sales_atmd_AM[y].sum()/1000 - bdr_sales_atmd_AMD[y].sum()/1000 - bdr_sales_atmd_VAL[y].sum()/1000
    sheet_atdm[f'{i}19'] = bdr_sales_atmd_OtherSales[y].sum()/1000

    sheet_atdm[f'{i}24'] = -expenses_atmd_FOT[y].sum()/1000
    sheet_atdm[f'{i}25'] = -expenses_atmd_ESN[y].sum()/1000
    sheet_atdm[f'{i}26'] = -expenses_atmd_Personnel[y].sum()/1000
    sheet_atdm[f'{i}27'] = -expenses_atmd_VGO[y].sum()/1000
    sheet_atdm[f'{i}28'] = -expenses_atmd_Material[y].sum()/1000
    sheet_atdm[f'{i}29'] = -expenses_atmd_Amortization[y].sum()/1000
    sheet_atdm[f'{i}51'] = -expenses_atmd_DebtInterest[y].sum()/1000
    sheet_atdm[f'{i}250'] = -expenses_atmd_Unknown[y].sum()/1000
    sheet_atdm[f'{i}36'] = -expenses_atmd_Taxes[y].sum()/1000

    shift = 1
    transcript_cell = 116

    sheet_atdm[f'A{transcript_cell-2}'] = 'По-статейные расшифровки:'
    sheet_atdm[f'A{transcript_cell-2}'].font = Font(underline='single', bold=True)
    sheet_atdm[f'A{transcript_cell-1}'] = 'Статья в Годовой модели'
    sheet_atdm[f'B{transcript_cell-1}'] = 'Статья в ERP'
    sheet_atdm[f'C{transcript_cell-1}'] = 'Примечание: в годовой модели расходы относятся к юр. лицу, а не ЦФО.'


    a3 = len(bdr_sales_atmd_OtherClients) + shift
    a4 = a3 + len(bdr_sales_atmd_OtherSales) + shift
    a5 = a4 + len(expenses_atmd_FOT) + shift
    a6 = a5 + len(expenses_atmd_ESN) + shift
    a7 = a6 + len(expenses_atmd_Personnel) + shift
    a8 = a7 + len(expenses_atmd_VGO) + shift
    a9 = a8 + len(expenses_atmd_Material) + shift
    a10 = a9 + len(expenses_atmd_Amortization) + shift
    a11 = a10 + len(expenses_atmd_DebtInterest) + shift
    a12 = a11 + len(expenses_atmd_Unknown) + shift
    a13 = a12 + len(expenses_atmd_Taxes) + shift

    for j in range(len(bdr_sales_atmd_OtherClients)):
        sheet_atdm[f'{i}{transcript_cell+j}'] = bdr_sales_atmd_OtherClients.iloc[j][y]/1000
        sheet_atdm[f'A{transcript_cell+j}'] = bdr_sales_atmd_OtherClients.iat[j, 0]
        sheet_atdm[f'B{transcript_cell+j}'] = bdr_sales_atmd_OtherClients.iat[j, 1]

    for j in range(len(bdr_sales_atmd_OtherSales)):
        sheet_atdm[f'{i}{transcript_cell+j+a3+shift}'] = bdr_sales_atmd_OtherSales.iloc[j][y]/1000
        sheet_atdm[f'A{transcript_cell+j+a3+shift}'] = bdr_sales_atmd_OtherSales.iat[j, 0]
        sheet_atdm[f'B{transcript_cell+j+a3+shift}'] = bdr_sales_atmd_OtherSales.iat[j, 1]

    for j in range(len(expenses_atmd_FOT)):
        sheet_atdm[f'{i}{transcript_cell+j+a4+shift}'] = expenses_atmd_FOT.iloc[j][y]/1000
        sheet_atdm[f'A{transcript_cell+j+a4+shift}'] = expenses_atmd_FOT.iat[j, 0]
        sheet_atdm[f'B{transcript_cell+j+a4+shift}'] = expenses_atmd_FOT.iat[j, 1]

    for j in range(len(expenses_atmd_ESN)):
        sheet_atdm[f'{i}{transcript_cell+j+a5+shift}'] = expenses_atmd_ESN.iloc[j][y]/1000
        sheet_atdm[f'A{transcript_cell+j+a5+shift}'] = expenses_atmd_ESN.iat[j, 0]
        sheet_atdm[f'B{transcript_cell+j+a5+shift}'] = expenses_atmd_ESN.iat[j, 1]

    for j in range(len(expenses_atmd_Personnel)):
        sheet_atdm[f'{i}{transcript_cell+j+a6+shift}'] = expenses_atmd_Personnel.iloc[j][y]/1000
        sheet_atdm[f'A{transcript_cell+j+a6+shift}'] = expenses_atmd_Personnel.iat[j, 0]
        sheet_atdm[f'B{transcript_cell+j+a6+shift}'] = expenses_atmd_Personnel.iat[j, 1]

    for j in range(len(expenses_atmd_VGO)):
        sheet_atdm[f'{i}{transcript_cell+j+a7+shift}'] = expenses_atmd_VGO.iloc[j][y]/1000
        sheet_atdm[f'A{transcript_cell+j+a7+shift}'] = expenses_atmd_VGO.iat[j, 0]
        sheet_atdm[f'B{transcript_cell+j+a7+shift}'] = expenses_atmd_VGO.iat[j, 1]

    for j in range(len(expenses_atmd_Material)):
        sheet_atdm[f'{i}{transcript_cell+j+a8+shift}'] = expenses_atmd_Material.iloc[j][y]/1000
        sheet_atdm[f'A{transcript_cell+j+a8+shift}'] = expenses_atmd_Material.iat[j, 0]
        sheet_atdm[f'B{transcript_cell+j+a8+shift}'] = expenses_atmd_Material.iat[j, 1]

    for j in range(len(expenses_atmd_Amortization)):
        sheet_atdm[f'{i}{transcript_cell+j+a9+shift}'] = expenses_atmd_Amortization.iloc[j][y]/1000
        sheet_atdm[f'A{transcript_cell+j+a9+shift}'] = expenses_atmd_Amortization.iat[j, 0]
        sheet_atdm[f'B{transcript_cell+j+a9+shift}'] = expenses_atmd_Amortization.iat[j, 1]

    for j in range(len(expenses_atmd_DebtInterest)):
        sheet_atdm[f'{i}{transcript_cell+j+a10+shift}'] = expenses_atmd_DebtInterest.iloc[j][y]/1000
        sheet_atdm[f'A{transcript_cell+j+a10+shift}'] = expenses_atmd_DebtInterest.iat[j, 0]
        sheet_atdm[f'B{transcript_cell+j+a10+shift}'] = expenses_atmd_DebtInterest.iat[j, 1]

    for j in range(len(expenses_atmd_Unknown)):
        sheet_atdm[f'{i}{transcript_cell+j+a11+shift}'] = expenses_atmd_Unknown.iloc[j][y]/1000
        sheet_atdm[f'A{transcript_cell+j+a11+shift}'] = expenses_atmd_Unknown.iat[j, 0]
        sheet_atdm[f'B{transcript_cell+j+a11+shift}'] = expenses_atmd_Unknown.iat[j, 1]

    for j in range(len(expenses_atmd_Taxes)):
        sheet_atdm[f'{i}{transcript_cell+j+a12+shift}'] = expenses_atmd_Taxes.iloc[j][y]/1000
        sheet_atdm[f'A{transcript_cell+j+a12+shift}'] = expenses_atmd_Taxes.iat[j, 0]
        sheet_atdm[f'B{transcript_cell+j+a12+shift}'] = expenses_atmd_Taxes.iat[j, 1]

    model_E.save('data_output/Model.xlsx')
