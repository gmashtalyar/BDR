import pandas as pd
from openpyxl import load_workbook
from pandas import read_excel
from openpyxl.styles import Font


def pasting_amd(i, y):
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
    data_amd = data.loc[(data['Company'] == 'АМД') | (data['Company'] == 'ИП') | (
            (data['Company'] == 'Металлы') & (data['Division'] == 'АТП'))]
    data_amd_checking = data_amd.loc[data_amd['Transaction type'] == 'Расходование']
    data_metally = data.loc[data['Company'] == 'Металлы']

    # Выбираем и группируем доходы
    data_amd_sales = data_amd.loc[data_amd['Transaction type'] == 'Поступление']
    data_metally_sales = data_metally.loc[data_metally['Transaction type'] == 'Поступление']
    data_metally_sales = data_metally_sales.loc[data_metally_sales['Division'] == 'Отдел транспортной логистики']

    # Корректировка ВГО и статей
    data_amd.loc[(data_amd['Transaction type'] == 'Расходование') & (
                data_amd['Counter Party'] == 'Вольфагролес'), 'Expense 1'] = 'Услуги ВГО'  # Вольагролес
    data_amd.loc[(data_amd['Transaction type'] == 'Расходование') & (
            data_amd['Статья бюджета'] == 'Штрафы, пени, неустойки'), 'Expense 1'] = 'Материальные затраты'  # штрафы пени и неустойки
    data_amd.loc[(data_amd['Transaction type'] == 'Расходование') & (
            data_amd['Статья бюджета'] == 'Списание задолжности (расход)'), 'Expense 1'] = 'Списание ДЗ'

    # Объединяем данные по Expense 1, смотрим по необходимости
    expenses_amd = data_amd.groupby(['Expense 1', 'Статья бюджета']).sum().reset_index()
    summa_amd_sales = data_amd_sales.groupby(['Expense type_', 'Counter Party']).sum().reset_index()
    summa_metally_sales = data_metally_sales.groupby('Expense type_').sum().reset_index()

    bdr_sales_amd_otherAM = summa_metally_sales
    bdr_sales_amd_AM = summa_amd_sales.loc[(summa_amd_sales['Expense type_'] == 'ВЫРУЧКА') & (summa_amd_sales['Counter Party'] == 'АРИЭЛЬ МЕТАЛЛ АО')]
    bdr_sales_amd_VAL = summa_amd_sales.loc[(summa_amd_sales['Expense type_'] == 'ВЫРУЧКА') & (summa_amd_sales['Counter Party'] == 'Вольфагролес')]
    bdr_sales_amd_AllClients = summa_amd_sales.loc[summa_amd_sales['Expense type_'] == 'ВЫРУЧКА']
    bdr_sales_amd_OtherSales = summa_amd_sales.loc[summa_amd_sales['Expense type_'] != 'ВЫРУЧКА']

    expenses_amd_FOT = expenses_amd.loc[expenses_amd['Expense 1'] == 'ФОТ']
    expenses_amd_ESN = expenses_amd.loc[expenses_amd['Expense 1'] == 'ЕСН']
    expenses_amd_Personnel = expenses_amd.loc[expenses_amd['Expense 1'] == 'Прочие расходы на персонал']
    expenses_amd_VGO = expenses_amd.loc[expenses_amd['Expense 1'] == 'Услуги ВГО']
    expenses_amd_Material = expenses_amd.loc[expenses_amd['Expense 1'] == 'Материальные затраты']
    expenses_amd_Amortization = expenses_amd.loc[expenses_amd['Expense 1'] == 'Амортизация']
    expenses_amd_Other_taxes = expenses_amd.loc[expenses_amd['Expense 1'] == 'Прочие налоги уплаченные']
    expenses_amd_writeoff = expenses_amd.loc[expenses_amd['Expense 1'] == 'Списание ДЗ']
    expenses_amd_Unknown = expenses_amd.loc[expenses_amd['Expense 1'] == '?!']
    expenses_amd_Taxes = expenses_amd.loc[expenses_amd['Expense 1'] == ' Налог на прибыль ']

    expense_check = expenses_amd_FOT[y].sum() + expenses_amd_ESN[y].sum() + expenses_amd_Personnel[y].sum() + \
                    expenses_amd_VGO[y].sum() + expenses_amd_Material[y].sum() + expenses_amd_Amortization[y].sum() \
                    + expenses_amd_Other_taxes[y].sum() + expenses_amd_writeoff[y].sum() + expenses_amd_Taxes[y].sum()

    print(f'Ошибка при переносе данных АМД в {y} равна {round(expense_check, 2) - round(data_amd_checking[y].sum(), 2)}')

    model_E = load_workbook('data_output/Model.xlsx')
    sheet_amd = model_E['АМД']

    sheet_amd[f'{i}28'] = bdr_sales_amd_otherAM[y].sum()/1000  # прочий доход от АМ
    sheet_amd[f'{i}29'] = bdr_sales_amd_VAL[y].sum()/1000  # ВАЛ
    sheet_amd[f'{i}31'] = bdr_sales_amd_AllClients[y].sum()/1000 - bdr_sales_amd_AM[y].sum()/1000 - bdr_sales_amd_VAL[y].sum()/1000
    sheet_amd[f'{i}32'] = bdr_sales_amd_OtherSales[y].sum()/1000  # прочий доход

    sheet_amd[f'{i}37'] = -expenses_amd_FOT[y].sum()/1000  # фот
    sheet_amd[f'{i}38'] = -expenses_amd_ESN[y].sum()/1000  # есн
    sheet_amd[f'{i}39'] = -expenses_amd_Personnel[y].sum()/1000  # прочие расходы на персонал
    sheet_amd[f'{i}40'] = -expenses_amd_VGO[y].sum()/1000  # услуги вго
    sheet_amd[f'{i}42'] = -expenses_amd_Material[y].sum()/1000  # материальные затраты
    sheet_amd[f'{i}43'] = -expenses_amd_Amortization[y].sum()/1000  # амортизация
    sheet_amd[f'{i}44'] = -expenses_amd_Other_taxes[y].sum()/1000  # прочие налоги
    sheet_amd[f'{i}50'] = -expenses_amd_Taxes[y].sum()/1000  # налог на прибыль
    sheet_amd[f'{i}96'] = -expenses_amd_writeoff[y].sum()/1000  # списание ДЗ
    sheet_amd[f'{i}250'] = -expenses_amd_Unknown[y].sum()/1000  # неизв

    shift = 1
    transcript_cell = 135

    sheet_amd[f'A{transcript_cell-2}'] = 'По-статейные расшифровки:'
    sheet_amd[f'A{transcript_cell-2}'].font = Font(underline='single', bold=True)
    sheet_amd[f'A{transcript_cell-1}'] = 'Статья в Годовой модели'
    sheet_amd[f'B{transcript_cell-1}'] = 'Статья в ERP'
    sheet_amd[f'C{transcript_cell-1}'] = 'Примечание: в годовой модели расходы относятся к юр. лицу, а не ЦФО.'

    a3 = len(bdr_sales_amd_AllClients) + shift
    a4 = a3 + len(bdr_sales_amd_OtherSales) + shift
    a5 = a3 + len(expenses_amd_FOT) + shift
    a6 = a5 + len(expenses_amd_ESN) + shift
    a7 = a6 + len(expenses_amd_Personnel) + shift
    a8 = a7 + len(expenses_amd_VGO) + shift
    a9 = a8 + len(expenses_amd_Material) + shift
    a10 = a9 + len(expenses_amd_Amortization) + shift
    a11 = a10 + len(expenses_amd_Other_taxes) + shift
    a12 = a11 + len(expenses_amd_Taxes) + shift
    a13 = a12 + len(expenses_amd_Unknown) + shift

    for j in range(len(bdr_sales_amd_AllClients)):
        sheet_amd[f'{i}{transcript_cell+j}'] = bdr_sales_amd_AllClients.iloc[j][y]/1000
        sheet_amd[f'A{transcript_cell+j}'] = bdr_sales_amd_AllClients.iat[j, 0]
        sheet_amd[f'B{transcript_cell+j}'] = bdr_sales_amd_AllClients.iat[j, 1]

    for j in range(len(bdr_sales_amd_OtherSales)):
        sheet_amd[f'{i}{260+j+a3+shift}'] = bdr_sales_amd_OtherSales.iloc[j][y]/1000
        sheet_amd[f'A{260+j+a3+shift}'] = bdr_sales_amd_OtherSales.iat[j, 0]
        sheet_amd[f'B{260+j+a3+shift}'] = bdr_sales_amd_OtherSales.iat[j, 1]

    for j in range(len(expenses_amd_FOT)):
        sheet_amd[f'{i}{transcript_cell+j+a3+shift}'] = expenses_amd_FOT.iloc[j][y]/1000
        sheet_amd[f'A{transcript_cell+j+a3+shift}'] = expenses_amd_FOT.iat[j, 0]
        sheet_amd[f'B{transcript_cell+j+a3+shift}'] = expenses_amd_FOT.iat[j, 1]

    for j in range(len(expenses_amd_ESN)):
        sheet_amd[f'{i}{transcript_cell+j+a5+shift}'] = expenses_amd_ESN.iloc[j][y]/1000
        sheet_amd[f'A{transcript_cell+j+a5+shift}'] = expenses_amd_ESN.iat[j, 0]
        sheet_amd[f'B{transcript_cell+j+a5+shift}'] = expenses_amd_ESN.iat[j, 1]

    for j in range(len(expenses_amd_Personnel)):
        sheet_amd[f'{i}{transcript_cell+j+a6+shift}'] = expenses_amd_Personnel.iloc[j][y]/1000
        sheet_amd[f'A{transcript_cell+j+a6+shift}'] = expenses_amd_Personnel.iat[j, 0]
        sheet_amd[f'B{transcript_cell+j+a6+shift}'] = expenses_amd_Personnel.iat[j, 1]

    for j in range(len(expenses_amd_VGO)):
        sheet_amd[f'{i}{transcript_cell+j+a7+shift}'] = expenses_amd_VGO.iloc[j][y]/1000
        sheet_amd[f'A{transcript_cell+j+a7+shift}'] = expenses_amd_VGO.iat[j, 0]
        sheet_amd[f'B{transcript_cell+j+a7+shift}'] = expenses_amd_VGO.iat[j, 1]

    for j in range(len(expenses_amd_Material)):
        sheet_amd[f'{i}{transcript_cell+j+a8+shift}'] = expenses_amd_Material.iloc[j][y]/1000
        sheet_amd[f'A{transcript_cell+j+a8+shift}'] = expenses_amd_Material.iat[j, 0]
        sheet_amd[f'B{transcript_cell+j+a8+shift}'] = expenses_amd_Material.iat[j, 1]

    for j in range(len(expenses_amd_Amortization)):
        sheet_amd[f'{i}{transcript_cell+j+a9+shift}'] = expenses_amd_Amortization.iloc[j][y]/1000
        sheet_amd[f'A{transcript_cell+j+a9+shift}'] = expenses_amd_Amortization.iat[j, 0]
        sheet_amd[f'B{transcript_cell+j+a9+shift}'] = expenses_amd_Amortization.iat[j, 1]

    for j in range(len(expenses_amd_Other_taxes)):
        sheet_amd[f'{i}{transcript_cell+j+a10+shift}'] = expenses_amd_Other_taxes.iloc[j][y]/1000
        sheet_amd[f'A{transcript_cell+j+a10+shift}'] = expenses_amd_Other_taxes.iat[j, 0]
        sheet_amd[f'B{transcript_cell+j+a10+shift}'] = expenses_amd_Other_taxes.iat[j, 1]

    for j in range(len(expenses_amd_Taxes)):
        sheet_amd[f'{i}{transcript_cell+j+a11+shift}'] = expenses_amd_Taxes.iloc[j][y]/1000
        sheet_amd[f'A{transcript_cell+j+a11+shift}'] = expenses_amd_Taxes.iat[j, 0]
        sheet_amd[f'B{transcript_cell+j+a11+shift}'] = expenses_amd_Taxes.iat[j, 1]

    for j in range(len(expenses_amd_Unknown)):
        sheet_amd[f'{i}{transcript_cell+j+a12+shift}'] = expenses_amd_Unknown.iloc[j][y]/1000
        sheet_amd[f'A{transcript_cell+j+a12+shift}'] = expenses_amd_Unknown.iat[j, 0]
        sheet_amd[f'B{transcript_cell+j+a12+shift}'] = expenses_amd_Unknown.iat[j, 1]

    # for j in range(len(bdr_sales_amd_otherAM)):
    #     sheet_amd[f'{i}{transcript_cell+a13+j+shift}'] = bdr_sales_amd_otherAM.iloc[j][y]
    #     sheet_amd[f'A{transcript_cell+a13+j+shift}'] = bdr_sales_amd_otherAM.iat[j, 0]
        # sheet_amd[f'B{260+a13+j+shift}'] = bdr_sales_amd_otherAM.iat[j, 1]

    model_E.save('data_output/Model.xlsx')
