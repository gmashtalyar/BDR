import pandas as pd
from openpyxl import load_workbook
from pandas import read_excel
from openpyxl.styles import Font



def pasting_val(i, y):
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
    data_val = data.loc[(data['Company'] == 'Вольфагролес') | ((data['Company'] == 'Металлы') & (data['Division'] == 'Склад Подольск'))]
    data_val_checking = data_val.loc[data_val['Transaction type'] == 'Расходование']

    # Выбираем и группируем доходы и расходы
    data_val_expenses = data_val.loc[data_val['Transaction type'] == 'Расходование']
    data_val_sales = data_val.loc[data_val['Transaction type'] == 'Поступление']

    # Корректировка ВГО и статей
    data_val_expenses.loc[
        (data_val_expenses['Transaction type'] == 'Расходование') & (data_val_expenses['Counter Party'] == 'АРИЭЛЬ МЕТАЛЛ АО') & (
                    data_val_expenses['Статья бюджета'] == 'Ремонт склада'), 'Expense 1'] = 'Услуги ВГО' # ремонт склада
    data_val_expenses.loc[(data_val_expenses['Transaction type'] == 'Расходование') & (
            data_val_expenses['Статья бюджета'] == 'Штрафы, пени, неустойки'), 'Expense 1'] = 'Материальные затраты' # штрафы пени и неустойки

    # Объединяем данные по Expense 1, смотрим по необходимости
    expenses_val = data_val_expenses.groupby(['Expense 1', 'Статья бюджета']).sum().reset_index()
    summa_val_sales = data_val_sales.groupby(['Expense type_', 'Counter Party']).sum().reset_index()

    bdr_sales_val_AM = summa_val_sales.loc[(summa_val_sales['Expense type_'] == 'ВЫРУЧКА') & (summa_val_sales['Counter Party'] == 'АРИЭЛЬ МЕТАЛЛ АО')]
    bdr_sales_val_AMD = summa_val_sales.loc[(summa_val_sales['Expense type_'] == 'ВЫРУЧКА') & (summa_val_sales['Counter Party'] == 'АМД')]
    bdr_sales_val_IP = summa_val_sales.loc[(summa_val_sales['Expense type_'] == 'ВЫРУЧКА') & (summa_val_sales['Counter Party'] == 'Овсянников С.М. ИП')]
    bdr_sales_val_ATMD = summa_val_sales.loc[(summa_val_sales['Expense type_'] == 'ВЫРУЧКА') & (summa_val_sales['Counter Party'] == 'АТМД')]
    bdr_sales_val_TS = summa_val_sales.loc[(summa_val_sales['Expense type_'] == 'ВЫРУЧКА') & (summa_val_sales['Counter Party'] == 'ТД ТУЛА-СТАЛЬ')]
    bdr_sales_val_AllClients = summa_val_sales.loc[summa_val_sales['Expense type_'] == 'ВЫРУЧКА']
    bdr_sales_val_OtherSales = summa_val_sales.loc[summa_val_sales['Expense type_'] != 'ВЫРУЧКА']
    bdr_sales_val_Other_AM = summa_val_sales.loc[summa_val_sales['Expense type_'] == 'Компенсация по выписке докум. (Логисты)']
    bdr_sales_val_Metally_AMD1 = summa_val_sales.loc[summa_val_sales['Expense type_'] == 'Аренда стоянки (доход)']
    bdr_sales_val_Metally_AMD2 = summa_val_sales.loc[summa_val_sales['Expense type_'] == 'Компенсация электр-во (доход)']
    bdr_sales_val_Metally = summa_val_sales

    expenses_val_FOT = expenses_val.loc[expenses_val['Expense 1'] == 'ФОТ']
    expenses_val_ESN = expenses_val.loc[expenses_val['Expense 1'] == 'ЕСН']
    expenses_val_Personnel = expenses_val.loc[expenses_val['Expense 1'] == 'Прочие расходы на персонал']
    expenses_val_VGO = expenses_val.loc[expenses_val['Expense 1'] == 'Услуги ВГО']
    expenses_val_PPGT_TS = expenses_val.loc[expenses_val['Expense 1'] == 'Компенсация ППЖТ ХРАНИТЕЛЯМИ без НДС']
    expenses_val_Material = expenses_val.loc[expenses_val['Expense 1'] == 'Материальные затраты']
    expenses_val_Amortization = expenses_val.loc[expenses_val['Expense 1'] == 'Амортизация']
    expenses_val_Other_taxes = expenses_val.loc[expenses_val['Expense 1'] == 'Прочие налоги уплаченные']
    expenses_val_Foreign_interest = expenses_val.loc[expenses_val['Expense 1'] == '%  уплаченный']
    expenses_val_interest_AM = expenses_val.loc[expenses_val['Expense 1'] == 'Проценты']
    expenses_val_Unknown = expenses_val.loc[expenses_val['Expense 1'] == '?!']

    expense_check = expenses_val_FOT[y].sum() + expenses_val_ESN[y].sum() + expenses_val_Personnel[y].sum() + \
                    expenses_val_VGO[y].sum() + expenses_val_Material[y].sum() + expenses_val_Amortization[y].sum() \
                    + expenses_val_Other_taxes[y].sum() + expenses_val_PPGT_TS[y].sum() + \
                    expenses_val_Foreign_interest[y].sum() + expenses_val_interest_AM[y].sum()

    print(f'Ошибка при переносе данных ВАЛ в {y} равна {round(expense_check, 2) - round(data_val_checking[y].sum(), 2)}')

    model_E = load_workbook('data_output/Model.xlsx')
    sheet_val = model_E['ВАЛ']
    sheet_am = model_E['АМ']

    sheet_val[f'{i}20'] = bdr_sales_val_Other_AM[y].sum()/1000  # прочие услугли АМ
    sheet_val[f'{i}21'] = bdr_sales_val_AMD[y].sum()/1000 + bdr_sales_val_IP[y].sum()/1000 + bdr_sales_val_Metally_AMD1[y].sum()/1000 + bdr_sales_val_Metally_AMD2[y].sum()/1000  # АМД
    sheet_val[f'{i}22'] = bdr_sales_val_ATMD[y].sum()/1000  # АТМД
    sheet_val[f'{i}23'] = bdr_sales_val_TS[y].sum()/1000 + expenses_val_PPGT_TS[y].sum()/1000 # ТС
    sheet_val[f'{i}24'] = bdr_sales_val_OtherSales[y].sum()/1000 - bdr_sales_val_Metally_AMD1[y].sum()/1000 - bdr_sales_val_Metally_AMD2[y].sum()/1000 - bdr_sales_val_Other_AM[y].sum()/1000 # прочие

    sheet_val[f'{i}26'] = -expenses_val_PPGT_TS[y].sum()/1000  # ППЖТ хранителей
    sheet_val[f'{i}30'] = -expenses_val_FOT[y].sum()/1000  # фот
    sheet_val[f'{i}31'] = -expenses_val_ESN[y].sum()/1000  # есн
    sheet_val[f'{i}32'] = -expenses_val_Personnel[y].sum()/1000  # прочие расходы на персонал
    sheet_val[f'{i}33'] = -expenses_val_VGO[y].sum()/1000  # услуги вго
    sheet_val[f'{i}35'] = -expenses_val_Material[y].sum()/1000  # материальные затраты
    sheet_val[f'{i}36'] = -expenses_val_Amortization[y].sum()/1000  # амортизация
    sheet_val[f'{i}37'] = -expenses_val_Other_taxes[y].sum()/1000  # прочие налоги
    sheet_val[f'{i}64'] = -expenses_val_Foreign_interest[y].sum()/1000  # проценты по стороннему займу
    sheet_am[f'{i}153'] = -expenses_val_interest_AM[y].sum()/1000  # проценты по стороннему займу
    sheet_val[f'{i}250'] = -expenses_val_Unknown[y].sum()/1000  # неизв


    shift = 1
    transcript_cell = 134

    sheet_val[f'A{transcript_cell-2}'] = 'По-статейные расшифровки:'
    sheet_val[f'A{transcript_cell-2}'].font = Font(underline='single', bold=True)
    sheet_val[f'A{transcript_cell-1}'] = 'Статья в Годовой модели'
    sheet_val[f'B{transcript_cell-1}'] = 'Статья в ERP'
    sheet_val[f'C{transcript_cell-1}'] = 'Примечание: в годовой модели расходы относятся к юр. лицу, а не ЦФО.'

    a3 = len(bdr_sales_val_AllClients) + shift
    a4 = a3 + len(bdr_sales_val_OtherSales) + shift
    a5 = a4 + len(expenses_val) + shift
    # a6 = a5 + len(expenses_val) + shift
    # a7 = a6 + len(expenses_amd_Personnel) + shift
    # a8 = a7 + len(expenses_amd_VGO) + shift
    # a9 = a8 + len(expenses_amd_Material) + shift
    # a10 = a9 + len(expenses_amd_Amortization) + shift
    # a11 = a10 + len(expenses_amd_Other_taxes) + shift
    # a12 = a11 + len(expenses_amd_Taxes) + shift
    # a13 = a12 + len(expenses_amd_Unknown) + shift

    for j in range(len(bdr_sales_val_AllClients)):
        sheet_val[f'{i}{transcript_cell+j}'] = bdr_sales_val_AllClients.iloc[j][y]/1000
        sheet_val[f'A{transcript_cell+j}'] = bdr_sales_val_AllClients.iat[j, 0]
        sheet_val[f'B{transcript_cell+j}'] = bdr_sales_val_AllClients.iat[j, 1]

    for j in range(len(bdr_sales_val_OtherSales)):
        sheet_val[f'{i}{transcript_cell+j+a3+shift}'] = bdr_sales_val_OtherSales.iloc[j][y]/1000
        sheet_val[f'A{transcript_cell+j+a3+shift}'] = bdr_sales_val_OtherSales.iat[j, 0]
        sheet_val[f'B{transcript_cell+j+a3+shift}'] = bdr_sales_val_OtherSales.iat[j, 1]

    for j in range(len(expenses_val)):
        sheet_val[f'{i}{transcript_cell+j+a4+shift}'] = expenses_val.iloc[j][y]/1000
        sheet_val[f'A{transcript_cell+j+a4+shift}'] = expenses_val.iat[j, 0]
        sheet_val[f'B{transcript_cell+j+a4+shift}'] = expenses_val.iat[j, 1]

    # for j in range(len(expenses_val)):
    #     sheet_val[f'{i}{transcript_cell+j+a5+shift}'] = expenses_val.iloc[j][y]/1000
    #     sheet_val[f'A{transcript_cell+j+a5+shift}'] = expenses_val.iat[j, 0]
    #     sheet_val[f'B{transcript_cell+j+a5+shift}'] = expenses_val.iat[j, 1]

    # for j in range(len(expenses_amd_Personnel)):
    #     sheet_val[f'{i}{260+j+a6+shift}'] = expenses_amd_Personnel.iloc[j][y]
    #     sheet_val[f'A{260+j+a6+shift}'] = expenses_amd_Personnel.iat[j, 0]
    #     sheet_val[f'B{260+j+a6+shift}'] = expenses_amd_Personnel.iat[j, 1]
    #
    # for j in range(len(expenses_amd_VGO)):
    #     sheet_val[f'{i}{260+j+a7+shift}'] = expenses_amd_VGO.iloc[j][y]
    #     sheet_val[f'A{260+j+a7+shift}'] = expenses_amd_VGO.iat[j, 0]
    #     sheet_val[f'B{260+j+a7+shift}'] = expenses_amd_VGO.iat[j, 1]
    #
    # for j in range(len(expenses_amd_Material)):
    #     sheet_val[f'{i}{260+j+a8+shift}'] = expenses_amd_Material.iloc[j][y]
    #     sheet_val[f'A{260+j+a8+shift}'] = expenses_amd_Material.iat[j, 0]
    #     sheet_val[f'B{260+j+a8+shift}'] = expenses_amd_Material.iat[j, 1]
    #
    # for j in range(len(expenses_amd_Amortization)):
    #     sheet_val[f'{i}{260+j+a9+shift}'] = expenses_amd_Amortization.iloc[j][y]
    #     sheet_val[f'A{260+j+a9+shift}'] = expenses_amd_Amortization.iat[j, 0]
    #     sheet_val[f'B{260+j+a9+shift}'] = expenses_amd_Amortization.iat[j, 1]
    #
    # for j in range(len(expenses_amd_Other_taxes)):
    #     sheet_val[f'{i}{260+j+a10+shift}'] = expenses_amd_Other_taxes.iloc[j][y]
    #     sheet_val[f'A{260+j+a10+shift}'] = expenses_amd_Other_taxes.iat[j, 0]
    #     sheet_val[f'B{260+j+a10+shift}'] = expenses_amd_Other_taxes.iat[j, 1]
    #
    # for j in range(len(expenses_amd_Taxes)):
    #     sheet_val[f'{i}{260+j+a11+shift}'] = expenses_amd_Taxes.iloc[j][y]
    #     sheet_val[f'A{260+j+a11+shift}'] = expenses_amd_Taxes.iat[j, 0]
    #     sheet_val[f'B{260+j+a11+shift}'] = expenses_amd_Taxes.iat[j, 1]
    #
    # for j in range(len(expenses_amd_Unknown)):
    #     sheet_val[f'{i}{260+j+a12+shift}'] = expenses_amd_Unknown.iloc[j][y]
    #     sheet_val[f'A{260+j+a12+shift}'] = expenses_amd_Unknown.iat[j, 0]
    #     sheet_val[f'B{260+j+a12+shift}'] = expenses_amd_Unknown.iat[j, 1]
    #
    # for j in range(len(bdr_sales_amd_otherAM)):
    #     sheet_val[f'{i}{260+a13+j+shift}'] = bdr_sales_amd_otherAM.iloc[j][y]
    #     sheet_val[f'A{260+a13+j+shift}'] = bdr_sales_amd_otherAM.iat[j, 0]
        # sheet_val[f'B{260+a13+j+shift}'] = bdr_sales_amd_otherAM.iat[j, 1]

    model_E.save('data_output/Model.xlsx')


