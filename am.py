import pandas as pd
from openpyxl import load_workbook
from pandas import read_excel
from openpyxl.styles import Font


def pasting_am(i, y):
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
    data.set_index('Division', inplace=True)
    data_divisions = data.merge(formatting_divisions, left_index=True, right_index=True, how='outer')
    """ MESSES UP WITH EXPENSES!!! """
    data_divisions.loc[(data_divisions['Transaction type'] == 'Поступление') & (
            data_divisions['Division'] == 'Отдел централизованных закупок'), 'Отдел'] = 'КД Проекты'
    # выбираем данные по Company
    data_am = data_divisions.loc[data_divisions['Отдел'] == 'АМ']

    data_metally = data_divisions.loc[data_divisions['Company'] == 'Металлы']
    # Выбираем и группируем доходы
    data_am_sales = data_am.loc[data_am['Transaction type'] == 'Поступление']
    data_am_sales = data_am_sales.loc[
        (data_am_sales['Company'] == 'Ариэль Металл') | (data_am_sales['Company'] == 'Металлы')]
    data_am_sales.loc[data_am_sales['Counter Party'] == 'Вольфагролес', 'Expense type_'] = 'дубли'
    data_am_sales = data_am_sales.loc[
        (data_am_sales['Отдел'] != 'КД МСК') | (data_am_sales['Отдел'] != 'КД СПБ') | (
                data_am_sales['Отдел'] != 'КД Проекты')]

    # data_am_sales.to_excel('data_output/data_am_sales.xlsx')

    # Корректировка ВГО - переписать метод
    data_am.loc[(data_am['Transaction type'] == 'Расходование') & (
            data_am['Company'] == 'Вольфагролес'), 'Expense 1'] = 'Лишнее'
    data_am.loc[(data_am['Transaction type'] == 'Расходование') & (data_am['Company'] == 'ИП'), 'Expense 1'] = 'Лишнее'
    data_am.loc[(data_am['Transaction type'] == 'Расходование') & (data_am['Company'] == 'АМД'), 'Expense 1'] = 'Лишнее'
    data_am.loc[
        (data_am['Transaction type'] == 'Расходование') & (data_am['Company'] == 'АТМД'), 'Expense 1'] = 'Лишнее'
    data_am.loc[(data_am['Transaction type'] == 'Расходование') & (data_am['Company'] == 'Металлы') & (
            data_am['Division'] == 'Отдел транспортной логистики'), 'Expense 1'] = 'дубли'
    data_am.loc[(data_am['Transaction type'] == 'Расходование') & (data_am['Company'] == 'Металлы') & (
            data_am['Division'] == 'Склад Подольск'), 'Expense 1'] = 'дубли'

    # Объединяем данные по Expense 1, смотрим по необходимости
    expenses_am = data_am.loc[data_am['Transaction type'] == 'Расходование']
    data_am_checking = expenses_am
    expenses_am = expenses_am.groupby(['Expense 1', 'Статья бюджета']).sum().reset_index()
    summa_am_sales = data_am_sales.groupby(['Expense type_', 'Counter Party']).sum().reset_index()
    summa_am_sales_transcript = summa_am_sales.groupby('Expense type_').sum().reset_index()

    bdr_sales_amMetally_OfficeLease = summa_am_sales.loc[summa_am_sales['Expense type_'] == 'Аренда офиса (доход)']
    bdr_sales_amMetally_Marketing = summa_am_sales.loc[
        summa_am_sales['Expense type_'] == 'Компенсация расходов по интернет-проектам (доход)']
    bdr_sales_amMetally_Tax1 = summa_am_sales.loc[summa_am_sales['Expense type_'] == 'Налог в УК']
    bdr_sales_amMetally_Tax2 = summa_am_sales.loc[summa_am_sales['Expense type_'] == 'Налог в УК (доход)']
    bdr_sales_amMetally_OtherServices = summa_am_sales.loc[
        summa_am_sales['Expense type_'] == 'Услуги прочие (доход)']
    bdr_sales_amMetally_AntiCor = summa_am_sales.loc[summa_am_sales['Expense type_'] == 'Антикор Полимер (прибыль)']
    bdr_sales_amMetally_SKTK = summa_am_sales.loc[summa_am_sales['Expense type_'] == 'ППУ (прибыль)']
    bdr_sales_amMetally_InterestReceived = summa_am_sales.loc[summa_am_sales['Expense type_'] == 'Процент по займам']

    bdr_sales_am_LatePayments = summa_am_sales.loc[summa_am_sales['Expense type_'] == 'Процент за польз.средствами']
    bdr_sales_am_SomeRevenues = summa_am_sales.loc[
        (summa_am_sales['Expense type_'] == 'ВЫРУЧКА') & (summa_am_sales['Counter Party'] != 'Доход_логистика')]
    bdr_sales_am_ExpenseReturn = summa_am_sales.loc[summa_am_sales['Expense type_'] == 'Возврат по затратным статьям']
    bdr_sales_am_Inventory = summa_am_sales.loc[summa_am_sales['Expense type_'] == 'Оприходование излишков (склад)']
    bdr_sales_am_OH = summa_am_sales.loc[summa_am_sales['Expense type_'] == 'Ответ.хранение']
    bdr_sales_am_UslugiProchee = summa_am_sales.loc[summa_am_sales['Expense type_'] == 'Услуги прочие (доход)']
    bdr_sales_am_Legal = summa_am_sales.loc[summa_am_sales['Expense type_'] == 'Юр-нотар услуги (доход)']
    bdr_sales_am_LiquidityMgmt = summa_am_sales.loc[
        summa_am_sales['Expense type_'] == 'Банковское обслуживание (доход)']
    bdr_sales_am_BadDebitReturn = summa_am_sales.loc[summa_am_sales['Expense type_'] == 'Списание задолжности (доход)']
    bdr_sales_am_InterestReceived = summa_am_sales.loc[summa_am_sales['Expense type_'] == 'Процент по займам']

    expenses_am_FOT = expenses_am.loc[expenses_am['Expense 1'] == 'ФОТ']
    expenses_am_ESN = expenses_am.loc[expenses_am['Expense 1'] == 'ЕСН']
    expenses_am_Personnel = expenses_am.loc[expenses_am['Expense 1'] == 'Прочие расходы на персонал']
    expenses_am_Material = expenses_am.loc[expenses_am['Expense 1'] == 'Материальные затраты']
    expenses_am_OtherTaxes = expenses_am.loc[expenses_am['Expense 1'] == 'Прочие налоги уплаченные']
    expenses_am_Amortization = expenses_am.loc[expenses_am['Expense 1'] == 'Амортизация']
    expenses_am_InterestBank = expenses_am.loc[expenses_am['Expense 1'] == '%  уплаченный']
    expenses_am_Unknown = expenses_am.loc[expenses_am['Expense 1'] == '?!']
    expenses_am_InterestMetally = expenses_am.loc[expenses_am['Expense 1'] == 'Проценты']
    expenses_am_Dubli = expenses_am.loc[expenses_am['Expense 1'] == 'дубли']
    expenses_am_VGO = expenses_am.loc[expenses_am['Expense 1'] == 'Услуги ВГО']
    expenses_am_Skad_prochee = expenses_am.loc[expenses_am['Expense 1'] == 'Услуги аутсорсеров (складские)']
    expenses_am_COGS = expenses_am.loc[expenses_am['Expense 1'] == 'СЕБЕСТОИМОСТЬ']
    expenses_am_excessdata = expenses_am.loc[expenses_am['Expense 1'] == 'Лишнее']
    expenses_am_bdrincometax = expenses_am.loc[expenses_am['Expense 1'] == ' Налог на прибыль ']
    expenses_am_nonfinancial = expenses_am.loc[expenses_am['Expense 1'] == 'non-financial']
    expenses_am_bad_debts = expenses_am.loc[expenses_am['Expense 1'] == 'Списание ДЗ']




    expense_check = expenses_am_FOT[y].sum() + expenses_am_ESN[y].sum() + expenses_am_Personnel[y].sum() +\
                    expenses_am_Material[y].sum() + expenses_am_OtherTaxes[y].sum() + expenses_am_Amortization[y].sum()\
                    + expenses_am_InterestBank[y].sum() + expenses_am_Unknown[y].sum() + \
                    expenses_am_InterestMetally[y].sum() + expenses_am_Dubli[y].sum() + expenses_am_VGO[y].sum() + \
                    expenses_am_Skad_prochee[y].sum() + expenses_am_COGS[y].sum() + expenses_am_excessdata[y].sum() + \
                    expenses_am_bdrincometax[y].sum() + expenses_am_nonfinancial[y].sum() + expenses_am_bad_debts[y].sum()

    print(f'Ошибка при переносе данных АМ в {y} равна {round(expense_check, 2) - round(data_am_checking[y].sum(), 2)}')

    model_E = load_workbook('data_output/Model.xlsx')
    sheet_am = model_E['АМ']

    sheet_am[f'{i}39'] = bdr_sales_am_LatePayments[y].sum() / 1000  # оплата просроченной ДБЗ
    sheet_am[f'{i}40'] = bdr_sales_am_SomeRevenues[y].sum() / 1000 + bdr_sales_am_ExpenseReturn[y].sum() / 1000 + \
                         bdr_sales_am_Inventory[y].sum() / 1000 + bdr_sales_am_OH[y].sum() / 1000 + \
                         bdr_sales_am_UslugiProchee[y].sum() / 1000 \
                         + bdr_sales_am_Legal[y].sum() / 1000 + bdr_sales_amMetally_OfficeLease[y].sum() / 1000 + \
                         bdr_sales_amMetally_Marketing[y].sum() / 1000 \
                         + bdr_sales_amMetally_Tax1[y].sum() / 1000 + bdr_sales_amMetally_Tax2[y].sum() / 1000
    # + bdr_sales_amMetally_OtherServices[y].sum() / 1000  # прочие доходы
    sheet_am[f'{i}41'] = bdr_sales_am_LiquidityMgmt[y].sum() / 1000  # проценты на отстатки
    sheet_am[f'{i}145'] = bdr_sales_am_InterestReceived[y].sum()/1000 + bdr_sales_amMetally_AntiCor[y].sum()/1000 + \
                          bdr_sales_amMetally_SKTK[y].sum()/1000 + bdr_sales_amMetally_InterestReceived[y].sum()/1000  # проценты полученные (внешние контрагенты)
    sheet_am[f'{i}186'] = bdr_sales_am_BadDebitReturn[y].sum()/1000 + expenses_am_bad_debts[y].sum()/1000  # списадие ДЗ

    sheet_am[f'{i}97'] = -expenses_am_FOT[y].sum() / 1000  # фот
    sheet_am[f'{i}98'] = -expenses_am_ESN[y].sum() / 1000  # есн
    sheet_am[f'{i}99'] = -expenses_am_Personnel[y].sum() / 1000  # прочие расходы на персонал
    sheet_am[f'{i}100'] = -expenses_am_Material[y].sum() / 1000 + expenses_am_Skad_prochee[
        y].sum() / 1000  # материальные затраты
    sheet_am[f'{i}102'] = -expenses_am_OtherTaxes[y].sum() / 1000  # прочие налоги
    sheet_am[f'{i}103'] = -expenses_am_Amortization[y].sum() / 1000  # амортизация
    sheet_am[f'{i}118'] = -expenses_am_InterestBank[y].sum() / 1000  # проценты по кредитам
    # sheet_am[f'{i}250'] = -expenses_am_Unknown[y].sum() / 1000  # неизв
    # sheet_am[f'{i}251'] = -summa_am.iat[3, x] / 1000  # дубли
    # sheet_am[f'{i}252'] = -summa_am.iat[12, x] / 1000  # вго
    sheet_am[f'{i}253'] = -expenses_am_InterestMetally[y].sum() / 1000  # проценты?!

    shift = 1
    transcript_cell = 235

    sheet_am[f'A{transcript_cell-2}'] = 'По-статейные расшифровки:'
    sheet_am[f'A{transcript_cell-2}'].font = Font(underline='single', bold=True)
    sheet_am[f'A{transcript_cell-1}'] = 'Статья в Годовой модели'
    sheet_am[f'B{transcript_cell-1}'] = 'Статья в ERP'
    sheet_am[f'C{transcript_cell-1}'] = 'Примечание: в годовой модели расходы относятся к юр. лицу, а не ЦФО.'

    a3 = len(summa_am_sales_transcript) + shift
    a4 = a3 + len(expenses_am_FOT) + shift
    a6 = a4 + len(expenses_am_Personnel) + shift
    a7 = a6 + len(expenses_am_Material) + shift
    a8 = a7 + len(expenses_am_Unknown) + shift
    # a9 = a8 + len(expenses_am_Unknown) + shift
    # a10 = a9 + len(expenses_amd_Amortization) + shift
    # a11 = a10 + len(expenses_amd_Other_taxes) + shift
    # a12 = a11 + len(expenses_amd_Taxes) + shift
    # a13 = a12 + len(expenses_amd_Unknown) + shift

    for j in range(len(summa_am_sales_transcript)):
        sheet_am[f'{i}{transcript_cell+j}'] = summa_am_sales_transcript.iloc[j][y]/1000
        sheet_am[f'A{transcript_cell+j}'] = summa_am_sales_transcript.iat[j, 0]
        sheet_am[f'B{transcript_cell+j}'] = summa_am_sales_transcript.iat[j, 1]

    for j in range(len(expenses_am_FOT)):
        sheet_am[f'{i}{transcript_cell+j+a3+shift}'] = expenses_am_FOT.iloc[j][y]/1000
        sheet_am[f'A{transcript_cell+j+a3+shift}'] = expenses_am_FOT.iat[j, 0]
        sheet_am[f'B{transcript_cell+j+a3+shift}'] = expenses_am_FOT.iat[j, 1]

    for j in range(len(expenses_am_Personnel)):
        sheet_am[f'{i}{transcript_cell+j+a4+shift}'] = expenses_am_Personnel.iloc[j][y]/1000
        sheet_am[f'A{transcript_cell+j+a4+shift}'] = expenses_am_Personnel.iat[j, 0]
        sheet_am[f'B{transcript_cell+j+a4+shift}'] = expenses_am_Personnel.iat[j, 1]

    for j in range(len(expenses_am_Material)):
        sheet_am[f'{i}{transcript_cell+j+a6+shift}'] = expenses_am_Material.iloc[j][y]/1000
        sheet_am[f'A{transcript_cell+j+a6+shift}'] = expenses_am_Material.iat[j, 0]
        sheet_am[f'B{transcript_cell+j+a6+shift}'] = expenses_am_Material.iat[j, 1]

    # expenses_am_Unknown
    # expenses_am
    for j in range(len(expenses_am)):
        sheet_am[f'{i}{transcript_cell+j+a7+shift}'] = expenses_am.iloc[j][y]/1000
        sheet_am[f'A{transcript_cell+j+a7+shift}'] = expenses_am.iat[j, 0]
        sheet_am[f'B{transcript_cell+j+a7+shift}'] = expenses_am.iat[j, 1]

    # for j in range(len(expenses_am_Unknown)):
    #     sheet_am[f'{i}{transcript_cell+j+a8+shift}'] = expenses_am_Unknown.iloc[j][y]/1000
    #     sheet_am[f'A{transcript_cell+j+a8+shift}'] = expenses_am_Unknown.iat[j, 0]
    #     sheet_am[f'B{transcript_cell+j+a8+shift}'] = expenses_am_Unknown.iat[j, 1]

    # for j in range(len(expenses_am)):
    #     sheet_am[f'{i}{transcript_cell+j+a9+shift}'] = expenses_am.iloc[j][y]
    #     sheet_am[f'A{transcript_cell+j+a9+shift}'] = expenses_am.iat[j, 0]
    #     sheet_am[f'B{transcript_cell+j+a9+shift}'] = expenses_am.iat[j, 1]

    model_E.save('data_output/Model.xlsx')

