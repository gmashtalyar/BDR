import pandas as pd
from openpyxl import load_workbook
from pandas import read_excel
from openpyxl.styles import Font


def pasting_msk(i, y):
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
    # выбираем данные по Company
    data_msk = data_divisions.loc[data_divisions['Отдел'] == 'КД МСК']
    # Выбираем и группируем доходы
    data_msk_sales = data_msk.loc[data_msk['Transaction type'] == 'Поступление']
    # data_msk_sales.loc[(data_msk_sales['Transaction type'] == 'Поступление') & (data_msk_sales['Expense type_'] == 'Корректировка поступлений'), 'Counter Party'] = 'Unknown'
    # Корректировка ВГО - переписать метод
    data_msk.loc[(data_msk['Transaction type'] == 'Расходование') & (
                data_msk['Company'] == 'Вольфагролес'), 'Expense 1'] = 'Лишнее'
    data_msk.loc[(data_msk['Transaction type'] == 'Расходование') & (
                data_msk['Counter Party'] == 'Вольфагролес'), 'Expense 1'] = 'Услуги СК ВАЛ (складские)'
    data_msk.loc[(data_msk['Transaction type'] == 'Расходование') & (
                data_msk['Expense 1'] == 'Проценты'), 'Expense 1'] = 'Материальные затраты'



    # Объединяем данные по Expense 1, смотрим по необходимости
    expenses_msk = data_msk.loc[data_msk['Transaction type'] == 'Расходование']
    data_msk_checking = expenses_msk
    expenses_msk = expenses_msk.groupby(['Expense 1', 'Статья бюджета']).sum().reset_index()
    summa_msk_sales = data_msk_sales.groupby(['Expense type_', 'Counter Party']).sum().reset_index()
    sales_data = read_excel('data_input/Sales.xlsx', sheet_name='SVOD')

    TrubaKrug = sales_data.loc[(sales_data['Sales id'] == 'ТРУБЫ КРУГЛЫЕ') & (sales_data['Department'] == 'MSK')]
    TrubaProf = sales_data.loc[(sales_data['Sales id'] == 'ТРУБЫ ПРОФИЛЬНЫЕ') & (sales_data['Department'] == 'MSK')]
    List = sales_data.loc[(sales_data['Sales id'] == 'ЛИСТ') & (sales_data['Department'] == 'MSK')]
    Fason = sales_data.loc[(sales_data['Sales id'] == 'ФАСОН') & (sales_data['Department'] == 'MSK')]
    Sort = sales_data.loc[(sales_data['Sales id'] == 'СОРТ') & (sales_data['Department'] == 'MSK')]
    TrubaProch = sales_data.loc[(sales_data['Sales id'] == 'ТРУБА ПРОЧАЯ') & (sales_data['Department'] == 'MSK')]
    Prochee = sales_data.loc[(sales_data['Sales id'] == 'ПРОЧЕЕ ') & (sales_data['Department'] == 'MSK')]
    revenues = sales_data.loc[(sales_data['Sales id'] == 'Revenue') & (sales_data['Department'] == 'MSK')]
    income = sales_data.loc[(sales_data['Sales id'] == 'Income') & (sales_data['Department'] == 'MSK')]

    bdr_sales_msk_Bonuses1 = summa_msk_sales.loc[summa_msk_sales['Expense type_'] == 'Ответ.хранение']
    bdr_sales_msk_Bonuses2 = summa_msk_sales.loc[summa_msk_sales['Expense type_'] == 'Премия заводов']
    bdr_sales_msk_Corrections = summa_msk_sales.loc[summa_msk_sales['Expense type_'] == 'Корректировка поступлений']
    bdr_sales_msk_LatePayments = summa_msk_sales.loc[summa_msk_sales['Expense type_'] == 'Процент за польз.средствами']
    bdr_sales_msk_Other1 = summa_msk_sales.loc[summa_msk_sales['Expense type_'] == 'Юр-нотар услуги (доход)']
    bdr_sales_msk_Other2 = summa_msk_sales.loc[summa_msk_sales['Expense type_'] == 'Списание задолжности (доход)']

    expenses_msk_FOT = expenses_msk.loc[expenses_msk['Expense 1'] == 'ФОТ']
    expenses_msk_ESN = expenses_msk.loc[expenses_msk['Expense 1'] == 'ЕСН']
    expenses_msk_Personnel = expenses_msk.loc[expenses_msk['Expense 1'] == 'Прочие расходы на персонал']
    expenses_msk_Material = expenses_msk.loc[expenses_msk['Expense 1'] == 'Материальные затраты']
    expenses_msk_VAL = expenses_msk.loc[expenses_msk['Expense 1'] == 'Услуги СК ВАЛ (складские)']
    expenses_msk_Ppgt = expenses_msk.loc[expenses_msk['Expense 1'] == 'Услуги ППЖТ']
    expenses_msk_VGO = expenses_msk.loc[expenses_msk['Expense 1'] == 'Услуги ВГО']
    expenses_msk_COGS = expenses_msk.loc[expenses_msk['Expense 1'] == 'СЕБЕСТОИМОСТЬ']
    expenses_msk_Unknown = expenses_msk.loc[expenses_msk['Expense 1'] == '?!']
    expenses_msk_Interest = expenses_msk.loc[expenses_msk['Expense 1'] == 'Проценты']
    expenses_msk_nonfinancial = expenses_msk.loc[expenses_msk['Expense 1'] == 'non-financial']
    expenses_msk_bad_debts = expenses_msk.loc[expenses_msk['Expense 1'] == 'Списание ДЗ']

    expense_check = expenses_msk_FOT[y].sum() + expenses_msk_ESN[y].sum() + expenses_msk_Personnel[y].sum() +\
                    expenses_msk_Material[y].sum() + expenses_msk_VAL[y].sum() + expenses_msk_Ppgt[y].sum()\
                    + expenses_msk_VGO[y].sum() + expenses_msk_COGS[y].sum() + expenses_msk_Unknown[y].sum() + \
                    expenses_msk_nonfinancial[y].sum() + expenses_msk_bad_debts[y].sum()

    print(f'Ошибка при переносе данных КД МСК в {y} равна {round(expense_check, 2) - round(data_msk_checking[y].sum(), 2)}')

    model_E = load_workbook('data_output/Model.xlsx')
    sheet_msk = model_E['КД МСК']

    sheet_msk[f'{i}6'] = TrubaKrug[y].sum()
    sheet_msk[f'{i}7'] = TrubaProf[y].sum()
    sheet_msk[f'{i}8'] = List[y].sum()
    sheet_msk[f'{i}9'] = Fason[y].sum()
    sheet_msk[f'{i}10'] = Sort[y].sum()
    sheet_msk[f'{i}11'] = TrubaProch[y].sum()
    sheet_msk[f'{i}12'] = Prochee[y].sum()
    sheet_msk[f'{i}21'] = revenues[y].sum()/1000
    sheet_msk[f'{i}22'] = income[y].sum()/1000

    sheet_msk[f'{i}23'] = bdr_sales_msk_Bonuses1[y].sum()/1000 + bdr_sales_msk_Bonuses2[y].sum()/1000  # бонусы и ох
    sheet_msk[f'{i}24'] = bdr_sales_msk_Corrections[y].sum()/1000  # корректировки
    sheet_msk[f'{i}25'] = bdr_sales_msk_LatePayments[y].sum()/1000  # оплата просроченной ДБЗ
    sheet_msk[f'{i}26'] = bdr_sales_msk_Other1[y].sum()/1000 + bdr_sales_msk_Other2[y].sum()/1000 + \
                          expenses_msk_bad_debts[y].sum()/1000 # прочее

    sheet_msk[f'{i}49'] = -expenses_msk_FOT[y].sum()/1000  # фот
    sheet_msk[f'{i}50'] = -expenses_msk_ESN[y].sum()/1000  # есн
    sheet_msk[f'{i}51'] = -expenses_msk_Personnel[y].sum()/1000  # прочие расходы на персонал
    sheet_msk[f'{i}52'] = -expenses_msk_Material[y].sum()/1000  # материальные затраты
    sheet_msk[f'{i}54'] = -expenses_msk_VGO[y].sum()/1000  # вго
    sheet_msk[f'{i}58'] = -expenses_msk_VAL[y].sum()/1000  # услуги СК ВАЛ
    sheet_msk[f'{i}60'] = -expenses_msk_Ppgt[y].sum()/1000  # ппжт
    # sheet_msk[f'{i}250'] = -expenses_msk_Unknown[y].sum()/1000  # неизв
    # sheet_msk[f'{i}252'] = -expenses_msk_Interest[y].sum()/1000  # проценты

    shift = 1
    transcript_cell = 81

    sheet_msk[f'A{transcript_cell-2}'] = 'По-статейные расшифровки:'
    sheet_msk[f'A{transcript_cell-2}'].font = Font(underline='single', bold=True)
    sheet_msk[f'A{transcript_cell-1}'] = 'Статья в Годовой модели'
    sheet_msk[f'B{transcript_cell-1}'] = 'Статья в ERP'
    sheet_msk[f'C{transcript_cell-1}'] = 'Примечание: в годовой модели расходы относятся к юр. лицу, а не ЦФО.'

    a3 = len(bdr_sales_msk_Bonuses1) + shift
    a4 = a3 + len(bdr_sales_msk_Bonuses2) + shift
    a5 = a4 + len(bdr_sales_msk_LatePayments) + shift
    a6 = a5 + len(bdr_sales_msk_Other1) + shift
    a7 = a6 + len(expenses_msk) + shift
    # a8 = a7 + len(expenses_spb) + shift
    # a9 = a8 + len(expenses_amd_Material) + shift
    # a10 = a9 + len(expenses_amd_Amortization) + shift
    # a11 = a10 + len(expenses_amd_Other_taxes) + shift
    # a12 = a11 + len(expenses_amd_Taxes) + shift
    # a13 = a12 + len(expenses_amd_Unknown) + shift

    for j in range(len(bdr_sales_msk_Bonuses1)):
        sheet_msk[f'{i}{transcript_cell+j}'] = bdr_sales_msk_Bonuses1.iloc[j][y]/1000
        sheet_msk[f'A{transcript_cell+j}'] = bdr_sales_msk_Bonuses1.iat[j, 0]
        sheet_msk[f'B{transcript_cell+j}'] = bdr_sales_msk_Bonuses1.iat[j, 1]

    for j in range(len(bdr_sales_msk_Bonuses2)):
        sheet_msk[f'{i}{transcript_cell+j+a3+shift}'] = bdr_sales_msk_Bonuses2.iloc[j][y]/1000
        sheet_msk[f'A{transcript_cell+j+a3+shift}'] = bdr_sales_msk_Bonuses2.iat[j, 0]
        sheet_msk[f'B{transcript_cell+j+a3+shift}'] = bdr_sales_msk_Bonuses2.iat[j, 1]

    for j in range(len(bdr_sales_msk_LatePayments)):
        sheet_msk[f'{i}{transcript_cell+j+a4+shift}'] = bdr_sales_msk_LatePayments.iloc[j][y]/1000
        sheet_msk[f'A{transcript_cell+j+a4+shift}'] = bdr_sales_msk_LatePayments.iat[j, 0]
        sheet_msk[f'B{transcript_cell+j+a4+shift}'] = bdr_sales_msk_LatePayments.iat[j, 1]

    for j in range(len(bdr_sales_msk_Other1)):
        sheet_msk[f'{i}{transcript_cell+j+a5+shift}'] = bdr_sales_msk_Other1.iloc[j][y]/1000
        sheet_msk[f'A{transcript_cell+j+a5+shift}'] = bdr_sales_msk_Other1.iat[j, 0]
        sheet_msk[f'B{transcript_cell+j+a5+shift}'] = bdr_sales_msk_Other1.iat[j, 1]

    for j in range(len(expenses_msk)):
        sheet_msk[f'{i}{transcript_cell+j+a6+shift}'] = expenses_msk.iloc[j][y]/1000
        sheet_msk[f'A{transcript_cell+j+a6+shift}'] = expenses_msk.iat[j, 0]
        sheet_msk[f'B{transcript_cell+j+a6+shift}'] = expenses_msk.iat[j, 1]

    # for j in range(len(expenses_spb)):
    #     sheet_spb[f'{i}{260+j+a7+shift}'] = expenses_spb.iloc[j][y]
    #     sheet_spb[f'A{260+j+a7+shift}'] = expenses_spb.iat[j, 0]
    #     sheet_spb[f'B{260+j+a7+shift}'] = expenses_spb.iat[j, 1]

    model_E.save('data_output/Model.xlsx')
