import pandas as pd
from openpyxl import load_workbook
from pandas import read_excel
from openpyxl.styles import Font


def pasting_spb(i, y):
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
    data_spb = data_divisions.loc[data_divisions['Отдел'] == 'КД СПБ']
    # Выбираем и группируем доходы
    data_spb_sales = data_spb.loc[data_spb['Transaction type'] == 'Поступление']
    data_spb_sales.loc[(data_spb_sales['Transaction type'] == 'Поступление') & (data_spb_sales['Expense type_'] == 'Корректировка поступлений'), 'Counter Party'] = 'NA'
    # Корректировка ВГО - переписать метод
    data_spb.loc[
        (data_spb['Transaction type'] == 'Расходование') & (data_spb['Company'] == 'АТМД'), 'Expense 1'] = 'Лишнее'
    # Объединяем данные по Expense 1, смотрим по необходимости
    expenses_spb = data_spb.loc[data_spb['Transaction type'] == 'Расходование']
    data_spb_checking = expenses_spb
    expenses_spb = expenses_spb.groupby(['Expense 1', 'Статья бюджета']).sum().reset_index()
    summa_spb_sales = data_spb_sales.groupby(['Expense type_', 'Counter Party']).sum().reset_index()
    sales_data = read_excel('data_input/Sales.xlsx', sheet_name='SVOD')

    TrubaKrug = sales_data.loc[(sales_data['Sales id'] == 'ТРУБЫ КРУГЛЫЕ') & (sales_data['Department'] == 'SPB')]
    TrubaProf = sales_data.loc[(sales_data['Sales id'] == 'ТРУБЫ ПРОФИЛЬНЫЕ') & (sales_data['Department'] == 'SPB')]
    List = sales_data.loc[(sales_data['Sales id'] == 'ЛИСТ') & (sales_data['Department'] == 'SPB')]
    Fason = sales_data.loc[(sales_data['Sales id'] == 'ФАСОН') & (sales_data['Department'] == 'SPB')]
    Sort = sales_data.loc[(sales_data['Sales id'] == 'СОРТ') & (sales_data['Department'] == 'SPB')]
    TrubaProch = sales_data.loc[(sales_data['Sales id'] == 'ТРУБА ПРОЧАЯ') & (sales_data['Department'] == 'SPB')]
    Prochee = sales_data.loc[(sales_data['Sales id'] == 'ПРОЧЕЕ ') & (sales_data['Department'] == 'SPB')]
    revenues = sales_data.loc[(sales_data['Sales id'] == 'Revenue') & (sales_data['Department'] == 'SPB')]
    income = sales_data.loc[(sales_data['Sales id'] == 'Income') & (sales_data['Department'] == 'SPB')]

    bdr_sales_spb_Bonuses1 = summa_spb_sales.loc[summa_spb_sales['Expense type_'] == 'Ответ.хранение']
    bdr_sales_spb_Bonuses2 = summa_spb_sales.loc[summa_spb_sales['Expense type_'] == 'Премия заводов']
    bdr_sales_spb_Corrections = summa_spb_sales.loc[summa_spb_sales['Expense type_'] == 'Корректировка поступлений']
    bdr_sales_spb_LatePayments = summa_spb_sales.loc[summa_spb_sales['Expense type_'] == 'Процент за польз.средствами']
    bdr_sales_spb_Other1 = summa_spb_sales.loc[summa_spb_sales['Expense type_'] == 'Юр-нотар услуги (доход)']
    bdr_sales_spb_Other2 = summa_spb_sales.loc[summa_spb_sales['Expense type_'] == 'Списание задолжности (доход)']

    expenses_spb_FOT = expenses_spb.loc[expenses_spb['Expense 1'] == 'ФОТ']
    expenses_spb_ESN = expenses_spb.loc[expenses_spb['Expense 1'] == 'ЕСН']
    expenses_spb_Personnel = expenses_spb.loc[expenses_spb['Expense 1'] == 'Прочие расходы на персонал']
    expenses_spb_Material = expenses_spb.loc[expenses_spb['Expense 1'] == 'Материальные затраты']
    expenses_spb_Amortization = expenses_spb.loc[expenses_spb['Expense 1'] == 'Амортизация']
    expenses_spb_Sklad = expenses_spb.loc[expenses_spb['Expense 1'] == 'Услуги аутсорсеров (складские)']
    expenses_spb_Ppgt = expenses_spb.loc[expenses_spb['Expense 1'] == 'Услуги ППЖТ']
    expenses_spb_OtherOther = expenses_spb.loc[expenses_spb['Expense 1'] == 'прочее (вычеты со знаком минус)']
    expenses_spb_COGS = expenses_spb.loc[expenses_spb['Expense 1'] == 'СЕБЕСТОИМОСТЬ']
    expenses_spb_Unknown = expenses_spb.loc[expenses_spb['Expense 1'] == '?!']
    expenses_spb_nonfinancial = expenses_spb.loc[expenses_spb['Expense 1'] == 'non-financial']
    expenses_spb_bad_debts = expenses_spb.loc[expenses_spb['Expense 1'] == 'Списание ДЗ']



    expense_check = expenses_spb_FOT[y].sum() + expenses_spb_ESN[y].sum() + expenses_spb_Personnel[y].sum() +\
                    expenses_spb_Material[y].sum() + expenses_spb_Amortization[y].sum() + expenses_spb_Sklad[y].sum()\
                    + expenses_spb_Ppgt[y].sum() + expenses_spb_OtherOther[y].sum() + expenses_spb_COGS[y].sum()+ \
                    expenses_spb_Unknown[y].sum() + expenses_spb_nonfinancial[y].sum() + expenses_spb_bad_debts[y].sum()

    print(f'Ошибка при переносе данных КД СПБ в {y} равна {round(expense_check, 2) - round(data_spb_checking[y].sum(), 2)}')

    model_E = load_workbook('data_output/Model.xlsx')
    sheet_spb = model_E['КД СПБ']

    sheet_spb[f'{i}6'] = TrubaKrug[y].sum()
    sheet_spb[f'{i}7'] = TrubaProf[y].sum()
    sheet_spb[f'{i}8'] = List[y].sum()
    sheet_spb[f'{i}9'] = Fason[y].sum()
    sheet_spb[f'{i}10'] = Sort[y].sum()
    sheet_spb[f'{i}11'] = TrubaProch[y].sum()
    sheet_spb[f'{i}12'] = Prochee[y].sum()
    sheet_spb[f'{i}21'] = revenues[y].sum()/1000
    sheet_spb[f'{i}22'] = income[y].sum()/1000

    sheet_spb[f'{i}23'] = bdr_sales_spb_Bonuses1[y].sum()/1000 + bdr_sales_spb_Bonuses2[y].sum()/1000  # бонусы и ох
    sheet_spb[f'{i}24'] = bdr_sales_spb_Corrections[y].sum()/1000  # корректировки
    sheet_spb[f'{i}25'] = bdr_sales_spb_LatePayments[y].sum()/1000  # оплата просроченной ДБЗ
    sheet_spb[f'{i}26'] = bdr_sales_spb_Other1[y].sum()/1000 + bdr_sales_spb_Other2[y].sum()/1000 +\
                          expenses_spb_OtherOther[y].sum()/1000 + expenses_spb_bad_debts[y].sum()/1000  # прочее

    sheet_spb[f'{i}49'] = -expenses_spb_FOT[y].sum()/1000  # фот
    sheet_spb[f'{i}50'] = -expenses_spb_ESN[y].sum()/1000  # есн
    sheet_spb[f'{i}51'] = -expenses_spb_Personnel[y].sum()/1000  # прочие расходы на персонал
    sheet_spb[f'{i}52'] = -expenses_spb_Material[y].sum()/1000  # материальные затраты
    sheet_spb[f'{i}56'] = -expenses_spb_Amortization[y].sum()/1000  # амортизация
    sheet_spb[f'{i}59'] = -expenses_spb_Sklad[y].sum()/1000  # услуги аутсорсеров (складские)
    sheet_spb[f'{i}60'] = -expenses_spb_Ppgt[y].sum()/1000  # ппжт
    # sheet_spb[f'{i}250'] = -expenses_spb_Unknown[y].sum()/1000  # неизв
    # sheet_spb[f'{i}251'] = -bdr_sales_spb_LatePayments[y].sum()/1000  # лишнее

    shift = 1
    transcript_cell = 81

    sheet_spb[f'A{transcript_cell-2}'] = 'По-статейные расшифровки:'
    sheet_spb[f'A{transcript_cell-2}'].font = Font(underline='single', bold=True)
    sheet_spb[f'A{transcript_cell-1}'] = 'Статья в Годовой модели'
    sheet_spb[f'B{transcript_cell-1}'] = 'Статья в ERP'
    sheet_spb[f'C{transcript_cell-1}'] = 'Примечание: в годовой модели расходы относятся к юр. лицу, а не ЦФО.'

    a3 = len(bdr_sales_spb_Bonuses1) + shift
    a4 = a3 + len(bdr_sales_spb_Bonuses2) + shift
    a5 = a4 + len(bdr_sales_spb_LatePayments) + shift
    a6 = a5 + len(bdr_sales_spb_Other1) + shift
    a7 = a6 + len(expenses_spb) + shift
    # a8 = a7 + len(expenses_spb) + shift
    # a9 = a8 + len(expenses_amd_Material) + shift
    # a10 = a9 + len(expenses_amd_Amortization) + shift
    # a11 = a10 + len(expenses_amd_Other_taxes) + shift
    # a12 = a11 + len(expenses_amd_Taxes) + shift
    # a13 = a12 + len(expenses_amd_Unknown) + shift

    for j in range(len(bdr_sales_spb_Bonuses1)):
        sheet_spb[f'{i}{transcript_cell+j}'] = bdr_sales_spb_Bonuses1.iloc[j][y]/1000
        sheet_spb[f'A{transcript_cell+j}'] = bdr_sales_spb_Bonuses1.iat[j, 0]
        sheet_spb[f'B{transcript_cell+j}'] = bdr_sales_spb_Bonuses1.iat[j, 1]

    for j in range(len(bdr_sales_spb_Bonuses2)):
        sheet_spb[f'{i}{transcript_cell+j+a3+shift}'] = bdr_sales_spb_Bonuses2.iloc[j][y]/1000
        sheet_spb[f'A{transcript_cell+j+a3+shift}'] = bdr_sales_spb_Bonuses2.iat[j, 0]
        sheet_spb[f'B{transcript_cell+j+a3+shift}'] = bdr_sales_spb_Bonuses2.iat[j, 1]

    for j in range(len(bdr_sales_spb_LatePayments)):
        sheet_spb[f'{i}{transcript_cell+j+a4+shift}'] = bdr_sales_spb_LatePayments.iloc[j][y]/1000
        sheet_spb[f'A{transcript_cell+j+a4+shift}'] = bdr_sales_spb_LatePayments.iat[j, 0]
        sheet_spb[f'B{transcript_cell+j+a4+shift}'] = bdr_sales_spb_LatePayments.iat[j, 1]

    for j in range(len(bdr_sales_spb_Other1)):
        sheet_spb[f'{i}{transcript_cell+j+a5+shift}'] = bdr_sales_spb_Other1.iloc[j][y]/1000
        sheet_spb[f'A{transcript_cell+j+a5+shift}'] = bdr_sales_spb_Other1.iat[j, 0]
        sheet_spb[f'B{transcript_cell+j+a5+shift}'] = bdr_sales_spb_Other1.iat[j, 1]

    for j in range(len(expenses_spb)):
        sheet_spb[f'{i}{transcript_cell+j+a6+shift}'] = expenses_spb.iloc[j][y]/1000
        sheet_spb[f'A{transcript_cell+j+a6+shift}'] = expenses_spb.iat[j, 0]
        sheet_spb[f'B{transcript_cell+j+a6+shift}'] = expenses_spb.iat[j, 1]

    # for j in range(len(expenses_spb)):
    #     sheet_spb[f'{i}{260+j+a7+shift}'] = expenses_spb.iloc[j][y]
    #     sheet_spb[f'A{260+j+a7+shift}'] = expenses_spb.iat[j, 0]
    #     sheet_spb[f'B{260+j+a7+shift}'] = expenses_spb.iat[j, 1]

    model_E.save('data_output/Model.xlsx')

