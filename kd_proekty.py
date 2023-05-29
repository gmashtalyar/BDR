import pandas as pd
from openpyxl import load_workbook
from pandas import read_excel
from openpyxl.styles import Font


def pasting_proekty(i, y):
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
    data_divisions.loc[(data_divisions['Transaction type'] == 'Поступление') & (data_divisions['Division'] == 'Отдел централизованных закупок'), 'Отдел'] = 'КД Проекты'
    # выбираем данные по Company
    data_proekty = data_divisions.loc[data_divisions['Отдел'] == 'КД Проекты']
    data_proekty_checking = data_proekty.loc[data_proekty['Transaction type'] == 'Расходование']

    # Выбираем и группируем доходы
    data_proekty_sales = data_proekty.loc[data_proekty['Transaction type'] == 'Поступление']

    # Корректировка ВГО - переписать метод
    data_proekty.loc[(data_proekty['Transaction type'] == 'Расходование') & (
                data_proekty['Counter Party'] == 'Вольфагролес'), 'Expense 1'] = 'Услуги СК ВАЛ (складские)'

    # Объединяем данные по Expense 1, смотрим по необходимости
    expenses_proekty = data_proekty.loc[data_proekty['Transaction type'] == 'Расходование']
    expenses_proekty = expenses_proekty.groupby(['Expense 1', 'Статья бюджета']).sum().reset_index()

    summa_proekty_sales = data_proekty_sales.groupby(['Expense type_', 'Counter Party']).sum().reset_index()
    sales_data = read_excel('data_input/Sales.xlsx', sheet_name='SVOD')


    TrubaKrug = sales_data.loc[(sales_data['Sales id'] == 'ТРУБЫ КРУГЛЫЕ') & (sales_data['Department'] == 'PROEKTY')]
    TrubaProf = sales_data.loc[(sales_data['Sales id'] == 'ТРУБЫ ПРОФИЛЬНЫЕ') & (sales_data['Department'] == 'PROEKTY')]
    List = sales_data.loc[(sales_data['Sales id'] == 'ЛИСТ') & (sales_data['Department'] == 'PROEKTY')]
    Fason = sales_data.loc[(sales_data['Sales id'] == 'ФАСОН') & (sales_data['Department'] == 'PROEKTY')]
    Sort = sales_data.loc[(sales_data['Sales id'] == 'СОРТ') & (sales_data['Department'] == 'PROEKTY')]
    TrubaProch = sales_data.loc[(sales_data['Sales id'] == 'ТРУБА ПРОЧАЯ') & (sales_data['Department'] == 'PROEKTY')]
    Prochee = sales_data.loc[(sales_data['Sales id'] == 'ПРОЧЕЕ ') & (sales_data['Department'] == 'PROEKTY')]
    revenues = sales_data.loc[(sales_data['Sales id'] == 'Revenue') & (sales_data['Department'] == 'PROEKTY')]
    income = sales_data.loc[(sales_data['Sales id'] == 'Income') & (sales_data['Department'] == 'PROEKTY')]

    bdr_sales_proekty_Bonuses1 = summa_proekty_sales.loc[summa_proekty_sales['Expense type_'] == 'Ответ.хранение']
    bdr_sales_proekty_Bonuses2 = summa_proekty_sales.loc[summa_proekty_sales['Expense type_'] == 'Премия заводов']
    bdr_sales_proekty_Corrections = summa_proekty_sales.loc[summa_proekty_sales['Expense type_'] == 'Корректировка поступлений']
    bdr_sales_proekty_LatePayments = summa_proekty_sales.loc[summa_proekty_sales['Expense type_'] == 'Процент за польз.средствами']
    bdr_sales_proekty_Other1 = summa_proekty_sales.loc[summa_proekty_sales['Expense type_'] == 'Возврат по затратным статьям']
    bdr_sales_proekty_Other2 = summa_proekty_sales.loc[summa_proekty_sales['Expense type_'] == 'Списание задолжности (доход)']

    expenses_proekty_FOT = expenses_proekty.loc[expenses_proekty['Expense 1'] == 'ФОТ']
    expenses_proekty_ESN = expenses_proekty.loc[expenses_proekty['Expense 1'] == 'ЕСН']
    expenses_proekty_Personnel = expenses_proekty.loc[expenses_proekty['Expense 1'] == 'Прочие расходы на персонал']
    expenses_proekty_Material = expenses_proekty.loc[expenses_proekty['Expense 1'] == 'Материальные затраты']
    expenses_proekty_VAL = expenses_proekty.loc[expenses_proekty['Expense 1'] == 'Услуги СК ВАЛ (складские)']
    expenses_proekty_COGS = expenses_proekty.loc[expenses_proekty['Expense 1'] == 'СЕБЕСТОИМОСТЬ']
    expenses_proekty_VGO = expenses_proekty.loc[expenses_proekty['Expense 1'] == 'Услуги ВГО']
    expenses_proekty_Unknown = expenses_proekty.loc[expenses_proekty['Expense 1'] == '?!']
    expenses_proekty_nonfinancial = expenses_proekty.loc[expenses_proekty['Expense 1'] == 'non-financial']

    expense_check = expenses_proekty_FOT[y].sum() + expenses_proekty_ESN[y].sum() + expenses_proekty_Personnel[y].sum()\
                    + expenses_proekty_Material[y].sum() + expenses_proekty_COGS[y].sum() + \
                    expenses_proekty_VAL[y].sum() + expenses_proekty_Unknown[y].sum() +\
                    expenses_proekty_nonfinancial[y].sum()

    print(f'Ошибка при переносе данных КД ПРОЕКТЫ в {y} равна {round(expense_check, 2) - round(data_proekty_checking[y].sum(), 2)}')


    model_E = load_workbook('data_output/Model.xlsx')
    sheet_proekty = model_E['КД ПРОЕКТЫ']

    sheet_proekty[f'{i}6'] = TrubaKrug[y].sum()
    sheet_proekty[f'{i}7'] = TrubaProf[y].sum()
    sheet_proekty[f'{i}8'] = List[y].sum()
    sheet_proekty[f'{i}9'] = Fason[y].sum()
    sheet_proekty[f'{i}10'] = Sort[y].sum()
    sheet_proekty[f'{i}11'] = TrubaProch[y].sum()
    sheet_proekty[f'{i}12'] = Prochee[y].sum()

    sheet_proekty[f'{i}21'] = revenues[y].sum()/1000
    sheet_proekty[f'{i}22'] = income[y].sum()/1000

    sheet_proekty[f'{i}23'] = bdr_sales_proekty_Bonuses1[y].sum()/1000 + bdr_sales_proekty_Bonuses2[y].sum()/1000  # бонусы и ох
    sheet_proekty[f'{i}24'] = bdr_sales_proekty_Corrections[y].sum()/1000  # корректировки
    sheet_proekty[f'{i}25'] = bdr_sales_proekty_LatePayments[y].sum()/1000  # оплата просроченной ДБЗ
    sheet_proekty[f'{i}26'] = bdr_sales_proekty_Other1[y].sum()/1000 + bdr_sales_proekty_Other2[y].sum()/1000  # прочее

    sheet_proekty[f'{i}49'] = -expenses_proekty_FOT[y].sum()/1000  # фот
    sheet_proekty[f'{i}50'] = -expenses_proekty_ESN[y].sum()/1000  # есн
    sheet_proekty[f'{i}51'] = -expenses_proekty_Personnel[y].sum()/1000  # прочие расходы на персонал
    sheet_proekty[f'{i}52'] = -expenses_proekty_Material[y].sum()/1000  # материальные затраты
    sheet_proekty[f'{i}58'] = -expenses_proekty_VAL[y].sum()/1000  # услуги СК ВАЛ


    shift = 1
    transcript_cell = 81

    sheet_proekty[f'A{transcript_cell-2}'] = 'По-статейные расшифровки:'
    sheet_proekty[f'A{transcript_cell-2}'].font = Font(underline='single', bold=True)
    sheet_proekty[f'A{transcript_cell-1}'] = 'Статья в Годовой модели'
    sheet_proekty[f'B{transcript_cell-1}'] = 'Статья в ERP'
    sheet_proekty[f'C{transcript_cell-1}'] = 'Примечание: в годовой модели расходы относятся к юр. лицу, а не ЦФО.'

    a3 = len(bdr_sales_proekty_Bonuses1) + shift
    a4 = a3 + len(bdr_sales_proekty_Bonuses2) + shift
    a5 = a4 + len(bdr_sales_proekty_Corrections) + shift
    a6 = a5 + len(bdr_sales_proekty_LatePayments) + shift
    a7 = a6 + len(bdr_sales_proekty_Other1) + shift
    a8 = a7 + len(bdr_sales_proekty_Other2) + shift
    a9 = a8 + len(expenses_proekty) + shift
    # a10 = a9 + len(expenses_amd_Amortization) + shift
    # a11 = a10 + len(expenses_amd_Other_taxes) + shift
    # a12 = a11 + len(expenses_amd_Taxes) + shift
    # a13 = a12 + len(expenses_amd_Unknown) + shift

    for j in range(len(bdr_sales_proekty_Bonuses1)):
        sheet_proekty[f'{i}{transcript_cell+j}'] = bdr_sales_proekty_Bonuses1.iloc[j][y]/1000
        sheet_proekty[f'A{transcript_cell+j}'] = bdr_sales_proekty_Bonuses1.iat[j, 0]
        sheet_proekty[f'B{transcript_cell+j}'] = bdr_sales_proekty_Bonuses1.iat[j, 1]

    for j in range(len(bdr_sales_proekty_Bonuses2)):
        sheet_proekty[f'{i}{transcript_cell+j+a3+shift}'] = bdr_sales_proekty_Bonuses2.iloc[j][y]/1000
        sheet_proekty[f'A{transcript_cell+j+a3+shift}'] = bdr_sales_proekty_Bonuses2.iat[j, 0]
        sheet_proekty[f'B{transcript_cell+j+a3+shift}'] = bdr_sales_proekty_Bonuses2.iat[j, 1]

    for j in range(len(bdr_sales_proekty_Corrections)):
        sheet_proekty[f'{i}{transcript_cell+j+a4+shift}'] = bdr_sales_proekty_Corrections.iloc[j][y]/1000
        sheet_proekty[f'A{transcript_cell+j+a4+shift}'] = bdr_sales_proekty_Corrections.iat[j, 0]
        sheet_proekty[f'B{transcript_cell+j+a4+shift}'] = bdr_sales_proekty_Corrections.iat[j, 1]

    for j in range(len(bdr_sales_proekty_LatePayments)):
        sheet_proekty[f'{i}{transcript_cell+j+a5+shift}'] = bdr_sales_proekty_LatePayments.iloc[j][y]/1000
        sheet_proekty[f'A{transcript_cell+j+a5+shift}'] = bdr_sales_proekty_LatePayments.iat[j, 0]
        sheet_proekty[f'B{transcript_cell+j+a5+shift}'] = bdr_sales_proekty_LatePayments.iat[j, 1]

    for j in range(len(bdr_sales_proekty_Other1)):
        sheet_proekty[f'{i}{transcript_cell+j+a6+shift}'] = bdr_sales_proekty_Other1.iloc[j][y]/1000
        sheet_proekty[f'A{transcript_cell+j+a6+shift}'] = bdr_sales_proekty_Other1.iat[j, 0]
        sheet_proekty[f'B{transcript_cell+j+a6+shift}'] = bdr_sales_proekty_Other1.iat[j, 1]

    for j in range(len(bdr_sales_proekty_Other2)):
        sheet_proekty[f'{i}{transcript_cell+j+a7+shift}'] = bdr_sales_proekty_Other2.iloc[j][y]/1000
        sheet_proekty[f'A{transcript_cell+j+a7+shift}'] = bdr_sales_proekty_Other2.iat[j, 0]
        sheet_proekty[f'B{transcript_cell+j+a7+shift}'] = bdr_sales_proekty_Other2.iat[j, 1]

    for j in range(len(expenses_proekty)):
        sheet_proekty[f'{i}{transcript_cell+j+a8+shift}'] = expenses_proekty.iloc[j][y]/1000
        sheet_proekty[f'A{transcript_cell+j+a8+shift}'] = expenses_proekty.iat[j, 0]
        sheet_proekty[f'B{transcript_cell+j+a8+shift}'] = expenses_proekty.iat[j, 1]


    model_E.save('data_output/Model.xlsx')

