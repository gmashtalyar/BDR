import pandas as pd
from openpyxl import load_workbook
from pandas import read_excel
from openpyxl.styles import Font


def pasting_tgn(i, y):
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
    data_tgn = data_divisions.loc[data_divisions['Отдел'] == 'КД ТГН']
    # Выбираем и группируем доходы
    data_tgn_sales = data_tgn.loc[data_tgn['Transaction type'] == 'Поступление']
    # Корректировка ВГО - переписать метод

    # Объединяем данные по Expense 1, смотрим по необходимости
    expenses_tgn = data_tgn.loc[data_tgn['Transaction type'] == 'Расходование']
    data_tgn_checking = expenses_tgn
    expenses_tgn = expenses_tgn.groupby(['Expense 1', 'Статья бюджета']).sum().reset_index()
    summa_tgn_sales = data_tgn_sales.groupby(['Expense type_', 'Counter Party']).sum().reset_index()
    sales_data = read_excel('data_input/Sales.xlsx', sheet_name='SVOD')

    TrubaKrug = sales_data.loc[(sales_data['Sales id'] == 'ТРУБЫ КРУГЛЫЕ') & (sales_data['Department'] == 'TGN')]
    TrubaProf = sales_data.loc[(sales_data['Sales id'] == 'ТРУБЫ ПРОФИЛЬНЫЕ') & (sales_data['Department'] == 'TGN')]
    List = sales_data.loc[(sales_data['Sales id'] == 'ЛИСТ') & (sales_data['Department'] == 'TGN')]
    Fason = sales_data.loc[(sales_data['Sales id'] == 'ФАСОН') & (sales_data['Department'] == 'TGN')]
    Sort = sales_data.loc[(sales_data['Sales id'] == 'СОРТ') & (sales_data['Department'] == 'TGN')]
    TrubaProch = sales_data.loc[(sales_data['Sales id'] == 'ТРУБА ПРОЧАЯ') & (sales_data['Department'] == 'TGN')]
    Prochee = sales_data.loc[(sales_data['Sales id'] == 'ПРОЧЕЕ ') & (sales_data['Department'] == 'TGN')]
    revenues = sales_data.loc[(sales_data['Sales id'] == 'Revenue') & (sales_data['Department'] == 'TGN')]
    income = sales_data.loc[(sales_data['Sales id'] == 'Income') & (sales_data['Department'] == 'TGN')]

    bdr_sales_tgn_Bonuses1 = summa_tgn_sales.loc[summa_tgn_sales['Expense type_'] == 'Ответ.хранение']
    bdr_sales_tgn_Bonuses2 = summa_tgn_sales.loc[summa_tgn_sales['Expense type_'] == 'Премия заводов']
    bdr_sales_tgn_LatePayments = summa_tgn_sales.loc[summa_tgn_sales['Expense type_'] == 'Процент за польз.средствами']
    bdr_sales_tgn_Other1 = summa_tgn_sales.loc[summa_tgn_sales['Expense type_'] == 'Юр-нотар услуги (доход)']
    bdr_sales_tgn_Other2 = summa_tgn_sales.loc[summa_tgn_sales['Expense type_'] == 'Списание задолжности (доход)']
    bdr_sales_tgn_Inventory = summa_tgn_sales.loc[summa_tgn_sales['Expense type_'] == 'Оприходование излишков (склад)']

    expenses_tgn_FOT = expenses_tgn.loc[expenses_tgn['Expense 1'] == 'ФОТ']
    expenses_tgn_ESN = expenses_tgn.loc[expenses_tgn['Expense 1'] == 'ЕСН']
    expenses_tgn_Personnel = expenses_tgn.loc[expenses_tgn['Expense 1'] == 'Прочие расходы на персонал']
    expenses_tgn_Material = expenses_tgn.loc[expenses_tgn['Expense 1'] == 'Материальные затраты']
    expenses_tgn_Sklad = expenses_tgn.loc[expenses_tgn['Expense 1'] == 'Услуги аутсорсеров (складские)']
    expenses_proekty_COGS = expenses_tgn.loc[expenses_tgn['Expense 1'] == 'СЕБЕСТОИМОСТЬ']
    expenses_tgn_Unknown = expenses_tgn.loc[expenses_tgn['Expense 1'] == '?!']
    expenses_tgn_nonfinancial = expenses_tgn.loc[expenses_tgn['Expense 1'] == 'non-financial']


    expense_check = expenses_tgn_FOT[y].sum() + expenses_tgn_ESN[y].sum() + expenses_tgn_Personnel[y].sum() +\
                    expenses_tgn_Material[y].sum() + expenses_tgn_Sklad[y].sum() + expenses_proekty_COGS[y].sum() +\
                    expenses_tgn_Unknown[y].sum() + expenses_tgn_nonfinancial[y].sum()

    print(f'Ошибка при переносе данных КД ТГН в {y} равна {round(expense_check, 2) - round(data_tgn_checking[y].sum(), 2)}')

    model_E = load_workbook('data_output/Model.xlsx')
    sheet_tgn = model_E['КД ТГН']

    sheet_tgn[f'{i}6'] = TrubaKrug[y].sum()
    sheet_tgn[f'{i}7'] = TrubaProf[y].sum()
    sheet_tgn[f'{i}8'] = List[y].sum()
    sheet_tgn[f'{i}9'] = Fason[y].sum()
    sheet_tgn[f'{i}10'] = Sort[y].sum()
    sheet_tgn[f'{i}11'] = TrubaProch[y].sum()
    sheet_tgn[f'{i}12'] = Prochee[y].sum()
    sheet_tgn[f'{i}21'] = revenues[y].sum()/1000
    sheet_tgn[f'{i}22'] = income[y].sum()/1000

    sheet_tgn[f'{i}23'] = bdr_sales_tgn_Bonuses1[y].sum()/1000 + bdr_sales_tgn_Bonuses2[y].sum()/1000  # бонусы и ох
    sheet_tgn[f'{i}25'] = bdr_sales_tgn_LatePayments[y].sum()/1000  # оплата просроченной ДБЗ
    sheet_tgn[f'{i}26'] = bdr_sales_tgn_Other1[y].sum()/1000 + bdr_sales_tgn_Other2[y].sum()/1000  # прочее
    sheet_tgn[f'{i}36'] = bdr_sales_tgn_Inventory[y].sum()/1000  # инвентаризация склада

    sheet_tgn[f'{i}49'] = -expenses_tgn_FOT[y].sum()/1000  # фот
    sheet_tgn[f'{i}50'] = -expenses_tgn_ESN[y].sum()/1000  # есн
    sheet_tgn[f'{i}51'] = -expenses_tgn_Personnel[y].sum()/1000  # прочие расходы на персонал
    sheet_tgn[f'{i}52'] = -expenses_tgn_Material[y].sum()/1000  # материальные затраты
    sheet_tgn[f'{i}59'] = -expenses_tgn_Sklad[y].sum()/1000  # услуги аутсорсеров (складские)
    # sheet_tgn[f'{i}250'] = -bdr_sales_tgn_Inventory[y].sum()/1000  # неизв

    shift = 1
    transcript_cell = 81

    sheet_tgn[f'A{transcript_cell-2}'] = 'По-статейные расшифровки:'
    sheet_tgn[f'A{transcript_cell-2}'].font = Font(underline='single', bold=True)
    sheet_tgn[f'A{transcript_cell-1}'] = 'Статья в Годовой модели'
    sheet_tgn[f'B{transcript_cell-1}'] = 'Статья в ERP'
    sheet_tgn[f'C{transcript_cell-1}'] = 'Примечание: в годовой модели расходы относятся к юр. лицу, а не ЦФО.'

    a3 = len(bdr_sales_tgn_Bonuses1) + shift
    a4 = a3 + len(bdr_sales_tgn_Bonuses2) + shift
    a5 = a4 + len(bdr_sales_tgn_Other1) + shift
    a6 = a5 + len(bdr_sales_tgn_Other2) + shift
    a7 = a6 + len(bdr_sales_tgn_Inventory) + shift
    a8 = a7 + len(expenses_tgn) + shift
    # a9 = a8 + len(expenses_amd_Material) + shift
    # a10 = a9 + len(expenses_amd_Amortization) + shift
    # a11 = a10 + len(expenses_amd_Other_taxes) + shift
    # a12 = a11 + len(expenses_amd_Taxes) + shift
    # a13 = a12 + len(expenses_amd_Unknown) + shift

    for j in range(len(bdr_sales_tgn_Bonuses1)):
        sheet_tgn[f'{i}{transcript_cell+j}'] = bdr_sales_tgn_Bonuses1.iloc[j][y]/1000
        sheet_tgn[f'A{transcript_cell+j}'] = bdr_sales_tgn_Bonuses1.iat[j, 0]
        sheet_tgn[f'B{transcript_cell+j}'] = bdr_sales_tgn_Bonuses1.iat[j, 1]

    for j in range(len(bdr_sales_tgn_Bonuses2)):
        sheet_tgn[f'{i}{transcript_cell+j+a3+shift}'] = bdr_sales_tgn_Bonuses2.iloc[j][y]/1000
        sheet_tgn[f'A{transcript_cell+j+a3+shift}'] = bdr_sales_tgn_Bonuses2.iat[j, 0]
        sheet_tgn[f'B{transcript_cell+j+a3+shift}'] = bdr_sales_tgn_Bonuses2.iat[j, 1]

    for j in range(len(bdr_sales_tgn_Other1)):
        sheet_tgn[f'{i}{transcript_cell+j+a4+shift}'] = bdr_sales_tgn_Other1.iloc[j][y]/1000
        sheet_tgn[f'A{transcript_cell+j+a4+shift}'] = bdr_sales_tgn_Other1.iat[j, 0]
        sheet_tgn[f'B{transcript_cell+j+a4+shift}'] = bdr_sales_tgn_Other1.iat[j, 1]

    for j in range(len(bdr_sales_tgn_Other2)):
        sheet_tgn[f'{i}{transcript_cell+j+a5+shift}'] = bdr_sales_tgn_Other2.iloc[j][y]/1000
        sheet_tgn[f'A{transcript_cell+j+a5+shift}'] = bdr_sales_tgn_Other2.iat[j, 0]
        sheet_tgn[f'B{transcript_cell+j+a5+shift}'] = bdr_sales_tgn_Other2.iat[j, 1]

    for j in range(len(bdr_sales_tgn_Inventory)):
        sheet_tgn[f'{i}{transcript_cell+j+a6+shift}'] = bdr_sales_tgn_Inventory.iloc[j][y]/1000
        sheet_tgn[f'A{transcript_cell+j+a6+shift}'] = bdr_sales_tgn_Inventory.iat[j, 0]
        sheet_tgn[f'B{transcript_cell+j+a6+shift}'] = bdr_sales_tgn_Inventory.iat[j, 1]

    for j in range(len(expenses_tgn)):
        sheet_tgn[f'{i}{transcript_cell+j+a7+shift}'] = expenses_tgn.iloc[j][y]/1000
        sheet_tgn[f'A{transcript_cell+j+a7+shift}'] = expenses_tgn.iat[j, 0]
        sheet_tgn[f'B{transcript_cell+j+a7+shift}'] = expenses_tgn.iat[j, 1]

    model_E.save('data_output/Model.xlsx')

