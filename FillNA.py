import pandas as pd
from openpyxl import load_workbook
from pandas import read_excel

# raw_data = read_excel('data_input/All.xlsx')
# raw_data.set_index(["Expense type"], inplace=True)
#
# print(raw_data.head(5))
# print(raw_data.columns)
#
#
# raw_data['Counter Party'] = raw_data['Counter Party'].fillna('no_data')
#
# print(raw_data.head(5))
# print(raw_data.columns)





def pasting_proekty(i, y):
    # читаем данные, устанавливаем expense id
    raw_data = read_excel('data_input/All.xlsx')
    raw_data.set_index(["Expense type"], inplace=True)
    print(raw_data.head(5))
    print(raw_data.columns)

    raw_data['Counter Party'] = raw_data['Counter Party'].fillna('no_data')
    print(raw_data.head(5))
    print(raw_data.columns)

    # # читаем форматирование расходов по годовой модели, устанавливаем expense id
    # formatting_expenses = read_excel('data_input/Format_expenses.xlsx')
    # formatting_expenses["Expense type"] = formatting_expenses["Статья бюджета"].astype(str)
    # formatting_expenses.set_index('Expense type', inplace=True)
    # # читаем форматирование отделов по годовой модели, устанавливаем division id
    # formatting_divisions = read_excel('data_input/Format_divisions.xlsx')
    # formatting_divisions["Division id"] = formatting_divisions["Division"].astype(str)
    # formatting_divisions.set_index('Division id', inplace=True)
    # # получаем отформатированный доступ к данным
    # data = raw_data.merge(formatting_expenses, left_index=True, right_index=True, how="outer")
    # data["Expense type_"] = data.index
    # data.set_index('Division', inplace=True)
    # data_divisions = data.merge(formatting_divisions, left_index=True, right_index=True, how='outer')
    # """ MESSES UP WITH EXPENSES!!! """
    # data_divisions.loc[(data_divisions['Transaction type'] == 'Поступление') & (data_divisions['Division'] == 'Отдел централизованных закупок'), 'Отдел'] = 'КД Проекты'
    # # выбираем данные по Company
    # data_proekty = data_divisions.loc[data_divisions['Отдел'] == 'КД Проекты']
    # # Выбираем и группируем доходы
    # data_proekty_sales = data_proekty.loc[data_proekty['Transaction type'] == 'Поступление']
    # data_proekty_sales2 = data_proekty_sales
    #
    # print(data_proekty_sales.head(5))
    # print(data_proekty_sales.columns)
    #
    # data_proekty_sales['Counter Party'] = data_proekty_sales2['Counter Party'].fillna('no_data')
    #
    # print(data_proekty_sales.head(5))
    # print(data_proekty_sales.columns)


pasting_proekty(1, 2)