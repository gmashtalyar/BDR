import pandas as pd
from openpyxl import load_workbook
from pandas import read_excel


def formatting_KD(YYYYYY, XXXXXXXX):
    #Первичная обработка данных
    raw_data = read_excel('data_input/All_qs.xlsx')  # читаем данные
    raw_data.set_index(["Статья бюджета"], inplace=True)  # устанавливаем id "Статья бюджета"
    raw_data['Контрагент'] = raw_data['Контрагент'].fillna('нд')  # устанавливаем пустых контрагентов
    raw_data = raw_data.loc[raw_data['Статья оборотов.Направление'] == 'Расходование']  # удаляем данные по поступлениям

    # читаем форматирование расходов по годовой модели, устанавливаем expense id
    formatting_expenses = read_excel('data_input/Format_expenses.xlsx')
    formatting_expenses.set_index('Статья бюджета', inplace=True)
    # читаем форматирование отделов по годовой модели, устанавливаем division id
    formatting_divisions = read_excel('data_input/Format_divisions.xlsx')
    formatting_divisions["ЦФО"] = formatting_divisions["Division"].astype(str)
    formatting_divisions.set_index('ЦФО', inplace=True)


    # получаем отформатированный доступ к данным
    data = raw_data.merge(formatting_expenses, left_index=True, right_index=True, how="outer")
    data["Статья бюджета"] = data.index
    data.set_index('ЦФО', inplace=True)
    data_divisions = data.merge(formatting_divisions, left_index=True, right_index=True, how='outer')


    # уточняем форматирование отделов и расходов вручную
    data_divisions.loc[(data_divisions['Организации'] == 'АТМД'), 'Отдел'] = 'АТМД'
    data_divisions.loc[(data_divisions['Организации'] == 'АМД'), 'Отдел'] = 'АМД'
    data_divisions.loc[(data_divisions['Организации'] == 'ИП'), 'Отдел'] = 'АМД'
    data_divisions.loc[(data_divisions['Организации'] == 'Металлы') & (
            data_divisions['Контрагент'] == 'Вольфагролес'), 'Отдел'] = 'АМД'
    data_divisions.loc[(data_divisions['Организации'] == 'Вольфагролес'), 'Отдел'] = 'ВАЛ'

    data_divisions.loc[(data_divisions['Организации'] == 'Ариэль Металл') & (
            data_divisions['Контрагент'] == 'Вольфагролес') & (
            data_divisions['Отдел'] != 'АМ'), 'Expense 1'] = 'Услуги СК ВАЛ (складские)'
    data_divisions.loc[(data_divisions['Организации'] == 'Ариэль Металл') & (
            data_divisions['Контрагент'] == 'Вольфагролес') & (
            data_divisions['Отдел'] == 'АМ'), 'Expense 1'] = 'Услуги ВГО'

    data_divisions = data_divisions.loc[(data_divisions['Expense 1'] != 'СЕБЕСТОИМОСТЬ')]
    data_divisions = data_divisions.loc[(data_divisions['Отдел'] == YYYYYY)]


    data_divisions['Январь 2022'] = data_divisions['Январь 2022'].apply(abs)
    data_divisions['Февраль 2022'] = data_divisions['Февраль 2022'].apply(abs)
    data_divisions['Март 2022'] = data_divisions['Март 2022'].apply(abs)
    data_divisions['Апрель 2022'] = data_divisions['Апрель 2022'].apply(abs)
    data_divisions['Май 2022'] = data_divisions['Май 2022'].apply(abs)
    data_divisions['Июнь 2022'] = data_divisions['Июнь 2022'].apply(abs)
    data_divisions['Июль 2022'] = data_divisions['Июль 2022'].apply(abs)
    data_divisions['Август 2022'] = data_divisions['Август 2022'].apply(abs)
    data_divisions['Сентябрь 2022'] = data_divisions['Сентябрь 2022'].apply(abs)



    # делаем удельные расходы на тн.
    data_divisions_tons = data_divisions
    sales_data = read_excel('data_input/Sales_new.xlsx', sheet_name='SVOD')

    total_tons_sold = sales_data.loc[(sales_data['Sales id'] == 'Total tons sold') & (sales_data['Department'] == XXXXXXXX)]

    tons_01 = total_tons_sold['January'].sum()
    tons_02 = total_tons_sold['February'].sum()
    tons_03 = total_tons_sold['March'].sum()
    tons_04 = total_tons_sold['April'].sum()
    tons_05 = total_tons_sold['May'].sum()
    tons_06 = total_tons_sold['June'].sum()
    tons_07 = total_tons_sold['July'].sum()
    tons_08 = total_tons_sold['August'].sum()
    tons_09 = total_tons_sold['September'].sum()


    data_divisions['Январь 2022 Т'] = data_divisions_tons['Январь 2022'].div(tons_01).round(2).apply(abs)
    data_divisions['Февраль 2022 Т'] = data_divisions_tons['Февраль 2022'].div(tons_02).round(2).apply(abs)
    data_divisions['Март 2022 Т'] = data_divisions_tons['Март 2022'].div(tons_03).round(2).apply(abs)
    data_divisions['Апрель 2022 Т'] = data_divisions_tons['Апрель 2022'].div(tons_04).round(2).apply(abs)
    data_divisions['Май 2022 Т'] = data_divisions_tons['Май 2022'].div(tons_05).round(2).apply(abs)
    data_divisions['Июнь 2022 Т'] = data_divisions_tons['Июнь 2022'].div(tons_06).round(2).apply(abs)
    data_divisions['Июль 2022 Т'] = data_divisions_tons['Июль 2022'].div(tons_07).round(2).apply(abs)
    data_divisions['Август 2022 Т'] = data_divisions_tons['Август 2022'].div(tons_08).round(2).apply(abs)
    data_divisions['Сентябрь 2022 Т'] = data_divisions_tons['Сентябрь 2022'].div(tons_09).round(2).apply(abs)


    data_divisions = data_divisions[data_divisions.index.notnull()]


    # data_divisions.to_excel('data_output/data_transpose.xlsx')

    data_divisions.to_excel(f'data_output/data_divisions_{XXXXXXXX}.xlsx')

    data_divisions = read_excel(f'data_output/data_divisions_{XXXXXXXX}.xlsx')  # читаем данные

    model_E = load_workbook(f'data_output/data_divisions_{XXXXXXXX}.xlsx')
    sheet = model_E['Sheet1']
    for j in range(len(data_divisions)):
        sheet[f'AB{2+j}'] = j
    sheet[f'AB1'] = 'unique id'
    model_E.save(f'data_output/data_divisions_w_id_{XXXXXXXX}.xlsx')
    """
    
    """
    data_divisions = read_excel(f'data_output/data_divisions_w_id_{XXXXXXXX}.xlsx')  # читаем данные


    data_melted_1 = data_divisions.melt(id_vars=[
        'unique id', 'Организации', 'Expense 1', 'Статья бюджета', 'Отдел'], value_vars=[
        'Январь 2022', 'Февраль 2022', 'Март 2022', 'Апрель 2022', 'Май 2022', 'Июнь 2022', 'Июль 2022', 'Август 2022',
        'Сентябрь 2022'])

    data_melted_2 = data_divisions.melt(id_vars=[
        'unique id', 'Организации', 'Expense 1', 'Статья бюджета', 'Отдел'], value_vars=[
        'Январь 2022 Т', 'Февраль 2022 Т', 'Март 2022 Т', 'Апрель 2022 Т', 'Май 2022 Т', 'Июнь 2022 Т', 'Июль 2022 Т',
        'Август 2022 Т', 'Сентябрь 2022 Т'])


    data_melted_1.to_excel(f'data_output/data_divisions_melted_1_{XXXXXXXX}.xlsx')

    data_melted_2.to_excel(f'data_output/data_divisions_melted_2_{XXXXXXXX}.xlsx')

    data_melted_1.rename(columns={'variable': 'Месяц', 'value': 'Руб'}, inplace=True)
    data_melted_1['РубТН'] = data_melted_2['value']
    data_melted_1.to_excel(f'data_output/data_divisions_melted_{XXXXXXXX}.xlsx')



formatting_KD('КД МСК', 'MSK')
formatting_KD('КД СПБ', 'SPB')
formatting_KD('КД ТГН', 'TGN')
formatting_KD('КД СМР', 'SMR')