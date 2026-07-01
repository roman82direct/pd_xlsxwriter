import xlsxwriter

workbook = xlsxwriter.Workbook('доли.xlsx')
worksheet = workbook.add_worksheet('Структура')

# Данные
categories = ['Электроника', 'Мебель', 'Одежда', 'Книги', 'Другое']
values = [450, 280, 320, 180, 70]

worksheet.write_row('A1', ['Категория', 'Продажи, тыс. руб.'])
worksheet.write_column('A2', categories)
worksheet.write_column('B2', values)

# Создаём круговую диаграмму
chart = workbook.add_chart({'type': 'pie'})

chart.add_series({
    'name': 'Структура продаж',
    'categories': '=Структура!$A$2:$A$6',
    'values': '=Структура!$B$2:$B$6',
    'data_labels': {'value': True, 'percentage': True, 'category': True}
})

chart.set_title({'name': 'Структура продаж по категориям'})
chart.set_legend({'position': 'right'})

worksheet.insert_chart('D2', chart)
workbook.close()
