import xlsxwriter

workbook = xlsxwriter.Workbook('сравнение.xlsx')
worksheet = workbook.add_worksheet('Сравнение')

# Данные
months = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май']
product_a = [120, 135, 140, 155, 170]
product_b = [80, 95, 110, 125, 140]

# Записываем данные
worksheet.write_row('A1', ['Месяц', 'Товар А', 'Товар Б'])
worksheet.write_column('A2', months)
worksheet.write_column('B2', product_a)
worksheet.write_column('C2', product_b)

# Создаём столбчатую диаграмму
chart = workbook.add_chart({'type': 'column'})

# Добавляем первую серию (Товар А)
chart.add_series({
    'name': 'Товар А',
    'categories': '=Сравнение!$A$2:$A$6',
    'values': '=Сравнение!$B$2:$B$6',
    'fill': {'color': '#4472C4'}
})

# Добавляем вторую серию (Товар Б)
chart.add_series({
    'name': 'Товар Б',
    'categories': '=Сравнение!$A$2:$A$6',
    'values': '=Сравнение!$C$2:$C$6',
    'fill': {'color': '#ED7D31'}
})

# Настройки
chart.set_title({'name': 'Сравнение продаж'})
chart.set_x_axis({'name': 'Месяц'})
chart.set_y_axis({'name': 'Тыс. руб.'})

worksheet.insert_chart('E2', chart)
workbook.close()
