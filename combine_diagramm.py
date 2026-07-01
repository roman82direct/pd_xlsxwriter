import xlsxwriter

workbook = xlsxwriter.Workbook('комби.xlsx')
worksheet = workbook.add_worksheet('Комбинированная')

# Данные
months = ['Янв', 'Фев', 'Мар', 'Апр', 'Май']
sales = [120, 135, 150, 165, 180]
target = [130, 130, 150, 150, 170]

worksheet.write_row('A1', ['Месяц', 'Продажи', 'План'])
worksheet.write_column('A2', months)
worksheet.write_column('B2', sales)
worksheet.write_column('C2', target)

# Создаём комбинированную диаграмму
chart = workbook.add_chart({'type': 'column'})

# Добавляем продажи (столбцы)
chart.add_series({
    'name': 'Продажи',
    'categories': '=Комбинированная!$A$2:$A$6',
    'values': '=Комбинированная!$B$2:$B$6',
    'fill': {'color': '#4472C4'}
})

# Добавляем план (линия)
chart.add_series({
    'name': 'План',
    'categories': '=Комбинированная!$A$2:$A$6',
    'values': '=Комбинированная!$C$2:$C$6',
    'type': 'line',
    'line': {'color': '#ED7D31', 'width': 2, 'dash_type': 'dash'},
    'marker': {'type': 'circle', 'size': 5}
})

chart.set_title({'name': 'Продажи vs План'})
chart.set_x_axis({'name': 'Месяц'})
chart.set_y_axis({'name': 'Тыс. руб.'})

worksheet.insert_chart('E2', chart)
workbook.close()
