import xlsxwriter

# Создаём файл
with xlsxwriter.Workbook('отчёт.xlsx') as workbook:

    # === СОЗДАЁМ ФОРМАТЫ ===

    # Формат для заголовков
    header_format = workbook.add_format({
        'bold': True,
        'font_name': 'Arial',
        'font_size': 12,
        'font_color': 'white',
        'bg_color': '#4472C4',      # синий фон
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'
    })

    # Формат для обычных ячеек с границами
    cell_format = workbook.add_format({
        'border': 1,
        'align': 'left',
        'valign': 'vcenter'
    })

    # Формат для чисел (с разделителями)
    number_format = workbook.add_format({
        'border': 1,
        'num_format': '# ##0.00',
        'align': 'right',
        'valign': 'vcenter'
    })

    # Формат для выделения (жёлтый фон)
    highlight_format = workbook.add_format({
        'border': 1,
        'bg_color': '#FFE699',
        'bold': True,
        'align': 'center'
    })

    # Формат для итоговой строки
    total_format = workbook.add_format({
        'bold': True,
        'border': 1,
        'bg_color': '#D9E1F2',
        'num_format': '# ##0.00',
        'align': 'right'
    })

    # === СОЗДАЁМ ЛИСТ И ЗАПИСЫВАЕМ ДАННЫЕ ===

    worksheet = workbook.add_worksheet('Продажи')

    # Заголовки
    headers = ['Товар', 'Категория', 'Цена', 'Количество', 'Сумма']
    for col_num, header in enumerate(headers):
        worksheet.write(0, col_num, header, header_format)

    # Данные
    data = [
        ['Ноутбук', 'Электроника', 89990, 5],
        ['Мышь', 'Электроника', 1500, 20],
        ['Стол', 'Мебель', 12990, 3],
        ['Стул', 'Мебель', 4990, 8],
        ['Бумага', 'Канцелярия', 349, 50]
    ]

    # Записываем данные с разными форматами
    for row_num, row_data in enumerate(data, start=1):
        # Товар (текст)
        worksheet.write(row_num, 0, row_data[0], cell_format)
        # Категория (текст)
        worksheet.write(row_num, 1, row_data[1], cell_format)
        # Цена (число)
        worksheet.write(row_num, 2, row_data[2], number_format)
        # Количество (число)
        worksheet.write(row_num, 3, row_data[3], number_format)
        # Сумма (формула)
        worksheet.write_formula(row_num, 4, f'=C{row_num+1}*D{row_num+1}',
                                number_format)

        # Выделяем товары с ценой выше 50000
        if row_data[2] > 50000:
            worksheet.write(row_num, 0, row_data[0], highlight_format)
            worksheet.write(row_num, 2, row_data[2], highlight_format)

    # Итоговая строка
    last_row = len(data) + 1
    worksheet.write(last_row, 0, 'ИТОГО:', total_format)
    worksheet.write(last_row, 1, '', cell_format)
    worksheet.write(last_row, 2, '', cell_format)
    worksheet.write_formula(last_row, 3, f'=SUM(D2:D{last_row})', total_format)
    worksheet.write_formula(last_row, 4, f'=SUM(E2:E{last_row})', total_format)

    # === НАСТРАИВАЕМ ШИРИНУ КОЛОНОК ===

    worksheet.set_column('A:A', 20)   # Товар
    worksheet.set_column('B:B', 15)   # Категория
    worksheet.set_column('C:C', 15)   # Цена
    worksheet.set_column('D:D', 12)   # Количество
    worksheet.set_column('E:E', 15)   # Сумма

    # Применяем формат заголовков ко всей первой строке (на всякий случай)
    # worksheet.set_row(0, None, header_format)

print("Красивый отчёт создан! Откройте отчёт.xlsx")
