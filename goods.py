import xlsxwriter


# Создаём файл
with xlsxwriter.Workbook('товары.xlsx') as workbook:
    worksheet = workbook.add_worksheet('Склад')

    # Заголовки
    headers = ['Наименование', 'Категория', 'Количество', 'Цена', 'Поставщик']
    worksheet.write_row('A1', headers)

    # Данные
    products = [
        ['Ноутбук Dell XPS 15', 'Электроника', 15, 89990, 'ООО "ТехноСнаб"'],
        ['Кофемашина Philips', 'Бытовая техника', 8, 45990, 'ИП Петров'],
        ['Стол офисный', 'Мебель', 23, 12990, 'Мебельная фабрика №1'],
        ['Бумага А4 (500 листов)', 'Канцелярия', 150, 349, 'ООО "Канцтовары"']
    ]

    # Записываем данные построчно
    for row_num, product in enumerate(products, start=1):
        worksheet.write_row(row_num, 0, product)

    # НАСТРАИВАЕМ ШИРИНУ КОЛОНОК
    worksheet.set_column('A:A', 25)   # Наименование (длинное)
    worksheet.set_column('B:B', 18)   # Категория
    worksheet.set_column('C:C', 12)   # Количество
    worksheet.set_column('D:D', 15)   # Цена
    worksheet.set_column('E:E', 22)   # Поставщик

    # Можно добавить ещё одну колонку с примечанием
    worksheet.write('F1', 'Примечание')
    worksheet.set_column('F:F', 20)   # И для неё ширину

print("Файл с настроенной шириной колонок создан!")
