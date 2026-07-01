import xlsxwriter


with xlsxwriter.Workbook('formats.xlsx') as workbook:
    # Создаём формат
    # наглядно
    format1 = workbook.add_format({
        'bold': True,
        'font_color': 'blue',
        'bg_color': '#FFFFCC'
    })

    table_border = workbook.add_format({
        'border': 1,
        'border_color': 'gray'
    })

    money_format = workbook.add_format({
        'num_format': '# ##0.00 ₽',
        'align': 'right'
    })

    percent_format = workbook.add_format({
        'num_format': '0.00%',
        'align': 'right'
    })

    date_format = workbook.add_format({
        'num_format': 'dd.mm.yyyy',
        'align': 'center'
    })

    # компактно:
    format2 = workbook.add_format(bold=True, font_color='blue',
                                  bg_color='#FFFFCC')

    bold = workbook.add_format({'bold': True})

    ws = workbook.add_worksheet('formats')
    # Применяем при записи
    ws.write('A1', 'Заголовок', bold)


# # Список основных цветов по именам:

# colors = [
#     'black', 'blue', 'brown', 'cyan', 'gray', 'green', 'lime',
#     'magenta', 'navy', 'orange', 'pink', 'purple', 'red',
#     'white', 'yellow', 'light_gray', 'dark_gray', 'dark_blue',
#     'dark_cyan', 'dark_green', 'dark_magenta', 'dark_red', 'dark_yellow'
# ]

# # Советы по форматированию
# # Создавайте форматы один раз в начале и переиспользуйте.
# # Не создавайте новый формат для каждой ячейки.
# # Именуйте форматы понятно:

# # Плохо
# f1 = workbook.add_format({'bold': True})
# f2 = workbook.add_format({'bold': True, 'bg_color': 'yellow'})

# # Хорошо
# header_bold = workbook.add_format({'bold': True})
# warning_yellow = workbook.add_format({'bold': True, 'bg_color': 'yellow'})

# # Для чисел всегда указывайте num_format.
# # Даже если число целое, лучше задать формат:

# int_format = workbook.add_format({'num_format': '0'})
# money_format = workbook.add_format({'num_format': '# ##0.00'})
# percent_format = workbook.add_format({'num_format': '0.00%'})

# # Комбинируйте форматы с шириной колонок:
# # Можно сразу задать и ширину, и формат
# worksheet.set_column('B:B', 15, number_format)
