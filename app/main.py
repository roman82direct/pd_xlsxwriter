import pandas as pd

# Данные
df = pd.DataFrame({
    'Товар': ['Ноутбук', 'Мышь', 'Стол', 'Стул', 'Бумага'],
    'Категория': ['Электроника', 'Электроника',
                  'Мебель', 'Мебель', 'Канцелярия'],
    'Цена': [89990, 1500, 12990, 4990, 349],
    'Количество': [5, 20, 3, 8, 50]
})
df['Сумма'] = df['Цена'] * df['Количество']


def auto_set_column_width(worksheet, df):
    """Автоматически подбирает ширину колонок на основе данных"""
    for i, col in enumerate(df.columns):
        # Длина заголовка
        max_len = len(str(col))

        # Длина данных в колонке
        col_data = df[col].astype(str)
        max_len = max(max_len, col_data.map(len).max())

        # Устанавливаем ширину с запасом
        worksheet.set_column(i, i, min(max_len + 2, 50)) 


# Сохраняем с форматированием
with pd.ExcelWriter('продажи.xlsx', engine='xlsxwriter') as writer:
    # Сначала сохраняем данные (без индексов)
    df.to_excel(writer, sheet_name='Отчёт', index=False)

    # Получаем объекты
    workbook = writer.book
    worksheet = writer.sheets['Отчёт']

    # Формат для заголовков (синий)
    header_format = workbook.add_format({
        'bold': True,
        'bg_color': '#4472C4',
        'font_color': 'white',
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'
    })

    # Формат для текстовых колонок
    text_format = workbook.add_format({
        'border': 1,
        'align': 'left',
        'valign': 'vcenter'
    })

    # Формат для денег
    money_format = workbook.add_format({
        'border': 1,
        'num_format': '# ##0 ₽',
        'align': 'right',
        'valign': 'vcenter'
    })

    # Формат для целых чисел
    int_format = workbook.add_format({
        'border': 1,
        'num_format': '# ##0',
        'align': 'right',
        'valign': 'vcenter'
    })

    # === ГЛАВНОЕ: принудительно перезаписываем заголовки с нужным форматом ===
    for col_num, col_name in enumerate(df.columns):
        worksheet.write(0, col_num, col_name, header_format)

    # Настраиваем ширину и форматы колонок
    worksheet.set_column('A:A', 20, text_format)   # Товар
    worksheet.set_column('B:B', 18, text_format)   # Категория
    worksheet.set_column('C:C', 12, money_format)  # Цена
    worksheet.set_column('D:D', 12, int_format)    # Количество
    worksheet.set_column('E:E', 15, money_format)  # Сумма

print("Файл создан! Откройте продажи.xlsx")
