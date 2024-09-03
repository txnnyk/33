import os
import re
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import BarChart, Reference
from telethon import TelegramClient
from contextlib import contextmanager

# Ваші api_id, api_hash та номер телефону
api_id = '24105027'
api_hash = 'e0f8668442e6d4ca79d6d5a9ee70e005'
phone_number = "380636182482"

# Функція для конвертації чисел, записаних словами, у числовий формат
def word_to_number(text):
    # Словник для конвертації слів у числа
    number_words = {
        'нуль': 0, 'один': 1, 'два': 2, 'три': 3, 'чотири': 4, 'п’ять': 5, 'пять': 5, 'шість': 6, 'сім': 7, 'вісім': 8, 'дев’ять': 9, 'девять': 9,
        'десять': 10, 'одинадцять': 11, 'дванадцять': 12, 'тринадцять': 13, 'чотирнадцять': 14, 'п’ятнадцять': 15, 'пятнадцять': 15,
        'шістнадцять': 16, 'сімнадцять': 17, 'вісімнадцять': 18, 'дев’ятнадцять': 19, 'девятнадцять': 19, 'двадцять': 20, 'тридцять': 30,
        'сорок': 40, 'п’ятдесят': 50, 'пятдесят': 50, 'шістдесят': 60, 'сімдесят': 70, 'вісімдесят': 80, 'дев’яносто': 90, 'девяносто': 90,
        'сто': 100, 'двісті': 200
    }

    # Розбиваємо текст на слова
    words = text.split()
    total = 0

    for word in words:
        if word in number_words:
            total += number_words[word]
    return total

# Ініціалізація клієнта Telegram
client = TelegramClient('session_name', api_id, api_hash)

@contextmanager
def save_excel_safe(filename):
    try:
        yield
    except PermissionError as e:
        print(f"Permission denied: {e}")
    except Exception as e:
        print(f"Failed to save the file: {e}")
    else:
        print(f"File saved successfully as {filename}")

async def main():
    await client.start(phone_number)

    # Вказуємо username або посилання на канал
    channel_username = 'ComAFUA'

    # Отримуємо канал за username
    channel = await client.get_entity(channel_username)

    # Підготовка даних для таблиці
    data = {
        'Дата': [],
        'Збиті БПЛА/Дрони/Ударні/Шахеди': [],
        'Збиті ракети': [],
        'Загальна кількість цілей': [],
        'Відсоток збитих цілей': []
    }

    # Фільтруємо повідомлення зі словом "збито" і витягуємо дані
    async for message in client.iter_messages(channel.id):
        if message.text and 'збито' in message.text.lower():  # Перевіряємо, чи є слово "збито" у нижньому регістрі
            date = message.date.strftime('%Y-%m-%d')

            # Регулярні вирази для пошуку кількості збитих "БПЛА", "Дронів", "Ударних БПЛА" та "Shahed", незалежно від регістру
            bpla_pattern = re.search(r'(\d+|[а-яА-Я]+)\s*(БПЛА|дрон\w*|ударн\w* БПЛА|Shahed\w*)', message.text, re.IGNORECASE)
            raketa_pattern = re.search(r'(\d+|[а-яА-Я]+)\s*(ракета|ракет\w*|X59\w*|X101|X69\w*)', message.text, re.IGNORECASE)
            targets_pattern = re.search(r'загальна\s*кількість\s*цілей\s*:\s*(\д+)', message.text, re.IGNORECASE)

            # Перевіряємо чи є результат у числовому або словесному форматі і конвертуємо його
            bpla_count = int(bpla_pattern.group(1)) if bpla_pattern and bpla_pattern.group(1).isdigit() else word_to_number(bpla_pattern.group(1)) if bpla_pattern else 0
            raketa_count = int(raketa_pattern.group(1)) if raketa_pattern and raketa_pattern.group(1).isdigit() else word_to_number(raketa_pattern.group(1)) if raketa_pattern else 0
            total_targets = int(targets_pattern.group(1)) if targets_pattern else 0

            # Обчислюємо відсоток збитих цілей
            total_destroyed = bpla_count + raketa_count
            if total_targets > 0:
                percentage_destroyed = (total_destroyed / total_targets) * 100
            else:
                percentage_destroyed = 0

            # Заповнення даних
            data['Дата'].append(date)
            data['Збиті БПЛА/Дрони/Ударні/Шахеди'].append(bpla_count)
            data['Збиті ракети'].append(raketa_count)
            data['Загальна кількість цілей'].append(total_targets)
            data['Відсоток збитих цілей'].append(percentage_destroyed)

    # Створюємо DataFrame
    df = pd.DataFrame(data)

    # Додаємо стовпець для місяця
    df['Місяць'] = pd.to_datetime(df['Дата']).dt.to_period('M').astype(str)  # Конвертація у строковий формат

    # Групуємо дані за місяцями і підраховуємо загальну кількість
    monthly_totals = df.groupby('Місяць')[['Збиті БПЛА/Дрони/Ударні/Шахеди', 'Збиті ракети', 'Загальна кількість цілей']].sum()
    monthly_totals['Відсоток збитих цілей'] = (monthly_totals['Збиті БПЛА/Дрони/Ударні/Шахеди'] + monthly_totals['Збиті ракети']) / monthly_totals['Загальна кількість цілей'] * 100

    # Підраховуємо середнє значення за місяць
    monthly_means = df.groupby('Місяць')[['Збиті БПЛА/Дрони/Ударні/Шахеди', 'Збиті ракети', 'Відсоток збитих цілей']].mean()

    # Створюємо новий Excel файл за допомогою openpyxl
    wb = Workbook()

    # Додавання листа з даними
    ws_data = wb.active
    ws_data.title = 'Дані'
    for r in dataframe_to_rows(df, index=False, header=True):
        ws_data.append(r)

    # Додавання листа із загальною кількістю за місяць
    ws_totals = wb.create_sheet(title='Загальна кількість за місяць')
    for r in dataframe_to_rows(monthly_totals.reset_index(), index=False, header=True):  # Перетворюємо індекс в стовпець
        ws_totals.append(r)

    # Додавання листа із середнім значенням за місяць
    ws_means = wb.create_sheet(title='Середнє значення за місяць')
    for r in dataframe_to_rows(monthly_means.reset_index(), index=False, header=True):  # Перетворюємо індекс в стовпець
        ws_means.append(r)

    # Діаграма загальної кількості за місяць
    chart_totals = BarChart()
    chart_totals.title = 'Загальна кількість збитих БПЛА/Дронів/Ударних/Шахедів та ракет за місяць'
    chart_totals.x_axis.title = 'Місяць'
    chart_totals.y_axis.title = 'Кількість'
    data = Reference(ws_totals, min_col=2, min_row=1, max_col=4, max_row=len(monthly_totals) + 1)
    categories = Reference(ws_totals, min_col=1, min_row=2, max_row=len(monthly_totals) + 1)
    chart_totals.add_data(data, titles_from_data=True)
    chart_totals.set_categories(categories)
    ws_totals.add_chart(chart_totals, 'E5')

    # Діаграма середнього значення за місяць
    chart_means = BarChart()
    chart_means.title = 'Середнє значення збитих БПЛА/Дронів/Ударних/Шахедів та ракет за місяць'
    chart_means.x_axis.title = 'Місяць'
    chart_means.y_axis.title = 'Середнє значення'
    data = Reference(ws_means, min_col=2, min_row=1, max_col=4, max_row=len(monthly_means) + 1)
    categories = Reference(ws_means, min_col=1, min_row=2, max_row=len(monthly_means) + 1)
    chart_means.add_data(data, titles_from_data=True)
    chart_means.set_categories(categories)
    ws_means.add_chart(chart_means, 'E5')

    # Збереження Excel файлу з обробкою виключень
    with save_excel_safe('zbyti_bpla_raketi_with_charts.xlsx'):
        wb.save('zbyti_bpla_raketi_with_charts.xlsx')

    await client.disconnect()

# Запускаємо асинхронну подію
with client:
    client.loop.run_until_complete(main())
