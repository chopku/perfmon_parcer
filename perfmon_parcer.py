import pandas as pd
import os
from pathlib import Path

file_path = input("Введите путь к файлу, выходной файл будет сохранен в ту же папку : ")
file_path = Path(file_path)

# Проверяем, существует ли файл
if not os.path.isfile(file_path):
    print(f"Файл {file_path} не найден.")
else:
    print(f'Обрабатываемый файл: {file_path}')

    try:
        # Читаем первую строку отдельно для получения названий счетчиков
        with open(file_path, 'r', encoding='cp1251') as f:
            counters = f.readline().strip().split(',')  # Используем запятую в качестве разделителя

        # Читаем файл, пропуская первую строку, так как в ней метаданные
        df = pd.read_csv(file_path, skiprows=1, encoding='cp1251')  # Читаем CSV без указания разделителя, так как по умолчанию используется запятая

        # Преобразуем все числовые колонки к типу float (если возможно)
        for col in df.columns[1:]:  # Пропускаем первую колонку с датой/временем
            df[col] = pd.to_numeric(df[col], errors='coerce')

        # Вычисляем максимальные, минимальные и средние значения по колонкам
        max_values = df.max(numeric_only=True).round(3)  # Округляем до 3 знаков
        min_values = df.min(numeric_only=True).round(3)  # Округляем до 3 знаков
        mean_values = df.mean(numeric_only=True).round(3)  # Округляем до 3 знаков

        # Список для хранения результатов
        all_results = []

        # Формируем список результатов для записи в Excel
        for name, (col, max_val), (_, min_val), (_, mean_val) in zip(counters[1:], max_values.items(), min_values.items(), mean_values.items()):
            all_results.append([name, max_val, min_val, mean_val, file_path])  # Добавляем имя файла в результаты

        # Преобразуем общий список в DataFrame для записи в Excel
        final_results_df = pd.DataFrame(all_results, columns=['Счетчик', 'Максимальное значение', 'Минимальное значение', 'Среднее значение', 'Файл'])

        # Сохраняем объединенные результаты в Excel файл
        output_filename = file_path.with_suffix(".xlsx")
        final_results_df.to_excel(output_filename, index=False)

        print(f'Все результаты объединены и сохранены в файл {output_filename}')

    except Exception as e:
        print(f"Ошибка при обработке файла {file_path}: {e}")
