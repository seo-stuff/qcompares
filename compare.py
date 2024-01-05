import pandas as pd

# Чтение файлов
wm_data = pd.read_excel('wm.xlsx', index_col=0, decimal=',')
topvisor_data = pd.read_excel('topvisor.xlsx', index_col=0)

# Получение общих запросов
common_queries = list(set(topvisor_data.index) & set(wm_data.index))

# Выводим количество общих запросов
print(f"Найдено {len(common_queries)} общих запросов.")

# Формирование DataFrame для записи результата
result_data = pd.DataFrame(index=common_queries)

# Заполнение DataFrame результатами анализа
result_data['Поисковые запросы'] = common_queries
result_data['Topvisor'] = topvisor_data.loc[common_queries].iloc[:, 0]  # Позиции Topvisor всегда в первой колонке
result_data['Webmaster'] = wm_data.loc[common_queries].iloc[:, 1]  # Позиции Webmaster всегда во второй колонке

# Рассчет разницы между позициями и создание колонки "Сравнение"
result_data['Сравнение'] = result_data['Topvisor'] - result_data['Webmaster']

# Рассчет средней частоты и создание колонки "Частота"
demand_columns = [col for col in wm_data.columns if col.endswith('_demand')]
result_data['Частота'] = round(wm_data.loc[common_queries, demand_columns].mean(axis=1))

# Сортировка по убыванию по колонке "Частота"
result_data.sort_values(by='Частота', ascending=False, inplace=True)

# Запись результатов в файл compare.xlsx
result_data.to_excel('compare.xlsx', index=False)

# Ожидание нажатия клавиши Enter
input("Нажмите Enter для завершения программы.")
