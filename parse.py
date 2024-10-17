import pandas as pd
import json

df = pd.read_excel('контрагенты УИД.xlsx')
df = df.iloc[3:].reset_index(drop=True)
df = df.rename(columns={'Unnamed: 0': 'Контрагент УИД', 'Unnamed: 1': 'Контрагент ссылка.Код', 'Unnamed: 2': 'Контрагент ссылка', 'Unnamed: 3': 'АГЕНТ'})
data = {}

# Пройдитесь по DataFrame и заполните словарь
for index, row in df.iterrows():
  name = row['АГЕНТ'].split()[0]
  number = [row['Контрагент ссылка'], row['Контрагент ссылка.Код'], row['Контрагент УИД']]
  if name in data:
    data[name].append(number)
  else:
    data[name] = [number]

json_data = json.dumps(data)
with open('agents.json', 'w') as f:
  f.write(json_data)


file_name = 'номенклатуры УИД (4).xlsx'

import pandas as pd
import numpy as np
import os
import shutil
from zipfile import ZipFile

# Создаем временную папку
tmp_folder = '/tmp/convert_wrong_excel/'
os.makedirs(tmp_folder, exist_ok=True)

# Распаковываем excel как zip в нашу временную папку
with ZipFile(file_name) as excel_container:
    excel_container.extractall(tmp_folder)

# Переименовываем файл с неверным названием
wrong_file_path = os.path.join(tmp_folder, 'xl', 'SharedStrings.xml')
correct_file_path = os.path.join(tmp_folder, 'xl', 'sharedStrings.xml')
os.rename(wrong_file_path, correct_file_path)

# Запаковываем excel обратно в zip и переименовываем в исходный файл
shutil.make_archive('test', 'zip', tmp_folder)
os.rename('test.zip', file_name)

df1 = pd.read_excel('номенклатуры УИД (4).xlsx')
df1 = df1.iloc[3:].reset_index(drop=True)
df1 = df1.drop('Unnamed: 1', axis=1)
df1 = df1.drop('Unnamed: 2', axis=1)
df1 = df1.drop('Unnamed: 4', axis=1)
df1 = df1.rename(columns={'Unnamed: 0': 'Продуктовое направление', 'Unnamed: 3': 'Номенклатура.Ссылка',
                          'Unnamed: 11': 'Номенклатурная группа УИД',
                          'Unnamed: 10': 'Номенклатурная группа', 'Unnamed: 5': 'Номенклатура УИД',
                          'Unnamed: 6': 'Плановая цена продажи в валюте', 'Unnamed: 7': 'Валюта плановой цены продажи',
                          'Unnamed: 8': 'Единица измерения', 'Unnamed: 9': 'Единица измерения УИД',
                          'Unnamed: 12': 'Номенклатура.Ссылка.Артикул', 'Unnamed: 13': 'Номенклатура.Ссылка.Архивный'})
name_dict1 = {}
df1['Номенклатура.Ссылка.Архивный'] = df1['Номенклатура.Ссылка.Архивный'].map({'Да': True, 'Нет': False})
for index, row in df1.iterrows():
  name = row['Номенклатурная группа']
  number = [row['Номенклатура.Ссылка'], row['Продуктовое направление'], row['Номенклатурная группа УИД'], row['Номенклатурная группа'], row['Номенклатура УИД'],
            row['Номенклатура.Ссылка.Артикул'], row['Номенклатура.Ссылка.Архивный'],
            row['Единица измерения УИД'], row['Единица измерения'],
            row['Плановая цена продажи в валюте'], row['Валюта плановой цены продажи']]
  if name in name_dict1:
    name_dict1[name].append(number)
  else:
    name_dict1[name] = [number]
json_data1 = json.dumps(name_dict1)
with open('goods.json', 'w') as f:
  f.write(json_data1)