import json
import os

# Чтение файла agents.json
with open('agents.json', 'r', encoding='utf-8') as f:
    agents_data = json.load(f)

# Создание содержимого для файла agents.js
js_content = f"const agents = {json.dumps(agents_data, indent=2)};"

# Запись данных в файл agents.js
with open('agents.js', 'w', encoding='utf-8') as f:
    f.write(js_content)

# Удаление файла agents.json
os.remove('agents.json')

print("Файл agents.js создан, файл agents.json удален.")


# Чтение файла agents.json
with open('goods.json', 'r', encoding='utf-8') as f:
    goods_data = json.load(f)

# Создание содержимого для файла agents.js
js_content1 = f"const goods = {json.dumps(goods_data, indent=2)};"

# Запись данных в файл agents.js
with open('goods.js', 'w', encoding='utf-8') as f:
    f.write(js_content1)

# Удаление файла agents.json
os.remove('goods.json')

print("Файл goods.js создан, файл goods.json удален.")
