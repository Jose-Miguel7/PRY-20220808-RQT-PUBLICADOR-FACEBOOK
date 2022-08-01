import json


category = 2
with open('./data/category.json', encoding="utf8") as r:
    data = json.loads(r.read())
    data_category = data[int(category)]
    category_label = data_category['category']

print(category_label)