import pandas as pd

# Шаг 1: Чтение данных из файла Excel
df = pd.read_excel('имя_файла.xlsx')

# Шаг 2: Обработка данных (необязательно)
# ...

# Шаг 3: Запись данных в файл Excel
df.to_excel('новый_файл.xlsx', index=False)  # Установите index=False, чтобы не сохранять индексы DataFrame
В этом примере сначала мы читаем данные из файла Excel в DataFrame с помощью функции read_excel(), затем обрабатываем данные (если это необходимо), и наконец записываем измененные данные в новый файл Excel с помощью метода to_excel().

Библиотека openpyxl:
python
Copy code
from openpyxl import load_workbook

# Шаг 1: Загрузка книги Excel
workbook = load_workbook('имя_файла.xlsx')

# Шаг 2: Получение активного листа
sheet = workbook.active

# Шаг 3: Чтение/изменение данных (необязательно)
# ...

# Шаг 4: Сохранение книги
workbook.save('новый_файл.xlsx')
Здесь мы открываем книгу Excel с помощью load_workbook(), получаем активный лист и вносим необходимые изменения в данные (если это необходимо), а затем сохраняем книгу с помощью метода save().

Библиотека xlrd/xlwt:
python
Copy code
import xlrd
import xlwt

# Шаг 1: Чтение данных из файла Excel
workbook = xlrd.open_workbook('имя_файла.xlsx')
sheet = workbook.sheet_by_index(0)

# Шаг 2: Получение данных (необязательно)
# ...

# Шаг 3: Создание новой книги Excel и листа
new_workbook = xlwt.Workbook()
new_sheet = new_workbook.add_sheet('Новый лист')

# Шаг 4: Запись данных в новую книгу Excel (необязательно)
# ...

# Шаг 5: Сохранение новой книги Excel
new_workbook.save('новый_файл.xlsx')
Здесь мы сначала открываем файл Excel с помощью xlrd.open_workbook(), читаем данные (если это необходимо), создаем новую книгу Excel с помощью xlwt.Workbook() и добавляем в нее новый лист. Затем мы записываем данные в новую книгу (если это необходимо) и сохраняем ее с помощью метода save().
