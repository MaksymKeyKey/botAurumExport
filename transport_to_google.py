import pymongo
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime

# Настройки для доступа к Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('./high-empire-431420-a2-77009419ab03.json', scope)
client = gspread.authorize(creds)

# Открытие Google Sheets
spreadsheet = client.open("botTest")
sheet = spreadsheet.sheet1  # Можно выбрать другой лист, если нужно

# Подключение к MongoDB
mongo_client = pymongo.MongoClient('mongodb+srv://seksikoleg5:se4HivNRYKdydnzc@cluster0.pdc2rrh.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0', 
                                   tls=True, tlsAllowInvalidCertificates=True)
db = mongo_client.botUsers
collection = db.users

# Получение данных из MongoDB
data = list(collection.find())

# Запись заголовков
header = [
    'Дата прибытия', 'Дата последнего взаимодействия', 'Имя', 'Фамилия', 'Имя пользователя',
    'Номер телефона', 'Направление', 'Формат обучения', 'Оплата за модуль №1',
    'Оплата за модуль №2', 'Оплата за модуль №3', 'Оплата за весь курс', 'Комментарий'
]
sheet.clear()  # Очистка всех данных в листе
sheet.insert_row(header, 1)

# Функция для форматирования даты
def format_date(date):
    return date.strftime('%Y-%m-%d %H:%M:%S') if isinstance(date, datetime) else date

# Функция для форматирования поля first_question
def format_first_question(question):
    if question == 'want_to_participate_in_events':
        return 'Хочу участвовать в событиях'
    elif question == 'want_to_be_rehabilitologist':
        return 'Хочу быть реабилитологом'
    elif question == 'want_to_upgrade_qualification':
        return 'Хочу повысить квалификацию'
    return ''

# Преобразование существующих данных в словарь для быстрого поиска
existing_data = {record['Имя пользователя']: record for record in sheet.get_all_records()}

# Обновление данных в Google Sheets
for user in data:
    row = [
        format_date(user.get('arrival_date', 'N/A')),
        format_date(user.get('last_interaction_format', 'N/A')),
        user.get('first_name', 'N/A'),
        user.get('last_name', 'N/A'),
        user.get('username', 'N/A'),
        user.get('contact', 'N/A'),
        format_first_question(user.get('first_question', '')),
        user.get('second_question', 'N/A'),
        'Оплачено' if user.get('payment_status_module_1') == 'Paid' else 'Не оплачено',
        'Оплачено' if user.get('payment_status_module_2') == 'Paid' else 'Не оплачено',
        'Оплачено' if user.get('payment_status_module_3') == 'Paid' else 'Не оплачено',
        'Оплачено' if user.get('payment_status_all_course') == 'Paid' else 'Не оплачено',
        user.get('comments', '')
    ]

    username = user.get('username', 'N/A')

    if username in existing_data:
        # Обновление существующей строки
        row_index = list(existing_data.keys()).index(username) + 2  # Индексы в Google Sheets начинаются с 1, +1 для заголовка
        sheet.update(f'A{row_index}:M{row_index}', [row])
    else:
        # Добавление новой строки
        sheet.append_row(row)

print("Data transferred successfully!")
