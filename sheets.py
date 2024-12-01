import os.path
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import logging

# Определяем область доступа
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
SERVICE_ACCOUNT_FILE = 'credentials.json'  # Учетные данные

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def get_service():
    # Создаем объект учетных данных
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    # Создаем объект службы для работы с Google Sheets API
    logging.info('Создание службы Google Sheets API')
    service = build('sheets', 'v4', credentials=creds)
    return service

def read_sheet(spreadsheet_id, range_name):
    # Получаем объект службы
    service = get_service()
    # Получаем данные из таблицы
    sheet = service.spreadsheets()
    try:
        logging.info(f'Извлечение данных из таблицы {spreadsheet_id}, диапазон {range_name}')
        result = sheet.values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
        values = result.get('values', [])
        logging.info('Данные успешно извлечены')
        return values
    except Exception as e:
        logging.error(f'Ошибка при извлечении данных: {e}')
        return []

def get_sheet_title(spreadsheet_id):
    """
    Получает название таблицы Google Sheets.
    """
    try:
        service = get_service()
        sheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        title = sheet.get('properties', {}).get('title', 'Без названия')
        logging.info(f"Получено название таблицы: {title}")
        return title
    except Exception as e:
        logging.error(f"Ошибка при получении названия таблицы: {e}")
        traceback.print_exc()
        return 'Без названия'
