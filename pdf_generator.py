from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import logging
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os

# Настройка логирования
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s: %(message)s')

def generate_pdf(data, filename='selected_personnel.pdf', sheet_title='Без названия'):
    """
    Генерирует PDF-файл с добавлением названия таблицы.
    """
    logging.info("Начало процесса генерации PDF")
    
    # Путь к файлу шрифта, который поддерживает кириллицу
    font_path = 'DejaVuSans.ttf'
    if not os.path.isfile(font_path):
        logging.error(f"Файл шрифта {font_path} не найден.")
        return
    
    # Регистрируем шрифт
    pdfmetrics.registerFont(TTFont('DejaVuSans', font_path))
    
    # Создаем PDF с помощью библиотеки ReportLab
    c = canvas.Canvas(filename, pagesize=letter)
    width, height = letter
    
    # Настройка шрифта, размера и отступов
    title_font_size = 16
    title_top_margin = height - 50  # Верхний отступ в пикселях
    title_left_margin = 50  # Левый отступ в пикселях
    c.setFont("DejaVuSans", title_font_size)
    c.setFillColorRGB(0, 0, 0)  # Черный цвет

    # Добавление названия таблицы
    c.drawString(title_left_margin, title_top_margin, f"Название таблицы: {sheet_title}")
    
    # Смещение для данных
    data_top_margin = title_top_margin - 30  # Смещение ниже заголовка
    line_spacing = 15  # Межстрочный интервал
    c.setFont("DejaVuSans", 12) 
    
    if not data:
        c.drawString(title_left_margin, data_top_margin, "Нет данных для отображения")
    else:
        y_position = data_top_margin
        for person in data:
            line = f"Имя: {person['name']}, Город: {person['city']}, Стоимость: {person['cost']}"
            c.drawString(title_left_margin, y_position, line)
            y_position -= line_spacing
            if y_position < 50:  # Проверка на переполнение страницы
                c.showPage()
                c.setFont("DejaVuSans", 12)
                y_position = height - 50
    
    c.save()
    logging.info(f"PDF сохранен как {filename}")
