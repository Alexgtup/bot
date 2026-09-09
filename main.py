import logging
import traceback
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    ContextTypes,
    CallbackQueryHandler,
    MessageHandler,
    filters,
    ConversationHandler,
)
from telegram.error import TimedOut
from sheets import read_sheet, get_sheet_title
from pdf_generator import generate_pdf
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt
import os
import subprocess
import requests
from PIL import Image as PIL_Image
from io import BytesIO
import uuid
from pptx.enum.text import PP_ALIGN

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler("bot_log.txt"),
        logging.StreamHandler(),
    ],
)

ORIGINAL_TEMPLATES_DIR = "original_templates"
UPDATED_TEMPLATES_DIR = "updated_templates"
SPREADSHEET_RANGE = os.getenv("GOOGLE_SHEETS_RANGE", "Sheet1!A2:J")

os.makedirs(ORIGINAL_TEMPLATES_DIR, exist_ok=True)
os.makedirs(UPDATED_TEMPLATES_DIR, exist_ok=True)

logging.info("Настройка логирования завершена. Логирование начато.")


def required_env(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise RuntimeError(f"Environment variable {name} is required")
    return value


def get_spreadsheet_id() -> str:
    return required_env("GOOGLE_SHEETS_ID")


def get_template_names(directory=ORIGINAL_TEMPLATES_DIR):
    template_files = os.listdir(directory)
    return [os.path.splitext(file)[0] for file in template_files if file.endswith(".pptx")]


def convert_drive_url(url):
    if "drive.google.com" in url:
        try:
            file_id = url.split("/d/")[1].split("/")[0]
            return f"https://drive.google.com/uc?export=download&id={file_id}"
        except IndexError:
            logging.error("Невозможно извлечь ID файла из Google Drive URL")
            return url
    return url


def download_photo(url):
    try:
        direct_url = convert_drive_url(url)
        response = requests.get(direct_url, timeout=20)
        response.raise_for_status()
        try:
            image = PIL_Image.open(BytesIO(response.content))
            image.load()
            unique_filename = f"/tmp/photo_{uuid.uuid4().hex}.jpg"
            image.save(unique_filename)
            return unique_filename
        except (IOError, PIL_Image.UnidentifiedImageError) as img_err:
            logging.error("Невалидный файл изображения: %s", img_err)
            return None
    except requests.exceptions.RequestException as exc:
        logging.error("Ошибка при загрузке фото: %s", exc)
        return None


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    logging.info("Команда /start вызвана")
    keyboard = [
        [InlineKeyboardButton("📋 Получить", callback_data="get_contractors")],
        [InlineKeyboardButton("➕ Загрузить шаблон", callback_data="upload_template")],
        [InlineKeyboardButton("📑 Показать шаблоны", callback_data="show_templates")],
        [InlineKeyboardButton("❌ Удалить шаблон", callback_data="delete_template")],
    ]
    await update.message.reply_text(
        "Добро пожаловать! Выберите действие:",
        reply_markup=InlineKeyboardMarkup(keyboard),
    )


async def get_contractors(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    try:
        values = read_sheet(get_spreadsheet_id(), SPREADSHEET_RANGE)
        target = update.callback_query.message if update.callback_query else update.message
        if not values:
            await target.reply_text("Нет данных в таблице.")
            return

        table_header = "Имя            | Фамилия | Город   | Стоимость | Часы | Мин. часы | Трансфер | Instagram         | Портфолио          | VK\n"
        table_divider = "---------------|---------|---------|-----------|------|-----------|----------|-------------------|--------------------|----\n"
        table_rows = ""
        for row in values:
            row = list(row) + [""] * max(0, 10 - len(row))
            name = row[1] or "Не указано"
            parts = name.split()
            surname = parts[1] if len(parts) > 1 else "N/A"
            city = row[2] or "N/A"
            cost = row[3] or "N/A"
            hours = row[4] or "N/A"
            min_hours = row[5] or "N/A"
            transfer = row[6] or "N/A"
            instagram = extract_username(row[7]) if row[7] else "N/A"
            portfolio = row[8] or "N/A"
            vk = extract_username(row[9]) if row[9] else "N/A"
            table_rows += (
                f"{name:<15} | {surname:<7} | {city:<7} | {cost:<9} | "
                f"{hours:<4} | {min_hours:<9} | {transfer:<8} | "
                f"{instagram:<17} | {portfolio:<18} | {vk}\n"
            )
        await target.reply_text(
            f"```\n{table_header}{table_divider}{table_rows}```",
            parse_mode="Markdown",
        )
    except Exception as exc:
        logging.error("Ошибка при получении подрядчиков: %s", exc)
        logging.error(traceback.format_exc())
        target = update.callback_query.message if update.callback_query else update.message
        await target.reply_text("Произошла ошибка при обработке запроса.")


async def show_personnel_list(update: Update, context: ContextTypes.DEFAULT_TYPE) -> str:
    personnel_data = context.user_data.get("personnel_data", [])
    selected_personnel = context.user_data.get("selected_personnel", set())
    keyboard = []
    for row in personnel_data:
        button_text = (
            f"Имя: {row['name']}\n"
            f"Фамилия: {row['surname']}\n"
            f"Город: {row['city']}\n"
            f"Стоимость: {row['cost']}\n"
            f"Часы: {row['hours']}\n"
            f"Мин. часы: {row['min_hours']}\n"
            f"Трансфер: {row['transfer']}\n"
            f"Instagram: {row['instagram']}\n"
            f"Портфолио: {row['portfolio']}\n"
            f"VK: {row['vk']}"
        )
        callback_data = f"select_{row['name']}_{row['surname']}"
        if tuple(row.items()) in selected_personnel:
            button_text = f"✅ {button_text}"
        keyboard.append([InlineKeyboardButton(button_text, callback_data=callback_data)])

    keyboard.append([InlineKeyboardButton("Ввести название и дату", callback_data="enter_title_date")])
    reply_markup = InlineKeyboardMarkup(keyboard)
    if update.callback_query:
        await update.callback_query.message.reply_text("Выберите персонал:", reply_markup=reply_markup)
    else:
        await update.message.reply_text("Выберите персонал:", reply_markup=reply_markup)
    return "SELECTING_PERSONNEL"


async def select_personnel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> str:
    query = update.callback_query
    await query.answer()
    personnel_data = context.user_data.get("personnel_data", [])
    selected_personnel = context.user_data.get("selected_personnel", set())
    selected_person = next(
        (
            row
            for row in personnel_data
            if f"select_{row['name']}_{row['surname']}" == query.data
        ),
        None,
    )
    if selected_person:
        selected_tuple = tuple(selected_person.items())
        if selected_tuple in selected_personnel:
            selected_personnel.remove(selected_tuple)
        else:
            selected_personnel.add(selected_tuple)
    context.user_data["selected_personnel"] = selected_personnel
    return await show_personnel_list(update, context)


async def choose_template(update: Update, context: ContextTypes.DEFAULT_TYPE) -> str:
    templates = get_template_names()
    if not templates:
        await update.callback_query.message.reply_text("Нет доступных шаблонов. Пожалуйста, загрузите шаблон.")
        return ConversationHandler.END
    keyboard = [
        [InlineKeyboardButton(template, callback_data=f"tpl_{idx}")]
        for idx, template in enumerate(templates)
    ]
    await update.callback_query.message.reply_text(
        "Выберите шаблон:",
        reply_markup=InlineKeyboardMarkup(keyboard),
    )
    return "CHOOSING_TEMPLATE"


async def select_template(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    await query.answer()
    selected_template_idx = int(query.data.split("_")[1])
    templates = get_template_names()
    selected_template = templates[selected_template_idx]
    await query.edit_message_text(text=f"Вы выбрали шаблон: {selected_template}")

    output_directory = "templates"
    os.makedirs(output_directory, exist_ok=True)
    selected_personnel = [dict(person) for person in context.user_data.get("selected_personnel", set())]
    title = context.user_data.get("title", "")
    date = context.user_data.get("date", "")
    sheet_title = context.user_data.get("sheet_title", "Без названия")

    pptx_template_path = os.path.join(output_directory, f"updated_{selected_template}.pptx")
    pdf_output_path = os.path.join(output_directory, f"updated_{selected_template}.pdf")
    fill_ppt_template(
        selected_personnel,
        selected_template,
        pptx_template_path,
        title=title,
        date=date,
        sheet_title=sheet_title,
    )
    convert_pptx_to_pdf(pptx_template_path, pdf_output_path)
    try:
        with open(pdf_output_path, "rb") as document:
            await query.message.reply_document(document)
    except TimedOut:
        logging.error("Время ожидания истекло при отправке документа. Повторная попытка...")
        with open(pdf_output_path, "rb") as document:
            await query.message.reply_document(document)
    return ConversationHandler.END


async def upload_template(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.callback_query.message.reply_text("Пожалуйста, загрузите файл шаблона (формат .pptx).")


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    document = update.message.document
    if document.mime_type != "application/vnd.openxmlformats-officedocument.presentationml.presentation":
        await update.message.reply_text("Поддерживаются только файлы .pptx")
        return
    file = await document.get_file()
    safe_name = os.path.basename(document.file_name)
    file_path = os.path.join(ORIGINAL_TEMPLATES_DIR, safe_name)
    await file.download_to_drive(file_path)
    await update.message.reply_text(f"Шаблон {safe_name} успешно загружен и сохранён.")


async def show_templates(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    templates = get_template_names()
    if templates:
        await update.callback_query.message.reply_text(f"Доступные шаблоны:\n{'\n'.join(templates)}")
    else:
        await update.callback_query.message.reply_text("Нет доступных шаблонов.")


async def delete_template(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    templates = get_template_names()
    if not templates:
        await update.callback_query.message.reply_text("Нет доступных шаблонов для удаления.")
        return
    keyboard = [
        [InlineKeyboardButton(template, callback_data=f"del_{idx}")]
        for idx, template in enumerate(templates)
    ]
    await update.callback_query.message.reply_text(
        "Выберите шаблон для удаления:",
        reply_markup=InlineKeyboardMarkup(keyboard),
    )


async def confirm_delete_template(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    await query.answer()
    templates = get_template_names()
    idx = int(query.data.split("_")[1])
    if idx < 0 or idx >= len(templates):
        await query.edit_message_text("Некорректный шаблон.")
        return
    selected_template = templates[idx]
    file_path = os.path.join(ORIGINAL_TEMPLATES_DIR, f"{selected_template}.pptx")
    try:
        os.remove(file_path)
        await query.edit_message_text(text=f"Шаблон {selected_template} успешно удалён.")
    except OSError as exc:
        logging.error("Ошибка при удалении шаблона: %s", exc)
        await query.edit_message_text(text="Не удалось удалить шаблон.")


async def button(update: Update, context: ContextTypes.DEFAULT_TYPE) -> str:
    query = update.callback_query
    await query.answer()
    if query.data == "get_contractors":
        return await get_personnel_data(update, context)
    if query.data == "enter_title_date":
        return await ask_for_title(update, context)
    if query.data == "upload_template":
        await upload_template(update, context)
        return ConversationHandler.END
    if query.data == "show_templates":
        await show_templates(update, context)
        return ConversationHandler.END
    if query.data == "delete_template":
        await delete_template(update, context)
        return ConversationHandler.END
    if query.data.startswith("del_"):
        await confirm_delete_template(update, context)
        return ConversationHandler.END
    if query.data.startswith("select_"):
        return await select_personnel(update, context)
    if query.data == "choose_template":
        return await choose_template(update, context)
    if query.data.startswith("tpl_"):
        return await select_template(update, context)
    return ConversationHandler.END


async def ask_for_title(update: Update, context: ContextTypes.DEFAULT_TYPE) -> str:
    if update.callback_query:
        await update.callback_query.message.reply_text("Пожалуйста, введите название:")
    else:
        await update.message.reply_text("Пожалуйста, введите название:")
    return "WAITING_FOR_TITLE"


async def receive_title(update: Update, context: ContextTypes.DEFAULT_TYPE) -> str:
    context.user_data["title"] = update.message.text
    await update.message.reply_text("Теперь введите дату:")
    return "WAITING_FOR_DATE"


async def receive_date(update: Update, context: ContextTypes.DEFAULT_TYPE) -> str:
    context.user_data["date"] = update.message.text
    keyboard = [[InlineKeyboardButton("Выбрать шаблон", callback_data="choose_template")]]
    await update.message.reply_text(
        "Выберите действие:",
        reply_markup=InlineKeyboardMarkup(keyboard),
    )
    return "CHOOSING_TEMPLATE"


async def get_personnel_data(update: Update, context: ContextTypes.DEFAULT_TYPE) -> str:
    data = read_sheet(get_spreadsheet_id(), SPREADSHEET_RANGE)
    if not data:
        await update.callback_query.message.reply_text("Не удалось получить данные из таблицы.")
        return ConversationHandler.END

    context.user_data["sheet_title"] = get_sheet_title(get_spreadsheet_id())
    personnel_data = []
    for row in data:
        row = list(row) + ["N/A"] * max(0, 10 - len(row))
        full_name = row[1] if row[1] and row[1] != "N/A" else "Не указано"
        parts = full_name.split()
        personnel_data.append(
            {
                "name": full_name,
                "surname": parts[1] if len(parts) > 1 else "N/A",
                "city": row[2],
                "cost": row[3],
                "hours": row[4],
                "min_hours": row[5],
                "transfer": row[6],
                "instagram": row[7],
                "portfolio": row[8],
                "vk": row[9],
                "photo": row[0],
            }
        )
    context.user_data["personnel_data"] = personnel_data
    return await show_personnel_list(update, context)


def extract_username(url):
    if not url:
        return "N/A"
    if "instagram.com" in url or "vk.com" in url:
        parts = url.split("/")
        if len(parts) > 3:
            return parts[3].split("?")[0]
    return url


def fill_ppt_template(selected_people, template_name, output_path="output.pptx", title="", date="", sheet_title=""):
    template_path = os.path.join(ORIGINAL_TEMPLATES_DIR, f"{template_name}.pptx")
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Шаблон {template_path} не найден.")

    presentation = Presentation(template_path)
    current_slide = presentation.slides[0] if presentation.slides else presentation.slides.add_slide(presentation.slide_layouts[0])

    title_shape = current_slide.shapes.title
    if title_shape:
        title_shape.text = title
        for paragraph in title_shape.text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.color.rgb = RGBColor(255, 255, 255)

    date_shape = current_slide.shapes.add_textbox(Inches(3.15), Inches(12.15), Inches(2), Inches(0.5))
    date_run = date_shape.text_frame.paragraphs[0].add_run()
    date_run.text = date
    date_run.font.name = "Helvetica"
    date_run.font.size = Pt(20)
    date_run.font.color.rgb = RGBColor(255, 255, 255)

    if sheet_title:
        sheet_box = current_slide.shapes.add_textbox(Inches(0.5), Inches(11.75), Inches(6), Inches(0.5))
        sheet_run = sheet_box.text_frame.paragraphs[0].add_run()
        sheet_run.text = sheet_title
        sheet_run.font.name = "Helvetica"
        sheet_run.font.size = Pt(32)
        sheet_run.font.color.rgb = RGBColor(0, 0, 0)

    slide_height = presentation.slide_height
    current_top = Inches(1400 / 96)
    left_margin_image = Inches(0.1)
    image_width = Inches(2.6)
    image_height = Inches(3.6)
    textbox_width = Inches(2.9)
    textbox_height = Inches(2)
    spacing_between_users = Inches(1)

    for person in selected_people:
        if current_top + max(image_height, textbox_height) > slide_height - Inches(0.5):
            current_slide = presentation.slides.add_slide(presentation.slide_layouts[5])
            current_top = Inches(1)

        photo_url = person.get("photo", "N/A")
        if photo_url != "N/A" and str(photo_url).startswith("http"):
            photo_path = download_photo(photo_url)
            if photo_path:
                current_slide.shapes.add_picture(
                    photo_path,
                    left_margin_image,
                    current_top,
                    width=image_width,
                    height=image_height,
                )
                try:
                    os.remove(photo_path)
                except OSError:
                    pass

        textbox_left = left_margin_image + image_width + Inches(0.2)
        textbox_top = current_top + Inches(22 / 96)
        textbox = current_slide.shapes.add_textbox(textbox_left, textbox_top, textbox_width, textbox_height)
        text_frame = textbox.text_frame
        text_frame.clear()
        text_frame.margin_left = Inches(0.52)

        lines = [
            (person.get("name"), 18, RGBColor(0, 0, 0)),
            (f"/ {person.get('city')}" if person.get("city") not in (None, "N/A") else None, 14, RGBColor(128, 128, 128)),
            (f"{person.get('cost')} / час" if person.get("cost") not in (None, "N/A") else None, 14, RGBColor(0, 0, 0)),
            (f"Минимально от {person.get('min_hours')} часов" if person.get("min_hours") not in (None, "N/A") else None, 14, RGBColor(128, 128, 128)),
            (f"+ {person.get('transfer')} / трансфер" if person.get("transfer") not in (None, "N/A") else None, 14, RGBColor(0, 0, 0)),
            (extract_username(person.get("instagram")) if person.get("instagram") not in (None, "N/A") else None, 20, RGBColor(0, 0, 0)),
            (person.get("portfolio") if person.get("portfolio") not in (None, "N/A") else None, 20, RGBColor(0, 0, 0)),
            (extract_username(person.get("vk")) if person.get("vk") not in (None, "N/A") else None, 20, RGBColor(0, 0, 0)),
        ]
        for text, size, color in lines:
            if not text:
                continue
            paragraph = text_frame.add_paragraph()
            run = paragraph.add_run()
            run.text = str(text).upper() if size == 18 else str(text)
            run.font.name = "Helvetica"
            run.font.size = Pt(size)
            run.font.color.rgb = color

        current_top += max(image_height, textbox_height) + spacing_between_users

    presentation.save(output_path)


def convert_pptx_to_pdf(input_pptx, output_pdf):
    try:
        subprocess.run(
            ["unoconv", "-f", "pdf", "-o", output_pdf, input_pptx],
            check=True,
        )
    except subprocess.CalledProcessError as exc:
        logging.error("Ошибка при конвертации PPTX в PDF: %s", exc)
        raise


def main():
    token = required_env("TELEGRAM_BOT_TOKEN")
    application = ApplicationBuilder().token(token).build()

    conv_handler = ConversationHandler(
        entry_points=[CallbackQueryHandler(button)],
        states={
            "WAITING_FOR_TITLE": [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_title)],
            "WAITING_FOR_DATE": [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_date)],
            "CHOOSING_TEMPLATE": [
                CallbackQueryHandler(select_template, pattern=r"^tpl_"),
                CallbackQueryHandler(choose_template, pattern=r"^choose_template$"),
            ],
            "SELECTING_PERSONNEL": [
                CallbackQueryHandler(select_personnel, pattern=r"^select_"),
                CallbackQueryHandler(button),
            ],
        },
        fallbacks=[],
    )

    application.add_handler(CommandHandler("start", start))
    application.add_handler(conv_handler)
    application.add_handler(MessageHandler(filters.Document.ALL, handle_document))

    logging.info("Бот запущен и готов к работе")
    application.run_polling()


if __name__ == "__main__":
    main()
