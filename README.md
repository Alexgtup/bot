# Telegram bot for document workflow

Python Telegram bot that works with contractor data, PowerPoint templates and generated documents.

The bot can read structured data from Google Sheets, let a user select records through Telegram, fill a `.pptx` template, convert the result to PDF and send the generated document back in chat. The project also includes template upload/selection flows and image loading from Google Drive links.

## What is implemented

- Telegram commands, inline keyboards and multi-step conversations;
- contractor data retrieval from Google Sheets;
- selection of records inside the bot;
- upload and management of PowerPoint templates;
- `.pptx` generation with `python-pptx`;
- conversion of generated presentations to PDF;
- image loading and processing with Pillow;
- logging and basic error handling for external operations.

## Stack

`Python` · `python-telegram-bot` · `Google Sheets` · `python-pptx` · `Pillow` · `requests`

## Project structure

- `main.py` — Telegram flows and document-generation logic;
- `sheets.py` — Google Sheets integration;
- `pdf_generator.py` — PDF-related generation helpers.

## More Telegram work

I build and improve Telegram bots, integrations and automation workflows.

- Telegram bot development: https://alexgtup.github.io/telegram-bots/
- Existing project repair and improvements: https://alexgtup.github.io/project-repair/
- Portfolio and cases: https://alexgtup.github.io/cases/

Main portfolio: https://alexgtup.github.io/
