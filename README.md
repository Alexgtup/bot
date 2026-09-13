# Telegram bot for document workflow

Python Telegram bot that works with contractor data, PowerPoint templates and generated documents.

The bot reads structured data from Google Sheets, lets a user select records through Telegram, fills a `.pptx` template, converts the result to PDF and sends the generated document back in chat. The project also includes template upload/selection flows and image loading from Google Drive links.

## What is implemented

- Telegram commands, inline keyboards and multi-step conversations
- contractor data retrieval from Google Sheets
- selection of records inside the bot
- upload and management of PowerPoint templates
- `.pptx` generation with `python-pptx`
- conversion of generated presentations to PDF
- image loading and processing with Pillow
- logging and basic error handling for external operations

## Stack

`Python` · `python-telegram-bot` · `Google Sheets` · `python-pptx` · `Pillow` · `requests`

## Secure configuration

Secrets are not stored in source code. Copy `.env.example` to your local environment or configure the same variables in your deployment platform:

```text
TELEGRAM_BOT_TOKEN=
GOOGLE_SHEETS_ID=
GOOGLE_SHEETS_RANGE=Sheet1!A2:J
```

Google service-account credentials are read from `credentials.json`. The file is excluded by `.gitignore` and must be provided securely at deploy time.

Never commit a Telegram token, service-account JSON, `.env` file or other production credentials.

## Validation

GitHub Actions runs on every push and pull request:

- `python -m py_compile` for the Python sources
- a repository scan for obvious Telegram tokens, private keys, GitHub tokens and AWS access keys

This is a guard against accidental credential leaks, not a replacement for rotating any credential that has ever been exposed publicly.

## Project structure

- `main.py` - Telegram flows, environment configuration and document-generation logic
- `sheets.py` - Google Sheets integration
- `pdf_generator.py` - PDF-related generation helpers
- `.env.example` - safe configuration template without real values
- `.github/workflows/security-ci.yml` - syntax and secret-leak guard

## Related development work

This repository is one example of my Python/Telegram work. More detailed service pages and cases are published in the portfolio:

- Telegram bot development: https://alexgtup.github.io/telegram-bots/
- Telegram bot for income and expense tracking - Fin Planner case: https://alexgtup.github.io/cases/fin-planner/
- Python development: https://alexgtup.github.io/python-development/
- Existing project repair and improvements: https://alexgtup.github.io/project-repair/
- Portfolio cases: https://alexgtup.github.io/cases/

Main portfolio: https://alexgtup.github.io/
