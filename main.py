import os
import logging
from aiogram import Bot, Dispatcher, types
from aiogram.contrib.middlewares.logging import LoggingMiddleware
from aiogram.types import ParseMode
from aiogram.utils import executor
from google.oauth2 import service_account
from googleapiclient.discovery import build
import difflib
from aiogram.types import InlineKeyboardMarkup, InlineKeyboardButton
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from datetime import datetime, timedelta

scheduler = AsyncIOScheduler()

from aiogram.dispatcher.middlewares import BaseMiddleware
from datetime import datetime, timedelta

API_TOKEN = '5994114826:AAEQmUHq_wHiPQtf5-eyLHjIV4k6fGfo0sE'

logging.basicConfig(level=logging.INFO)

bot = Bot(token=API_TOKEN)
dp = Dispatcher(bot)
dp.middleware.setup(LoggingMiddleware())

# Google Sheets API credentials
SERVICE_ACCOUNT_FILE = 'botpaeser-c99c11fcf53a.json'
SHEET_ID = '18zD9ctaQ74wUq68uTD2MKJpvG1jk6osJkUTzEIUVPH0'
RANGE_NAME = 'A1:Z1000'

credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=['https://www.googleapis.com/auth/spreadsheets.readonly'])

service = build('sheets', 'v4', credentials=credentials)

# Create an inline keyboard with the "Поиск" button
search_keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True)
search_button = types.KeyboardButton("Поиск")
search_keyboard.add(search_button)

def get_search_keyboard():
    markup = InlineKeyboardMarkup()
    markup.add(InlineKeyboardButton("Поиск слова", switch_inline_query_current_chat=""))
    return markup


@dp.message_handler(commands=['start'])
async def start_command(message: types.Message):
    await message.reply("Нажмите кнопку ниже, чтобы начать поиск:", reply_markup=search_keyboard)

@dp.message_handler(lambda message: message.text == 'Поиск')
async def search_button_handler(message: types.Message):
    await message.reply("Напишите слово, которое хотите найти")

@dp.inline_handler(lambda query: query.query.strip() != '')
async def search_inline(query: types.InlineQuery):
    result_text = await check_google_sheet(query.query)
    result = types.InlineQueryResultArticle(
        id='1',
        title="Результаты поиска",
        input_message_content=types.InputTextMessageContent(result_text, parse_mode=ParseMode.HTML),
        description=result_text
    )
    await query.answer([result])

async def check_google_sheet(word):
    try:
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SHEET_ID, range=RANGE_NAME).execute()
        values = result.get('values', [])
    except Exception as e:
        logging.exception("Произошла ошибка при обращении к Google Sheets API")
        return f"Ошибка: не удалось получить данные из таблицы. Пожалуйста, попробуйте позже."

    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SHEET_ID, range=RANGE_NAME).execute()
    values = result.get('values', [])

    search_word = word.strip().lower()
    found_entries = []
    close_match_found = False

    for row_index, row in enumerate(values):
        for col_index, cell in enumerate(row):
            cell_lower = cell.strip().lower()
            if search_word in cell_lower:
                position = row[1] if len(row) > 1 else "не указана"
                
                date_row_index = row_index
                while date_row_index > 0 and not values[date_row_index][0].strip():
                    date_row_index -= 1
                event_date = values[date_row_index][0].strip()

                entry_text = (f'{cell}\n'
                              f'<b>Дата: {event_date}</b> \n'
                              f'<b>Пункт: {position}</b> ')
                found_entries.append(entry_text)

            # Если совпадений не найдено, проверьте, есть ли похожие слова
            elif not found_entries and search_word not in [entry.split()[0].lower() for entry in found_entries]:
                close_matches = difflib.get_close_matches(search_word, [cell_lower], n=1, cutoff=0.6)
                if close_matches and not close_match_found:
                    close_match_found = True
                    search_word = close_matches[0]

    if found_entries:
        response = f"Я нашел {len(found_entries)} запросов по таблице:\n"
        response += "\n\n".join(found_entries)
        return response

    if close_match_found:
        return f"Слово '{word}' не найдено, но есть '{cell}'.\n\n" + await check_google_sheet(cell)

    return f"Слово '{word}' не найдено"

def get_column_letter(column_number):
    letter = ""
    while column_number > 0:
        column_number, remainder = divmod(column_number - 1, 26)
        letter = chr(65 + remainder) + letter
    return letter

@dp.message_handler(lambda message: message.text not in ['/start', 'Поиск'])
async def search_word_handler(message: types.Message):
    word = message.text
    result_text = await check_google_sheet(word)
    await message.reply(result_text, parse_mode=ParseMode.HTML)

if __name__ == '__main__':
    from aiogram import executor
    executor.start_polling(dp, skip_updates=True)
    scheduler.start()