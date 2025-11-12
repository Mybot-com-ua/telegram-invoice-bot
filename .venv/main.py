from aiogram import Bot, Dispatcher, F, types
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.context import FSMContext
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import Message, ReplyKeyboardMarkup, KeyboardButton, ReplyKeyboardRemove
from aiogram.enums import ParseMode
from aiogram import Router
from aiogram.types import FSInputFile
from aiogram.utils.keyboard import ReplyKeyboardBuilder
from aiogram.utils.markdown import hbold
from aiogram.client.session.aiohttp import AiohttpSession
from aiogram import Dispatcher
from openpyxl import load_workbook
from num2words import num2words
from aiogram import BaseMiddleware
import asyncio
import logging
import os
from openpyxl.cell.cell import MergedCell
import subprocess
from dotenv import load_dotenv

load_dotenv()
API_TOKEN = os.getenv('API_TOKEN')

bot = Bot(token=API_TOKEN)
storage = MemoryStorage()
dp = Dispatcher(storage=storage)

os.makedirs("docs", exist_ok=True)
TEMPLATE_PATH = "Рахунок шаблон 8.xlsx"


class InvoiceStates(StatesGroup):
    фирма = State()
    номер_счета = State()
    дата = State()
    добавление_услуги = State()
    услуга = State()
    количество = State()
    цена = State()

start_keyboard = ReplyKeyboardMarkup(resize_keyboard=True, keyboard=[
    [KeyboardButton(text="Создать счёт")]
])

@dp.message(F.text.in_(["/start", "Создать счёт"]))
async def cmd_start(message: Message, state: FSMContext):
    await state.clear()
    await message.answer("Введите название фирмы:")
    await state.set_state(InvoiceStates.фирма)

@dp.message(InvoiceStates.фирма)
async def process_firma(message: Message, state: FSMContext):
    await state.update_data(фирма=message.text, услуги=[])
    await message.answer("Введите номер счёта:")
    await state.set_state(InvoiceStates.номер_счета)

@dp.message(InvoiceStates.номер_счета)
async def process_number(message: Message, state: FSMContext):
    from datetime import datetime
    today = datetime.today().strftime('%d.%m.%Y')
    await state.update_data(номер_счета=message.text)
    keyboard = ReplyKeyboardMarkup(resize_keyboard=True, keyboard=[[KeyboardButton(text=today)]])
    await message.answer(f"Введите дату (например, 16.04.2025): Или выберите сегодняшнюю:", reply_markup=keyboard)
    await state.set_state(InvoiceStates.дата)

@dp.message(InvoiceStates.дата)
async def process_date(message: Message, state: FSMContext):
    await state.update_data(дата=message.text)
    await message.answer("Введите название услуги:")
    await state.set_state(InvoiceStates.услуга)

@dp.message(InvoiceStates.услуга)
async def process_service(message: Message, state: FSMContext):
    await state.update_data(текущая_услуга={'название': message.text})
    await message.answer("Введите количество:")
    await state.set_state(InvoiceStates.количество)

@dp.message(InvoiceStates.количество)
async def process_quantity(message: Message, state: FSMContext):
    try:
        quantity = round(float(message.text.replace(',', '.')), 2)
    except ValueError:
        await message.answer("❌ **Ошибка!** Введите, пожалуйста, количество числом (например, 5 или 1.5).")
        return

    data = await state.get_data()
    data['текущая_услуга']['количество'] = quantity
    await state.update_data(текущая_услуга=data['текущая_услуга'])
    await message.answer("Введите цену за единицу:")
    await state.set_state(InvoiceStates.цена)

@dp.message(InvoiceStates.цена)
async def process_price(message: Message, state: FSMContext):
    try:
        price = round(float(message.text.replace(',', '.')), 2)
    except ValueError:
        await message.answer("❌ **Ошибка!** Введите, пожалуйста, цену числом (например, 1000.50).")
        return

    data = await state.get_data()
    услуга = data['текущая_услуга']
    услуга['цена'] = price
    услуга['сумма'] = услуга['количество'] * услуга['цена']
    data['услуги'].append(услуга)
    await state.update_data(услуги=data['услуги'])

    # ... остальной код функции (добавление/завершение) ...
    if len(data['услуги']) >= 10:
        await finalize_invoice(state, message)
        return

    keyboard = ReplyKeyboardMarkup(resize_keyboard=True, keyboard=[
        [KeyboardButton(text="Добавить ещё услугу")],
        [KeyboardButton(text="Завершить счёт")]
    ])
    await message.answer("Услуга добавлена. Что дальше?", reply_markup=keyboard)
    await state.set_state(InvoiceStates.добавление_услуги)

@dp.message(InvoiceStates.добавление_услуги)
async def process_add_more(message: Message, state: FSMContext):
    if message.text.lower() == "добавить ещё услугу":
        await message.answer("Введите название следующей услуги:", reply_markup=ReplyKeyboardRemove())
        await state.set_state(InvoiceStates.услуга)
    else:
        await finalize_invoice(state, message)

@dp.message(F.text.lower() == "сделать акт")
async def handle_create_akt(message: Message, state: FSMContext):
    data = await state.get_data()
    if not data.get('услуги'):
        await message.answer("Нельзя создать акт — счёт ещё не сформирован.")
        return
    await finalize_akt(state, message)

def convert_xlsx_to_pdf(xlsx_path: str, output_dir: str):
    subprocess.run([
        "libreoffice",
        "--headless",
        "--convert-to", "pdf",
        "--outdir", output_dir,
        xlsx_path
    ])

async def finalize_invoice(state: FSMContext, message: Message):
    data = await state.get_data()

    # Выбор шаблона в зависимости от количества услуг
    if len(data['услуги']) == 1:
        template_path = "Рахунок шаблон 1.xlsx"
    elif len(data['услуги']) == 2:
        template_path = "Рахунок шаблон 2.xlsx"
    elif len(data['услуги']) == 3:
        template_path = "Рахунок шаблон 3.xlsx"
    elif len(data['услуги']) == 4:
        template_path = "Рахунок шаблон 4.xlsx"
    elif len(data['услуги']) == 5:
        template_path = "Рахунок шаблон 5.xlsx"
    elif len(data['услуги']) == 6:
        template_path = "Рахунок шаблон 6.xlsx"
    elif len(data['услуги']) == 7:
        template_path = "Рахунок шаблон 7.xlsx"
    elif len(data['услуги']) == 8:
        template_path = "Рахунок шаблон 8.xlsx"
    else:
        template_path = TEMPLATE_PATH  # шаблон по умолчанию

    wb = load_workbook(template_path)
    ws = wb.active


    ws['E11'] = data['номер_счета']
    ws['C8'] = data['фирма']
    ws['D12'] = f"Від {data['дата']} року"

    start_row = 14
    total = 0
    for i, услуга in enumerate(data['услуги']):
        row = start_row + i
        ws[f'B{row}'] = услуга['название'].capitalize()
        ws[f'F{row}'] = услуга['количество']
        ws[f'G{row}'] = услуга['цена']
        ws[f'H{row}'] = услуга['сумма']
        total += услуга['сумма']
    num_total = len(data['услуги']) + 14
    ws[f'H{num_total}'] = total

    def сумма_прописью_укр(amount: float) -> str:
        гривны = int(amount)
        копейки = round((amount - гривны) * 100)

        гривны_txt = num2words(гривны, lang='uk').capitalize()
        гривны_слово = "гривня" if гривны % 10 == 1 and гривны % 100 != 11 else \
            "гривні" if 2 <= гривны % 10 <= 4 and not 12 <= гривны % 100 <= 14 else \
                "гривень"

        копейки_txt = f"{копейки:02d} копійок"

        return f"{гривны_txt} {гривны_слово} {копейки_txt}"

    ws[f'A{num_total + 3}'] = сумма_прописью_укр(total) + ' Без ПДВ'

    await state.update_data(готово_к_акту=True)
    output_filename = f"Счет_{data['номер_счета']}.xlsx"
    output_path = os.path.join("docs", output_filename)
    wb.save(output_path)
    wb.close()

    convert_xlsx_to_pdf(output_path, "docs")
    pdf_path = output_path.replace(".xlsx", ".pdf")
    await message.answer_document(FSInputFile(pdf_path), caption="PDF-версия")

    await state.set_state(None)  # снимаем состояние
    keyboard = ReplyKeyboardMarkup(resize_keyboard=True, keyboard=[
        [KeyboardButton(text="Сделать акт")]
    ])
    await message.answer_document(FSInputFile(output_path), caption="Готовый рахунок", reply_markup=keyboard)

async def finalize_akt(state: FSMContext, message: Message):
    data = await state.get_data()

    # Выбор шаблона в зависимости от количества услуг
    if len(data['услуги']) == 1:
        template_path = "Акт работ шаблон 1.xlsx"
    elif len(data['услуги']) == 2:
        template_path = "Акт работ шаблон 2.xlsx"
    elif len(data['услуги']) == 3:
        template_path = "Рахунок шаблон 3.xlsx"
    elif len(data['услуги']) == 4:
        template_path = "Рахунок шаблон 4.xlsx"
    elif len(data['услуги']) == 5:
        template_path = "Рахунок шаблон 5.xlsx"
    elif len(data['услуги']) == 6:
        template_path = "Рахунок шаблон 6.xlsx"
    elif len(data['услуги']) == 7:
        template_path = "Рахунок шаблон 7.xlsx"
    elif len(data['услуги']) == 8:
        template_path = "Рахунок шаблон 8.xlsx"
    else:
        template_path = TEMPLATE_PATH  # шаблон по умолчанию

    wb = load_workbook(template_path)
    ws = wb.active

    ws['C9'] = data['номер_счета']
    ws['C12'] = data['номер_счета']
    # ws['C8'] = data['фирма']
    ws['E9'] = data['дата']
    ws['E12'] = data['дата']
    ws['A50'] = data['дата']
    ws['D50'] = data['дата']

    start_row = 16
    total = 0
    for i, услуга in enumerate(data['услуги']):
        row = start_row + i
        ws[f'B{row}'] = услуга['название'].capitalize()
        ws[f'C{row}'] = услуга['количество']
        ws[f'E{row}'] = услуга['цена']
        ws[f'F{row}'] = услуга['сумма']
        total += услуга['сумма']
    num_total = len(data['услуги']) + 16
    ws[f'F{num_total}'] = total

    def сумма_прописью_укр(amount: float) -> str:
        гривны = int(amount)
        копейки = round((amount - гривны) * 100)

        гривны_txt = num2words(гривны, lang='uk').capitalize()
        гривны_слово = "гривня" if гривны % 10 == 1 and гривны % 100 != 11 else \
            "гривні" if 2 <= гривны % 10 <= 4 and not 12 <= гривны % 100 <= 14 else \
                "гривень"

        копейки_txt = f"{копейки:02d} копійок"

        return f"{гривны_txt} {гривны_слово} {копейки_txt}"

    ws[f'C{num_total + 2}'] = сумма_прописью_укр(total) + ' Без ПДВ'

    output_filename = f"Акт_{data['номер_счета']}.xlsx"
    output_path = os.path.join("docs", output_filename)
    wb.save(output_path)
    wb.close()

    convert_xlsx_to_pdf(output_path, "docs")
    pdf_path = output_path.replace(".xlsx", ".pdf")
    await message.answer_document(FSInputFile(pdf_path), caption="PDF-версия")

    keyboard = ReplyKeyboardMarkup(resize_keyboard=True, keyboard=[
        [KeyboardButton(text="Создать счёт")]
    ])
    await message.answer_document(FSInputFile(output_path), caption="Акт готовий", reply_markup=keyboard)
    await state.clear()


if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO)
    dp.run_polling(bot)