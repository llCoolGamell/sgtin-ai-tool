from aiogram import Router
from aiogram.filters import CommandStart
from aiogram.types import Message

router = Router()

WELCOME_MESSAGE = """
👋 <b>Добро пожаловать в МаркировкаГид!</b>

Я — AI-помощник по маркировке товаров в системах МДЛП и ГИС МТ (Честный ЗНАК).

<b>Мои возможности:</b>

📋 /table — Таблица сроков вступления маркировки
   (лекарства, БАДы, мед. изделия, косметика, масла, ТСР)

❓ /ask <i>вопрос</i> — Задать вопрос по маркировке
   Примеры:
   • /ask Какой документ подать при возврате БАДа?
   • /ask Какой штраф за продажу без маркировки?
   • /ask Когда начинается поэкземплярный учёт для ТСР?

➕ /add — Добавить информацию в базу знаний (для админов)

📊 /summary — Краткая сводка по всем группам товаров

ℹ️ /help — Показать это сообщение

<i>Или просто напишите свой вопрос — я постараюсь помочь!</i>
"""


@router.message(CommandStart())
async def cmd_start(message: Message) -> None:
    await message.answer(WELCOME_MESSAGE, parse_mode="HTML")


@router.message(lambda msg: msg.text and msg.text.strip().lower() == "/help")
async def cmd_help(message: Message) -> None:
    await message.answer(WELCOME_MESSAGE, parse_mode="HTML")
