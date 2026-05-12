from aiogram import Router
from aiogram.types import Message, CallbackQuery, InlineKeyboardMarkup, InlineKeyboardButton

from bot.services.knowledge_base import format_deadlines_table, format_deadlines_short

router = Router()

CATEGORY_KEYBOARD = InlineKeyboardMarkup(inline_keyboard=[
    [
        InlineKeyboardButton(text="💊 Лекарства", callback_data="table_medicines"),
        InlineKeyboardButton(text="🧬 БАДы", callback_data="table_bads"),
    ],
    [
        InlineKeyboardButton(text="🩺 Мед. изделия", callback_data="table_medical"),
        InlineKeyboardButton(text="💄 Косметика", callback_data="table_cosmetics"),
    ],
    [
        InlineKeyboardButton(text="🛢 Масла", callback_data="table_oils"),
        InlineKeyboardButton(text="♿ ТСР", callback_data="table_tsr"),
    ],
    [
        InlineKeyboardButton(text="📋 Все группы", callback_data="table_all"),
    ],
])


@router.message(lambda msg: msg.text and msg.text.strip().lower() == "/table")
async def cmd_table(message: Message) -> None:
    summary = format_deadlines_short()
    await message.answer(
        f"{summary}\n\nВыберите группу для подробной информации:",
        parse_mode="HTML",
        reply_markup=CATEGORY_KEYBOARD,
    )


@router.message(lambda msg: msg.text and msg.text.strip().lower() == "/summary")
async def cmd_summary(message: Message) -> None:
    summary = format_deadlines_short()
    await message.answer(summary, parse_mode="HTML")


CATEGORY_FILTERS = {
    "table_medicines": "Лекарственные",
    "table_bads": "БАД",
    "table_medical": "Медицинские",
    "table_cosmetics": "косметик",
    "table_oils": "Масла",
    "table_tsr": "ТСР",
    "table_all": None,
}


@router.callback_query(lambda c: c.data and c.data.startswith("table_"))
async def callback_table(callback: CallbackQuery) -> None:
    category_filter = CATEGORY_FILTERS.get(callback.data)
    table = format_deadlines_table(category_filter)

    if len(table) > 4096:
        parts = split_message(table, 4096)
        for part in parts:
            await callback.message.answer(part, parse_mode="HTML")
    else:
        await callback.message.answer(table, parse_mode="HTML")

    await callback.answer()


def split_message(text: str, max_length: int) -> list[str]:
    parts = []
    while len(text) > max_length:
        split_pos = text.rfind("\n", 0, max_length)
        if split_pos == -1:
            split_pos = max_length
        parts.append(text[:split_pos])
        text = text[split_pos:]
    if text:
        parts.append(text)
    return parts
