from aiogram import Router
from aiogram.types import Message
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup

from bot.config import ADMIN_IDS
from bot.services.knowledge_base import add_knowledge_entry

router = Router()


class AddKnowledge(StatesGroup):
    waiting_for_title = State()
    waiting_for_content = State()


@router.message(lambda msg: msg.text and msg.text.strip().lower() == "/add")
async def cmd_add(message: Message, state: FSMContext) -> None:
    if message.from_user.id not in ADMIN_IDS:
        await message.answer(
            "⛔ Эта команда доступна только администраторам бота.\n"
            "Обратитесь к администратору для добавления информации."
        )
        return

    await message.answer(
        "📝 <b>Добавление в базу знаний</b>\n\n"
        "Введите заголовок новой записи:\n"
        "(или /cancel для отмены)",
        parse_mode="HTML",
    )
    await state.set_state(AddKnowledge.waiting_for_title)


@router.message(AddKnowledge.waiting_for_title)
async def process_title(message: Message, state: FSMContext) -> None:
    if message.text and message.text.strip().lower() == "/cancel":
        await state.clear()
        await message.answer("❌ Отменено.")
        return

    await state.update_data(title=message.text.strip())
    await message.answer(
        "📄 Теперь введите содержание записи:\n"
        "(можно использовать несколько сообщений, отправьте /done когда закончите)\n"
        "(или /cancel для отмены)",
        parse_mode="HTML",
    )
    await state.set_state(AddKnowledge.waiting_for_content)


@router.message(AddKnowledge.waiting_for_content)
async def process_content(message: Message, state: FSMContext) -> None:
    if message.text and message.text.strip().lower() == "/cancel":
        await state.clear()
        await message.answer("❌ Отменено.")
        return

    if message.text and message.text.strip().lower() == "/done":
        data = await state.get_data()
        title = data.get("title", "")
        content = data.get("content", "")

        if not content:
            await message.answer("⚠️ Содержание пустое. Введите текст или /cancel")
            return

        filepath = add_knowledge_entry(title, content)
        await state.clear()
        await message.answer(
            f"✅ Запись добавлена в базу знаний!\n\n"
            f"<b>Заголовок:</b> {title}\n"
            f"<b>Файл:</b> {filepath}",
            parse_mode="HTML",
        )
        return

    data = await state.get_data()
    existing_content = data.get("content", "")
    new_content = existing_content + "\n" + message.text.strip() if existing_content else message.text.strip()
    await state.update_data(content=new_content)
    await message.answer("✏️ Принято. Продолжайте ввод или отправьте /done")
