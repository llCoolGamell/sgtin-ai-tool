from aiogram import Router
from aiogram.types import Message

from bot.services.ai_service import get_ai_response

router = Router()


@router.message(lambda msg: msg.text and msg.text.strip().lower().startswith("/ask"))
async def cmd_ask(message: Message) -> None:
    text = message.text.strip()
    question = text[4:].strip() if len(text) > 4 else ""

    if not question:
        await message.answer(
            "❓ Пожалуйста, задайте вопрос после команды /ask\n\n"
            "Примеры:\n"
            "• /ask Какой документ подать при возврате БАДа?\n"
            "• /ask Какой штраф за продажу без маркировки?\n"
            "• /ask Когда вступает маркировка косметики?",
            parse_mode="HTML",
        )
        return

    processing_msg = await message.answer("🔄 Обрабатываю ваш вопрос...")

    response = await get_ai_response(question)

    await processing_msg.delete()
    await message.answer(response, parse_mode="HTML")


@router.message(
    lambda msg: (
        msg.text
        and not msg.text.startswith("/")
        and len(msg.text.strip()) > 3
    )
)
async def handle_free_text(message: Message) -> None:
    processing_msg = await message.answer("🔄 Обрабатываю ваш вопрос...")

    response = await get_ai_response(message.text.strip())

    await processing_msg.delete()
    await message.answer(response, parse_mode="HTML")
