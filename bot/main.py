import asyncio
import logging

from aiogram import Bot, Dispatcher
from aiogram.fsm.storage.memory import MemoryStorage

from bot.config import TELEGRAM_BOT_TOKEN
from bot.handlers import start, table, ask, add

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger(__name__)


def create_bot() -> Bot:
    if not TELEGRAM_BOT_TOKEN:
        raise ValueError(
            "TELEGRAM_BOT_TOKEN не установлен. "
            "Скопируйте .env.example в .env и укажите токен бота."
        )
    return Bot(token=TELEGRAM_BOT_TOKEN)


def create_dispatcher() -> Dispatcher:
    dp = Dispatcher(storage=MemoryStorage())
    dp.include_router(start.router)
    dp.include_router(table.router)
    dp.include_router(add.router)
    dp.include_router(ask.router)
    return dp


async def main() -> None:
    logger.info("Запуск бота МаркировкаГид...")
    bot = create_bot()
    dp = create_dispatcher()

    await bot.delete_webhook(drop_pending_updates=True)
    logger.info("Бот запущен и готов к работе!")
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
