from openai import AsyncOpenAI

from bot.config import OPENAI_API_KEY
from bot.services.knowledge_base import get_all_knowledge_texts, search_knowledge

client = AsyncOpenAI(api_key=OPENAI_API_KEY) if OPENAI_API_KEY else None

SYSTEM_PROMPT = """Ты — AI-помощник по маркировке товаров в России (системы МДЛП и ГИС МТ «Честный ЗНАК»).

Твоя специализация:
- Лекарственные препараты (МДЛП)
- БАДы (биологически активные добавки)
- Медицинские изделия
- Косметика и парфюмерия
- Масла (смазочные материалы)
- ТСР (технические средства реабилитации)

Ты помогаешь с:
1. Документооборотом — какие документы подавать в конкретных ситуациях
2. Сроками — когда вступают/вступили этапы маркировки
3. Штрафами — последствия нарушений
4. Техническими вопросами — API, интеграция, коды маркировки
5. Нестандартными ситуациями — возвраты, перемаркировка, расхождения

Правила:
- Отвечай на русском языке
- Давай конкретные ответы со ссылками на нормативные акты где возможно
- Если не уверен — скажи об этом, не выдумывай
- Используй предоставленный контекст из базы знаний
- Будь кратким, но информативным
"""


async def get_ai_response(user_question: str) -> str:
    if not client:
        return (
            "⚠️ AI-ответы недоступны: не настроен OPENAI_API_KEY.\n\n"
            "Попробуйте использовать /table для просмотра сроков маркировки "
            "или обратитесь к администратору бота."
        )

    relevant_docs = search_knowledge(user_question)
    if not relevant_docs:
        relevant_docs = get_all_knowledge_texts()

    context = "\n\n---\n\n".join(relevant_docs[:3])

    messages = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {
            "role": "user",
            "content": (
                f"Контекст из базы знаний:\n\n{context}\n\n"
                f"---\n\nВопрос пользователя: {user_question}"
            ),
        },
    ]

    try:
        response = await client.chat.completions.create(
            model="gpt-4o-mini",
            messages=messages,
            max_tokens=2000,
            temperature=0.3,
        )
        return response.choices[0].message.content or "Не удалось получить ответ."
    except Exception as e:
        return f"❌ Ошибка при обращении к AI: {e}"
