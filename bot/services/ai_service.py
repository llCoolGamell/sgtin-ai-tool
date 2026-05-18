import logging

from bot.config import OPENAI_API_KEY, GEMINI_API_KEY, GROQ_API_KEY
from bot.services.knowledge_base import get_all_knowledge_texts, search_knowledge
from bot.services.web_search import search_marking_sites

logger = logging.getLogger(__name__)

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
- Используй предоставленный контекст из базы знаний и результаты веб-поиска
- Если есть данные из веб-поиска — приоритет у них (они актуальнее)
- Указывай источники информации когда возможно
- Будь кратким, но информативным
"""


async def _build_context(user_question: str) -> str:
    parts = []

    # Web search (priority — most up to date)
    try:
        web_results = await search_marking_sites(user_question)
        if web_results:
            parts.append("=== Результаты веб-поиска (актуальные данные) ===\n" + web_results)
            logger.info("Web search returned results")
    except Exception as e:
        logger.warning(f"Web search failed: {e}")

    # Local knowledge base
    relevant_docs = search_knowledge(user_question)
    if not relevant_docs:
        relevant_docs = get_all_knowledge_texts()
    if relevant_docs:
        parts.append("=== Локальная база знаний ===\n" + "\n\n".join(relevant_docs[:2]))

    return "\n\n---\n\n".join(parts) if parts else ""


async def _try_groq(user_question: str, context: str) -> str:
    from groq import AsyncGroq

    client = AsyncGroq(api_key=GROQ_API_KEY)
    response = await client.chat.completions.create(
        model="llama-3.1-8b-instant",
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {
                "role": "user",
                "content": f"Контекст из базы знаний:\n\n{context}\n\n---\n\nВопрос пользователя: {user_question}",
            },
        ],
        max_tokens=2000,
        temperature=0.3,
    )
    return response.choices[0].message.content or ""


async def _try_gemini(user_question: str, context: str) -> str:
    from google import genai

    client = genai.Client(api_key=GEMINI_API_KEY)
    full_prompt = (
        f"{SYSTEM_PROMPT}\n\n"
        f"Контекст из базы знаний:\n\n{context}\n\n---\n\n"
        f"Вопрос пользователя: {user_question}"
    )
    response = client.models.generate_content(
        model="gemini-2.0-flash",
        contents=full_prompt,
    )
    return response.text or ""


async def _try_openai(user_question: str, context: str) -> str:
    from openai import AsyncOpenAI

    client = AsyncOpenAI(api_key=OPENAI_API_KEY)
    response = await client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {
                "role": "user",
                "content": f"Контекст из базы знаний:\n\n{context}\n\n---\n\nВопрос пользователя: {user_question}",
            },
        ],
        max_tokens=2000,
        temperature=0.3,
    )
    return response.choices[0].message.content or ""


async def get_ai_response(user_question: str) -> str:
    providers = []
    if GROQ_API_KEY:
        providers.append(("Groq (Llama 3.1)", _try_groq))
    if GEMINI_API_KEY:
        providers.append(("Google Gemini", _try_gemini))
    if OPENAI_API_KEY:
        providers.append(("OpenAI", _try_openai))

    if not providers:
        return (
            "⚠️ AI-ответы недоступны: не настроен ни один AI-провайдер.\n\n"
            "Поддерживаются: GROQ_API_KEY, GEMINI_API_KEY, OPENAI_API_KEY\n\n"
            "Попробуйте использовать /table для просмотра сроков маркировки."
        )

    context = await _build_context(user_question)
    errors = []

    for name, provider_fn in providers:
        try:
            logger.info(f"Trying AI provider: {name}")
            result = await provider_fn(user_question, context)
            if result:
                return result
        except Exception as e:
            logger.warning(f"Provider {name} failed: {e}")
            errors.append(f"{name}: {e}")

    error_details = "\n".join(errors)
    return f"❌ Все AI-провайдеры недоступны:\n\n{error_details}"
