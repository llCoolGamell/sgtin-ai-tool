import os
import json
from datetime import datetime, date
from typing import Optional

from bot.config import KNOWLEDGE_DIR, DEADLINES_FILE


def load_deadlines() -> dict:
    with open(DEADLINES_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def get_all_knowledge_texts() -> list[str]:
    texts = []
    for filename in os.listdir(KNOWLEDGE_DIR):
        if filename.endswith(".md"):
            filepath = os.path.join(KNOWLEDGE_DIR, filename)
            with open(filepath, "r", encoding="utf-8") as f:
                texts.append(f.read())
    return texts


def format_deadlines_table(category_filter: Optional[str] = None) -> str:
    data = load_deadlines()
    lines = []
    lines.append("📋 <b>Таблица сроков вступления маркировки</b>")
    lines.append(f"<i>Обновлено: {data['last_updated']}</i>")
    lines.append("")

    today = date.today()

    for cat in data["categories"]:
        if category_filter and category_filter.lower() not in cat["name"].lower():
            continue

        lines.append(f"━━━━━━━━━━━━━━━━━━━━")
        lines.append(f"<b>📦 {cat['name']}</b>")
        lines.append(f"<i>Система: {cat['system']}</i>")
        lines.append("")

        for stage in cat["stages"]:
            stage_date = datetime.strptime(stage["date"], "%Y-%m-%d").date()
            if stage["status"] == "действует":
                icon = "✅"
            elif stage["status"] == "завершено":
                icon = "☑️"
            elif stage_date <= today:
                icon = "✅"
            else:
                icon = "⏳"

            lines.append(f"  {icon} {stage['stage']}")
            lines.append(f"     📅 {stage['date']} | {stage['status']}")
            lines.append("")

    lines.append("━━━━━━━━━━━━━━━━━━━━")
    lines.append(f"<i>Источник: {data['source']}</i>")

    return "\n".join(lines)


def format_deadlines_short() -> str:
    data = load_deadlines()
    lines = []
    lines.append("📋 <b>Краткая таблица сроков маркировки</b>")
    lines.append("")

    for cat in data["categories"]:
        upcoming = [
            s for s in cat["stages"]
            if s["status"] == "предстоит"
        ]
        active_count = sum(1 for s in cat["stages"] if s["status"] == "действует")

        lines.append(f"<b>{cat['name']}</b> ({cat['system']})")
        lines.append(f"  Активных этапов: {active_count}")
        if upcoming:
            next_stage = upcoming[0]
            lines.append(f"  ⏳ Ближайший: {next_stage['stage']} — {next_stage['date']}")
        else:
            lines.append("  ✅ Все этапы введены")
        lines.append("")

    return "\n".join(lines)


def add_knowledge_entry(title: str, content: str) -> str:
    filename = f"{title.replace(' ', '_').lower()}.md"
    filepath = os.path.join(KNOWLEDGE_DIR, filename)

    with open(filepath, "w", encoding="utf-8") as f:
        f.write(f"# {title}\n\n{content}\n")

    return filepath


def search_knowledge(query: str) -> list[str]:
    query_lower = query.lower()
    results = []

    for filename in os.listdir(KNOWLEDGE_DIR):
        if filename.endswith(".md"):
            filepath = os.path.join(KNOWLEDGE_DIR, filename)
            with open(filepath, "r", encoding="utf-8") as f:
                content = f.read()
            if query_lower in content.lower():
                results.append(content)

    return results
