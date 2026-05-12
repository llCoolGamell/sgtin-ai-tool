import os
from dotenv import load_dotenv

load_dotenv()

TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
ADMIN_IDS = [
    int(x.strip())
    for x in os.getenv("ADMIN_IDS", "").split(",")
    if x.strip().isdigit()
]

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
KNOWLEDGE_DIR = os.path.join(DATA_DIR, "knowledge")
DEADLINES_FILE = os.path.join(DATA_DIR, "deadlines.json")
